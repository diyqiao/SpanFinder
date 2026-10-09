using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading.Tasks;
using Windows.ApplicationModel.DataTransfer;

namespace Span.Helpers;

/// <summary>
/// WeChat 4.x (and similar apps) put delayed-render CF_HDROP data on the OLE
/// data object during drag-and-drop. WinUI3's DataPackageView fails to map it
/// to StorageItems, so the drop arrives with no usable format (Issue #73).
///
/// This helper provides:
/// 1. HasDroppableData: Accepts OLE formats (FileDrop, FileNameW, etc.) during DragOver.
/// 2. ExtractPathsAsync: Multi-tier fallback extracting real file paths via OLE,
///    WinRT Custom Format, and WeChat RWTemp immediate decrypted file detection.
/// </summary>
internal static class Win32DragDropHelper
{
    private const int S_OK = 0;
    private const short CF_HDROP = 15;

    private static string? _cachedWeChatDataDir = null;

    /// <summary>
    /// Determines whether the DragEventArgs contains droppable content,
    /// supporting internal Span drags, standard StorageItems/Text, and
    /// external OLE formats like FileDrop, FileNameW, FileName from WeChat 4.x.
    /// </summary>
    public static bool HasDroppableData(DataPackageView? dataView)
    {
        if (dataView == null) return false;
        try
        {
            if (dataView.Properties.ContainsKey("SourcePaths") ||
                dataView.Contains(StandardDataFormats.StorageItems) ||
                dataView.Contains(StandardDataFormats.Text))
                return true;

            var formats = dataView.AvailableFormats;
            if (formats != null && formats.Count > 0)
            {
                foreach (var f in formats)
                {
                    if (f.IndexOf("File", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        f.IndexOf("Drop", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        f.IndexOf("URL", StringComparison.OrdinalIgnoreCase) >= 0)
                        return true;
                }
            }
        }
        catch { }
        return false;
    }

    /// <summary>
    /// Extracts drop paths from external apps using OLE CF_HDROP, WinRT custom formats,
    /// and WeChat RWTemp immediate file detection.
    /// </summary>
    public static async Task<List<string>> ExtractExternalPathsAsync(DataPackageView dataView, IntPtr hwndDropTarget)
    {
        var result = new List<string>();
        try
        {
            var fmts = dataView.AvailableFormats;
            DebugLogger.Log($"[Win32DragDrop] ExtractExternalPathsAsync start, formats: {(fmts != null ? string.Join(", ", fmts) : "null")}");
        }
        catch { }

        // 1. Try live OLE IDataObject CF_HDROP
        try
        {
            var livePaths = GetLiveDragFileDrop(hwndDropTarget);
            if (livePaths != null && livePaths.Count > 0)
            {
                DebugLogger.Log($"[Win32DragDrop] Live OLE CF_HDROP found {livePaths.Count} items");
                return livePaths;
            }
        }
        catch (Exception ex)
        {
            DebugLogger.Log($"[Win32DragDrop] GetLiveDragFileDrop error: {ex.Message}");
        }

        // 2. Try DataPackageView custom formats: "FileNameW", "FileName", "FileDrop"
        try
        {
            var formats = dataView.AvailableFormats;
            if (formats != null)
            {
                foreach (var fmt in new[] { "FileNameW", "FileName", "FileDrop" })
                {
                    if (formats.Contains(fmt))
                    {
                        try
                        {
                            var data = await dataView.GetDataAsync(fmt);
                            DebugLogger.Log($"[Win32DragDrop] GetDataAsync('{fmt}') returned: {data?.GetType().FullName ?? "null"}");
                            if (data is string s && !string.IsNullOrWhiteSpace(s) && (File.Exists(s) || Directory.Exists(s)))
                            {
                                result.Add(s);
                                return result;
                            }
                            if (data is IEnumerable<string> strList)
                            {
                                var existing = strList.Where(p => File.Exists(p) || Directory.Exists(p)).ToList();
                                if (existing.Count > 0) return existing;
                            }
                            if (data is Windows.Storage.Streams.IRandomAccessStream stream)
                            {
                                using var reader = new Windows.Storage.Streams.DataReader(stream.GetInputStreamAt(0));
                                await reader.LoadAsync((uint)stream.Size);
                                var bytes = new byte[stream.Size];
                                reader.ReadBytes(bytes);
                                var pathStr = ParseNullTerminatedUnicode(bytes);
                                if (!string.IsNullOrWhiteSpace(pathStr) && (File.Exists(pathStr) || Directory.Exists(pathStr)))
                                {
                                    result.Add(pathStr);
                                    return result;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            DebugLogger.Log($"[Win32DragDrop] GetDataAsync('{fmt}') error: {ex.Message}");
                        }
                    }
                }
            }
        }
        catch { }

        // 3. Fallback: WeChat 4.x on-the-fly decrypted temp image detection
        // WeChat 4.x decrypts chat images into temp\RWTemp immediately upon drag start.
        try
        {
            var formats = dataView.AvailableFormats;
            bool isWeChatLike = formats != null && formats.Any(f =>
                f.IndexOf("File", StringComparison.OrdinalIgnoreCase) >= 0 ||
                f.IndexOf("Drop", StringComparison.OrdinalIgnoreCase) >= 0 ||
                f.IndexOf("URL", StringComparison.OrdinalIgnoreCase) >= 0);

            if (isWeChatLike)
            {
                var wechatFiles = GetRecentWeChatDropFiles(30);
                if (wechatFiles.Count > 0)
                {
                    DebugLogger.Log($"[Win32DragDrop] WeChat RWTemp detected {wechatFiles.Count} file(s): {string.Join("; ", wechatFiles)}");
                    return wechatFiles;
                }
            }
        }
        catch (Exception ex)
        {
            DebugLogger.Log($"[Win32DragDrop] WeChat RWTemp fallback error: {ex.Message}");
        }

        return result;
    }

    /// <summary>
    /// Reads CF_HDROP paths from the OLE data object of the drag currently
    /// in progress over <paramref name="hwndDropTarget"/>.
    /// </summary>
    public static List<string>? GetLiveDragFileDrop(IntPtr hwndDropTarget)
    {
        try
        {
            if (hwndDropTarget == IntPtr.Zero) return null;

            IntPtr pUnk = GetProp(hwndDropTarget, "OleDropTargetInterface");
            if (pUnk == IntPtr.Zero) return null;

            IDataObject? dataObj = null;
            try
            {
                dataObj = (IDataObject)Marshal.GetObjectForIUnknown(pUnk);
            }
            catch { return null; }
            if (dataObj == null) return null;

            try
            {
                var fmt = MakeFormat(CF_HDROP);
                if (dataObj.QueryGetData(ref fmt) != S_OK) return null;

                dataObj.GetData(ref fmt, out STGMEDIUM medium);
                try
                {
                    if (medium.unionmember == IntPtr.Zero) return null;
                    var paths = ReadPathsFromHDrop(medium.unionmember);
                    return paths.Count > 0 ? paths : null;
                }
                finally { NativeMethods.ReleaseStgMedium(ref medium); }
            }
            finally
            {
                Marshal.ReleaseComObject(dataObj);
            }
        }
        catch (Exception ex)
        {
            DebugLogger.Log($"[Win32DragDrop] GetLiveDragFileDrop exception: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Finds recent files created in WeChat's RWTemp directory within the last <paramref name="maxSeconds"/>.
    /// </summary>
    public static List<string> GetRecentWeChatDropFiles(int maxSeconds = 30)
    {
        var result = new List<string>();
        try
        {
            var rwTempDirs = FindWeChatRWTempDirs();
            var nowUtc = DateTime.UtcNow;
            var thresholdUtc = nowUtc.AddSeconds(-maxSeconds);
            var nowLocal = DateTime.Now;
            var thresholdLocal = nowLocal.AddSeconds(-maxSeconds);

            var recentFiles = new List<FileInfo>();
            foreach (var rwDir in rwTempDirs)
            {
                if (!Directory.Exists(rwDir)) continue;
                var dirInfo = new DirectoryInfo(rwDir);
                try
                {
                    var files = dirInfo.EnumerateFiles("*.*", SearchOption.AllDirectories)
                        .Where(f => f.LastWriteTimeUtc >= thresholdUtc || f.LastWriteTime >= thresholdLocal ||
                                    f.CreationTimeUtc >= thresholdUtc || f.CreationTime >= thresholdLocal)
                        .OrderByDescending(f => f.LastWriteTimeUtc)
                        .Take(10);
                    recentFiles.AddRange(files);
                }
                catch { }
            }

            if (recentFiles.Count > 0)
            {
                // Group by directory (a single drag creates files in one unique GUID folder)
                var topGroup = recentFiles
                    .GroupBy(f => f.DirectoryName)
                    .OrderByDescending(g => g.Max(f => f.LastWriteTimeUtc))
                    .FirstOrDefault();

                if (topGroup != null)
                {
                    foreach (var f in topGroup)
                    {
                        result.Add(f.FullName);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            DebugLogger.Log($"[Win32DragDrop] GetRecentWeChatDropFiles exception: {ex.Message}");
        }
        return result;
    }

    /// <summary>
    /// Enumerates every file currently present under WeChat's RWTemp directory.
    /// Used by the native drop target to snapshot the directory at DragEnter and
    /// then diff against it at Drop/poll time to detect freshly-decrypted images.
    /// </summary>
    public static List<string> EnumerateWeChatRwTempFiles()
    {
        var result = new List<string>();
        try
        {
            foreach (var rwDir in FindWeChatRWTempDirs())
            {
                if (!Directory.Exists(rwDir)) continue;
                try { result.AddRange(Directory.EnumerateFiles(rwDir, "*.*", SearchOption.AllDirectories)); }
                catch { }
            }
        }
        catch (Exception ex)
        {
            DebugLogger.Log($"[Win32DragDrop] EnumerateWeChatRwTempFiles error: {ex.Message}");
        }
        return result;
    }

    private static List<string> FindWeChatRWTempDirs()
    {
        var list = new List<string>();

        if (!string.IsNullOrEmpty(_cachedWeChatDataDir) && Directory.Exists(_cachedWeChatDataDir))
        {
            ScanDataDirForRWTemp(_cachedWeChatDataDir, list);
            if (list.Count > 0) return list;
        }

        var candidates = new List<string>();

        try
        {
            foreach (var drive in DriveInfo.GetDrives().Where(d => d.IsReady))
            {
                candidates.Add(Path.Combine(drive.RootDirectory.FullName, "微信", "xwechat_files"));
                candidates.Add(Path.Combine(drive.RootDirectory.FullName, "xwechat_files"));
                candidates.Add(Path.Combine(drive.RootDirectory.FullName, "WeChat Files"));
            }
        }
        catch { }

        var docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        if (!string.IsNullOrEmpty(docs))
        {
            candidates.Add(Path.Combine(docs, "微信", "xwechat_files"));
            candidates.Add(Path.Combine(docs, "xwechat_files"));
            candidates.Add(Path.Combine(docs, "WeChat Files"));
        }

        foreach (var cand in candidates)
        {
            if (Directory.Exists(cand))
            {
                _cachedWeChatDataDir = cand;
                ScanDataDirForRWTemp(cand, list);
                if (list.Count > 0) return list;
            }
        }

        return list;
    }

    private static void ScanDataDirForRWTemp(string dataDir, List<string> list)
    {
        try
        {
            foreach (var userDir in Directory.GetDirectories(dataDir))
            {
                var rw = Path.Combine(userDir, "temp", "RWTemp");
                if (Directory.Exists(rw))
                {
                    list.Add(rw);
                }
            }
        }
        catch { }
    }

    private static string? ParseNullTerminatedUnicode(byte[] bytes)
    {
        if (bytes == null || bytes.Length < 2) return null;
        int nullIdx = -1;
        for (int i = 0; i < bytes.Length - 1; i += 2)
        {
            if (bytes[i] == 0 && bytes[i + 1] == 0)
            {
                nullIdx = i;
                break;
            }
        }
        if (nullIdx <= 0) nullIdx = bytes.Length;
        var str = System.Text.Encoding.Unicode.GetString(bytes, 0, nullIdx).Trim();
        return string.IsNullOrEmpty(str) ? null : str;
    }

    private static List<string> ReadPathsFromHDrop(IntPtr hDrop)
    {
        var result = new List<string>();
        uint count = DragQueryFileW(hDrop, 0xFFFFFFFF, null, 0);
        for (uint i = 0; i < count; i++)
        {
            uint len = DragQueryFileW(hDrop, i, null, 0);
            if (len == 0) continue;
            var sb = new System.Text.StringBuilder((int)len + 1);
            if (DragQueryFileW(hDrop, i, sb, (uint)sb.Capacity) > 0)
            {
                var path = sb.ToString();
                if (!string.IsNullOrWhiteSpace(path)) result.Add(path);
            }
        }
        return result;
    }

    private static FORMATETC MakeFormat(short cfFormat) => new()
    {
        cfFormat = cfFormat,
        dwAspect = DVASPECT.DVASPECT_CONTENT,
        lindex = -1,
        ptd = IntPtr.Zero,
        tymed = TYMED.TYMED_HGLOBAL
    };

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr GetProp(IntPtr hWnd, string lpString);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode, EntryPoint = "DragQueryFileW")]
    private static extern uint DragQueryFileW(IntPtr hDrop, uint iFile, System.Text.StringBuilder? lpszFile, uint cch);
}
