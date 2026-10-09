using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace Span.Helpers;

/// <summary>
/// WeChat 4.x (and a few other apps) put delayed-render CF_HDROP data on the OLE
/// data object during drag-and-drop. WinUI 3's XAML drop pipeline never maps that
/// object to StorageItems, so <c>DragOver</c>/<c>Drop</c> never fire — the user
/// sees the 🚫 cursor and the drop is silently rejected (Issue #73).
///
/// WinUI 3 does NOT register a classic Win32 <c>RegisterDragDrop</c> target on any
/// of its windows (verified by enumerating window properties), which is why a plain
/// Win32 window receives the drag but SPAN does not. This class registers a classic
/// OLE <c>IDropTarget</c> directly on the top-level HWND, reads CF_HDROP itself and
/// hands the real file paths back to the app. This is the same strategy Directory
/// Opus uses to accept WeChat's virtual-file drags.
///
/// The COM object is implemented as a hand-built vtable (no built-in COM interop /
/// CCW) so it works even though <c>WinRT.ComWrappersSupport</c> owns marshalling.
/// </summary>
internal sealed unsafe class Win32DropTarget : IDisposable
{
    private const int S_OK = 0;
    private const short CF_HDROP = 15;
    private const uint DROPEFFECT_COPY = 1;
    private const uint DATADIR_GET = 1;

    private static readonly Guid IID_IUnknown = new("00000000-0000-0000-C000-000000000046");
    private static readonly Guid IID_IDropTarget = new("00000122-0000-0000-C000-000000000046");

    private readonly IntPtr _hwnd;
    public delegate void FilesDroppedHandler(List<string> paths, int screenX, int screenY);
    private readonly FilesDroppedHandler _onFilesDropped;

    private IntPtr _pVtbl;
    private IntPtr _pObj;
    private GCHandle _selfHandle;
    private bool _disposed;
    private int _refCount = 1;

    // WeChat decrypts the dragged image into RWTemp lazily. Snapshot the directory
    // at DragEnter so a Drop that arrives before decryption completes can still find
    // the file by diffing (and, if needed, polling) against this baseline.
    private HashSet<string> _rwTempSnapshot = new(StringComparer.OrdinalIgnoreCase);

    // ── Registration ────────────────────────────────────────────────────────────

    public static Win32DropTarget Register(IntPtr hwnd, FilesDroppedHandler onFilesDropped)
    {
        if (hwnd == IntPtr.Zero)
            throw new ArgumentException("hwnd is zero", nameof(hwnd));

        // RegisterDragDrop requires OLE to be initialised on the calling thread.
        // OleInitialize is idempotent (S_OK/S_FALSE are both success).
        int oleInit = OleInitialize(IntPtr.Zero);
        DebugLogger.Log($"[Win32DropTarget] OleInitialize hr=0x{oleInit:X8}");

        var target = new Win32DropTarget(hwnd, onFilesDropped);
        int hr = RegisterDragDrop(hwnd, target._pObj);
        DebugLogger.Log($"[Win32DropTarget] RegisterDragDrop hr=0x{hr:X8}");
        if (hr != S_OK)
        {
            target.Dispose();
            throw new InvalidOperationException($"RegisterDragDrop failed: 0x{hr:X8}");
        }
        return target;
    }

    private Win32DropTarget(IntPtr hwnd, FilesDroppedHandler onFilesDropped)
    {
        _hwnd = hwnd;
        _onFilesDropped = onFilesDropped;

        // IDropTarget vtable: IUnknown (QueryInterface, AddRef, Release)
        //                    + DragEnter, DragOver, DragLeave, Drop
        _pVtbl = Marshal.AllocHGlobal(7 * IntPtr.Size);
        *(IntPtr*)_pVtbl = (IntPtr)(delegate* unmanaged[Stdcall]<IntPtr, Guid*, IntPtr*, int>)&QueryInterface;
        *((IntPtr*)_pVtbl + 1) = (IntPtr)(delegate* unmanaged[Stdcall]<IntPtr, uint>)&AddRef;
        *((IntPtr*)_pVtbl + 2) = (IntPtr)(delegate* unmanaged[Stdcall]<IntPtr, uint>)&Release;
        *((IntPtr*)_pVtbl + 3) = (IntPtr)(delegate* unmanaged[Stdcall]<IntPtr, IntPtr, uint, POINTL, uint*, int>)&DragEnter;
        *((IntPtr*)_pVtbl + 4) = (IntPtr)(delegate* unmanaged[Stdcall]<IntPtr, uint, POINTL, uint*, int>)&DragOver;
        *((IntPtr*)_pVtbl + 5) = (IntPtr)(delegate* unmanaged[Stdcall]<IntPtr, int>)&DragLeave;
        *((IntPtr*)_pVtbl + 6) = (IntPtr)(delegate* unmanaged[Stdcall]<IntPtr, IntPtr, uint, POINTL, uint*, int>)&Drop;

        // Object layout: [vtable pointer][GCHandle] — the GCHandle lets the vtable
        // methods recover this managed instance from the raw "this" pointer.
        _pObj = Marshal.AllocHGlobal(2 * IntPtr.Size);
        *(IntPtr*)_pObj = _pVtbl;
        _selfHandle = GCHandle.Alloc(this);
        *(IntPtr*)(_pObj + IntPtr.Size) = GCHandle.ToIntPtr(_selfHandle);
    }

    private static Win32DropTarget? GetThis(IntPtr pThis)
    {
        if (pThis == IntPtr.Zero) return null;
        try
        {
            IntPtr gc = *(IntPtr*)(pThis + IntPtr.Size);
            if (gc == IntPtr.Zero) return null;
            return GCHandle.FromIntPtr(gc).Target as Win32DropTarget;
        }
        catch
        {
            return null;
        }
    }

    // ── IUnknown ────────────────────────────────────────────────────────────────

    [UnmanagedCallersOnly(CallConvs = new[] { typeof(CallConvStdcall) })]
    private static int QueryInterface(IntPtr pThis, Guid* riid, IntPtr* ppv)
    {
        try
        {
            if (ppv == null || riid == null) return unchecked((int)0x80004003); // E_POINTER
            *ppv = IntPtr.Zero;
            if (*riid == IID_IUnknown || *riid == IID_IDropTarget)
            {
                *ppv = pThis;
                var t = GetThis(pThis);
                if (t != null) Interlocked.Increment(ref t._refCount);
                return S_OK;
            }
            return unchecked((int)0x80004002); // E_NOINTERFACE
        }
        catch
        {
            return unchecked((int)0x80004005); // E_FAIL
        }
    }

    [UnmanagedCallersOnly(CallConvs = new[] { typeof(CallConvStdcall) })]
    private static uint AddRef(IntPtr pThis)
    {
        try
        {
            var t = GetThis(pThis);
            if (t == null) return 1;
            return (uint)Interlocked.Increment(ref t._refCount);
        }
        catch
        {
            return 1;
        }
    }

    [UnmanagedCallersOnly(CallConvs = new[] { typeof(CallConvStdcall) })]
    private static uint Release(IntPtr pThis)
    {
        try
        {
            var t = GetThis(pThis);
            if (t == null) return 0;
            return (uint)t.InternalRelease();
        }
        catch
        {
            return 0;
        }
    }

    private int InternalRelease()
    {
        int n = Interlocked.Decrement(ref _refCount);
        if (n <= 0)
        {
            FreeNativeResources();
        }
        return Math.Max(0, n);
    }

    private void FreeNativeResources()
    {
        try
        {
            if (_pObj != IntPtr.Zero)
            {
                *(IntPtr*)(_pObj + IntPtr.Size) = IntPtr.Zero;
                Marshal.FreeHGlobal(_pObj);
                _pObj = IntPtr.Zero;
            }
            if (_pVtbl != IntPtr.Zero)
            {
                Marshal.FreeHGlobal(_pVtbl);
                _pVtbl = IntPtr.Zero;
            }
            if (_selfHandle.IsAllocated)
            {
                _selfHandle.Free();
            }
        }
        catch { }
    }

    // ── IDropTarget ─────────────────────────────────────────────────────────────

    [UnmanagedCallersOnly(CallConvs = new[] { typeof(CallConvStdcall) })]
    private static int DragEnter(IntPtr pThis, IntPtr pDataObj, uint grfKeyState, POINTL pt, uint* pdwEffect)
    {
        try
        {
            var t = GetThis(pThis);
            if (t != null && pdwEffect != null)
                return t.OnDragEnter(pDataObj, grfKeyState, pt, pdwEffect);
        }
        catch (Exception ex)
        {
            DebugLogger.Log($"[Win32DropTarget] DragEnter exception: {ex.Message}");
        }
        if (pdwEffect != null) *pdwEffect = 0;
        return S_OK;
    }

    [UnmanagedCallersOnly(CallConvs = new[] { typeof(CallConvStdcall) })]
    private static int DragOver(IntPtr pThis, uint grfKeyState, POINTL pt, uint* pdwEffect)
    {
        try
        {
            var t = GetThis(pThis);
            if (t != null && pdwEffect != null)
                return t.OnDragOver(grfKeyState, pt, pdwEffect);
        }
        catch (Exception ex)
        {
            DebugLogger.Log($"[Win32DropTarget] DragOver exception: {ex.Message}");
        }
        if (pdwEffect != null) *pdwEffect = 0;
        return S_OK;
    }

    [UnmanagedCallersOnly(CallConvs = new[] { typeof(CallConvStdcall) })]
    private static int DragLeave(IntPtr pThis)
    {
        try
        {
            var t = GetThis(pThis);
            t?.OnDragLeave();
        }
        catch (Exception ex)
        {
            DebugLogger.Log($"[Win32DropTarget] DragLeave exception: {ex.Message}");
        }
        return S_OK;
    }

    [UnmanagedCallersOnly(CallConvs = new[] { typeof(CallConvStdcall) })]
    private static int Drop(IntPtr pThis, IntPtr pDataObj, uint grfKeyState, POINTL pt, uint* pdwEffect)
    {
        try
        {
            var t = GetThis(pThis);
            if (t != null && pdwEffect != null)
                return t.OnDrop(pDataObj, grfKeyState, pt, pdwEffect);
        }
        catch (Exception ex)
        {
            DebugLogger.Log($"[Win32DropTarget] Drop exception: {ex.Message}");
        }
        if (pdwEffect != null) *pdwEffect = 0;
        return S_OK;
    }

    // ── Managed handlers (run on the UI thread via OLE marshalling) ─────────────

    private int OnDragEnter(IntPtr pDataObj, uint grfKeyState, POINTL pt, uint* pdwEffect)
    {
        // Baseline for detecting WeChat files materialised during the drag.
        try { _rwTempSnapshot = new HashSet<string>(Win32DragDropHelper.EnumerateWeChatRwTempFiles(), StringComparer.OrdinalIgnoreCase); }
        catch { _rwTempSnapshot = new HashSet<string>(StringComparer.OrdinalIgnoreCase); }

        bool accept = ShouldAccept(pDataObj);
        DebugLogger.Log($"[Win32DropTarget] DragEnter ({pt.X},{pt.Y}) accept={accept}");
        *pdwEffect = accept ? DROPEFFECT_COPY : 0;
        return S_OK;
    }

    private int OnDragOver(uint grfKeyState, POINTL pt, uint* pdwEffect)
    {
        // DragEnter already accepted; keep the copy cursor.
        *pdwEffect = DROPEFFECT_COPY;
        return S_OK;
    }

    private void OnDragLeave()
    {
        // Nothing to clean up.
    }

    private int OnDrop(IntPtr pDataObj, uint grfKeyState, POINTL pt, uint* pdwEffect)
    {
        // 1) Preferred path: CF_HDROP already carries the real (decrypted) file.
        var paths = ReadCfHdrop(pDataObj);
        if (paths.Count > 0)
        {
            DebugLogger.Log($"[Win32DropTarget] Drop: CF_HDROP → {paths.Count} file(s)");
            *pdwEffect = DROPEFFECT_COPY;
            SafeInvoke(paths, pt.X, pt.Y);
            return S_OK;
        }

        // 2) WeChat decrypted lazily during the drag — diff RWTemp against the DragEnter snapshot.
        var diff = DiffRwTemp();
        if (diff.Count > 0)
        {
            DebugLogger.Log($"[Win32DropTarget] Drop: RWTemp diff → {diff.Count} file(s)");
            *pdwEffect = DROPEFFECT_COPY;
            SafeInvoke(diff, pt.X, pt.Y);
            return S_OK;
        }

        // 3) Fresh image: only poll RWTemp if WeChat process is active
        bool isWeChatRunning = false;
        try { isWeChatRunning = System.Diagnostics.Process.GetProcessesByName("Weixin").Length > 0; } catch { }
        if (!isWeChatRunning)
        {
            DebugLogger.Log("[Win32DropTarget] Drop: no files found, WeChat not running - skip polling");
            *pdwEffect = 0;
            return S_OK;
        }

        DebugLogger.Log("[Win32DropTarget] Drop: no file yet, polling RWTemp");
        *pdwEffect = DROPEFFECT_COPY;
        var snapshot = _rwTempSnapshot;
        int dropX = pt.X;
        int dropY = pt.Y;
        _ = System.Threading.Tasks.Task.Run(() =>
        {
            for (int i = 0; i < 20; i++)
            {
                System.Threading.Thread.Sleep(300);
                try
                {
                    var current = Win32DragDropHelper.EnumerateWeChatRwTempFiles();
                    var fresh = current.Where(f => !snapshot.Contains(f) && File.Exists(f)).ToList();
                    if (fresh.Count > 0)
                    {
                        var group = GroupNewest(fresh);
                        if (group.Count > 0)
                        {
                            DebugLogger.Log($"[Win32DropTarget] RWTemp poll → {group.Count} file(s)");
                            SafeInvoke(group, dropX, dropY);
                        }
                        return;
                    }
                }
                catch { }
            }
            DebugLogger.Log("[Win32DropTarget] RWTemp poll timed out (no file)");
        });
        return S_OK;
    }

    // ── Format probing & path extraction ────────────────────────────────────────

    /// <summary>
    /// Accept file drags (CF_HDROP / FileNameW / FileGroupDescriptorW / FileContents)
    /// and the WeChat delayed-render case where the data object advertises no formats
    /// at DragEnter. Reject text-only drags so WinUI's own text handling keeps working.
    /// </summary>
    private bool ShouldAccept(IntPtr pDataObj)
    {
        if (pDataObj == IntPtr.Zero) return false;

        if (HasFormat(pDataObj, CF_HDROP)) return true;

        foreach (var name in new[] { "FileNameW", "FileName", "FileGroupDescriptorW", "FileContents" })
        {
            ushort cf = NativeMethods.RegisterClipboardFormatW(name);
            if (cf != 0 && HasFormat(pDataObj, unchecked((short)cf))) return true;
        }

        // Delayed-render: only accept unadvertised format drags if WeChat is actively running
        if (CountFormats(pDataObj) == 0)
        {
            try { return System.Diagnostics.Process.GetProcessesByName("Weixin").Length > 0; }
            catch { return false; }
        }

        return false;
    }

    private bool HasFormat(IntPtr pDataObj, short cfFormat)
    {
        try
        {
            IntPtr vtbl = *(IntPtr*)pDataObj;
            var queryGetData = *(delegate* unmanaged[Stdcall]<IntPtr, IntPtr, int>*)(vtbl + 5 * IntPtr.Size);

            var fmt = new FORMATETC
            {
                cfFormat = cfFormat,
                dwAspect = DVASPECT.DVASPECT_CONTENT,
                lindex = -1,
                ptd = IntPtr.Zero,
                tymed = TYMED.TYMED_HGLOBAL
            };
            IntPtr pFmt = Marshal.AllocHGlobal(Marshal.SizeOf<FORMATETC>());
            try
            {
                Marshal.StructureToPtr(fmt, pFmt, false);
                return queryGetData(pDataObj, pFmt) == S_OK;
            }
            finally { Marshal.FreeHGlobal(pFmt); }
        }
        catch { return false; }
    }

    private int CountFormats(IntPtr pDataObj)
    {
        int count = 0;
        try
        {
            IntPtr vtbl = *(IntPtr*)pDataObj;
            var enumFormatEtc = *(delegate* unmanaged[Stdcall]<IntPtr, uint, IntPtr*, int>*)(vtbl + 8 * IntPtr.Size);

            IntPtr pEnum = IntPtr.Zero;
            if (enumFormatEtc(pDataObj, DATADIR_GET, &pEnum) != S_OK || pEnum == IntPtr.Zero)
                return 0;

            try
            {
                IntPtr enumVtbl = *(IntPtr*)pEnum;
                var next = *(delegate* unmanaged[Stdcall]<IntPtr, uint, IntPtr, int*, int>*)(enumVtbl + 3 * IntPtr.Size);

                IntPtr pFmt = Marshal.AllocHGlobal(Marshal.SizeOf<FORMATETC>());
                try
                {
                    int fetched = 0;
                    while (next(pEnum, 1, pFmt, &fetched) == S_OK && fetched == 1)
                    {
                        count++;
                        if (count > 64) break;
                        fetched = 0;
                    }
                }
                finally { Marshal.FreeHGlobal(pFmt); }
            }
            finally
            {
                IntPtr enumVtbl = *(IntPtr*)pEnum;
                var release = *(delegate* unmanaged[Stdcall]<IntPtr, uint>*)(enumVtbl + 2 * IntPtr.Size);
                release(pEnum);
            }
        }
        catch { }
        return count;
    }

    private List<string> ReadCfHdrop(IntPtr pDataObj)
    {
        var result = new List<string>();
        if (pDataObj == IntPtr.Zero) return result;
        try
        {
            IntPtr vtbl = *(IntPtr*)pDataObj;
            var getData = *(delegate* unmanaged[Stdcall]<IntPtr, IntPtr, IntPtr, int>*)(vtbl + 3 * IntPtr.Size);
            var queryGetData = *(delegate* unmanaged[Stdcall]<IntPtr, IntPtr, int>*)(vtbl + 5 * IntPtr.Size);

            var fmt = new FORMATETC
            {
                cfFormat = CF_HDROP,
                dwAspect = DVASPECT.DVASPECT_CONTENT,
                lindex = -1,
                ptd = IntPtr.Zero,
                tymed = TYMED.TYMED_HGLOBAL
            };
            IntPtr pFmt = Marshal.AllocHGlobal(Marshal.SizeOf<FORMATETC>());
            try
            {
                Marshal.StructureToPtr(fmt, pFmt, false);
                if (queryGetData(pDataObj, pFmt) != S_OK) return result;

                IntPtr pMed = Marshal.AllocHGlobal(Marshal.SizeOf<STGMEDIUM>());
                var med = default(STGMEDIUM);
                try
                {
                    if (getData(pDataObj, pFmt, pMed) != S_OK) return result;
                    med = Marshal.PtrToStructure<STGMEDIUM>(pMed);
                    if (med.unionmember != IntPtr.Zero)
                        result = ReadHDrop(med.unionmember);
                }
                finally
                {
                    if (med.unionmember != IntPtr.Zero)
                        NativeMethods.ReleaseStgMedium(ref med);
                    Marshal.FreeHGlobal(pMed);
                }
            }
            finally { Marshal.FreeHGlobal(pFmt); }
        }
        catch (Exception ex)
        {
            DebugLogger.Log($"[Win32DropTarget] ReadCfHdrop error: {ex.Message}");
        }
        return result;
    }

    private static List<string> ReadHDrop(IntPtr hDrop)
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
                if (!string.IsNullOrWhiteSpace(path))
                    result.Add(path);
            }
        }
        return result;
    }

    private List<string> DiffRwTemp()
    {
        var result = new List<string>();
        try
        {
            var current = Win32DragDropHelper.EnumerateWeChatRwTempFiles();
            var fresh = current.Where(f => !_rwTempSnapshot.Contains(f) && File.Exists(f)).ToList();
            result = GroupNewest(fresh);
        }
        catch (Exception ex)
        {
            DebugLogger.Log($"[Win32DropTarget] DiffRwTemp error: {ex.Message}");
        }
        return result;
    }

    /// <summary>One drag creates files inside a single GUID folder; pick the newest group.</summary>
    private static List<string> GroupNewest(List<string> files)
    {
        if (files.Count == 0) return new List<string>();
        var group = files.Select(f => new FileInfo(f))
            .GroupBy(f => f.DirectoryName)
            .OrderByDescending(g => g.Max(f => f.LastWriteTimeUtc))
            .FirstOrDefault();
        return group == null ? new List<string>() : group.Select(f => f.FullName).ToList();
    }

    private void SafeInvoke(List<string> paths, int screenX, int screenY)
    {
        try
        {
            if (paths == null || paths.Count == 0) return;
            _onFilesDropped(paths, screenX, screenY);
        }
        catch (Exception ex)
        {
            DebugLogger.Log($"[Win32DropTarget] onFilesDropped error: {ex.Message}");
        }
    }

    // ── Cleanup ─────────────────────────────────────────────────────────────────

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        try
        {
            if (_hwnd != IntPtr.Zero)
            {
                int hr = RevokeDragDrop(_hwnd);
                DebugLogger.Log($"[Win32DropTarget] RevokeDragDrop hr=0x{hr:X8}");
            }
        }
        catch (Exception ex)
        {
            DebugLogger.Log($"[Win32DropTarget] RevokeDragDrop error: {ex.Message}");
        }

        // Release the reference held by the managed wrapper.
        // If OLE still holds references (e.g. in-flight drag), unmanaged resources
        // remain valid until OLE calls Release.
        InternalRelease();
    }

    // ── P/Invoke ────────────────────────────────────────────────────────────────

    [StructLayout(LayoutKind.Sequential)]
    private struct POINTL { public int X; public int Y; }

    [DllImport("ole32.dll")]
    private static extern int OleInitialize(IntPtr pvReserved);

    [DllImport("ole32.dll")]
    private static extern int RegisterDragDrop(IntPtr hwnd, IntPtr pDropTarget);

    [DllImport("ole32.dll")]
    private static extern int RevokeDragDrop(IntPtr hwnd);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode, EntryPoint = "DragQueryFileW")]
    private static extern uint DragQueryFileW(IntPtr hDrop, uint iFile, System.Text.StringBuilder? lpszFile, uint cch);
}
