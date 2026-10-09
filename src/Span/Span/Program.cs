using Microsoft.Windows.AppLifecycle;
using System;

namespace Span;

class Program
{
    [STAThread]
    static int Main(string[] args)
    {
        // Issue #36: ColorCode 등 라이브러리의 Regex catastrophic backtracking이 UI 스레드를
        // 멈추지 않도록 기본 매치 타임아웃을 1초로 건다.
        //
        // 반드시 프로세스에서 가장 먼저 실행돼야 한다. Regex는 기본 타임아웃을 처음 쓰일 때
        // 정적 필드에 읽어 고정하므로, 어떤 코드든 Regex를 먼저 만들면 이후 SetData는 무시된다.
        // 이전에는 App 생성자 중간에 있었는데 그 앞의 Sentry 초기화(CrashReportingService)가
        // Regex를 먼저 만들어서, 이 타임아웃은 출하 빌드에서 한 번도 적용된 적이 없었다
        // (실측: 무한). App 생성자의 시작 로그가 실제 적용값을 기록한다.
        AppDomain.CurrentDomain.SetData("REGEX_DEFAULT_MATCH_TIMEOUT", TimeSpan.FromSeconds(1));

        // 와일드카드 필터·검색(Helpers.WildcardRegex)은 '*'가 여럿인 패턴(*a*b)에 NonBacktracking
        // 엔진을 쓴다. 이 엔진은 오토마톤 크기에 상한이 있어 기본값에서는 199자 패턴부터 생성이
        // 실패하고, 그러면 백트래킹 엔진으로 되돌아간다. 그 모양은 백트래킹에서 다항 시간이라
        // '*a' x 100 + 'b' 같은 필터가 긴 이름에서 끝나지 않는다(실측 5초 이상). 상한을 올려
        // 그 경로를 막는다. 이 값은 Regex를 만들 때마다 읽으므로 위치 제약은 없다. int여야 한다.
        AppDomain.CurrentDomain.SetData("REGEX_NONBACKTRACKING_MAX_AUTOMATA_SIZE", 100_000);

        WinRT.ComWrappersSupport.InitializeComWrappers();

        // Kill any conflicting older instances (e.g. Store app in WindowsApps)
        try
        {
            var curPid = System.Diagnostics.Process.GetCurrentProcess().Id;
            var curExe = Environment.ProcessPath;
            foreach (var p in System.Diagnostics.Process.GetProcessesByName("Span"))
            {
                if (p.Id != curPid)
                {
                    try
                    {
                        var pPath = p.MainModule?.FileName;
                        if (!string.Equals(pPath, curExe, StringComparison.OrdinalIgnoreCase) ||
                            pPath?.Contains("WindowsApps", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            p.Kill();
                            p.WaitForExit(1000);
                        }
                    }
                    catch { }
                }
            }
        }
        catch { }

        var isRedirect = DecideRedirection();
        if (!isRedirect)
        {
            Microsoft.UI.Xaml.Application.Start((p) =>
            {
                var context = new Microsoft.UI.Dispatching.DispatcherQueueSynchronizationContext(
                    Microsoft.UI.Dispatching.DispatcherQueue.GetForCurrentThread());
                System.Threading.SynchronizationContext.SetSynchronizationContext(context);
                new App();
            });
        }

        return 0;
    }

    private static bool DecideRedirection()
    {
        var appInstance = AppInstance.FindOrRegisterForKey("SPAN_FINDER_MAIN");

        if (appInstance.IsCurrent)
            return false; // 첫 인스턴스 — 정상 실행

        // 기존 인스턴스로 활성화 리다이렉트
        var activatedArgs = AppInstance.GetCurrent().GetActivatedEventArgs();
        appInstance.RedirectActivationToAsync(activatedArgs).AsTask().Wait();
        return true; // 현재 프로세스 종료
    }
}
