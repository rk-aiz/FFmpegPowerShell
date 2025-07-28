using System;
using System.Windows.Media;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.ConstrainedExecution;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Input;
using Microsoft.Win32.SafeHandles;


namespace HelperClasses
{
    internal class SafeThreadHandle : SafeHandleZeroOrMinusOneIsInvalid
    {
        [DllImport("kernel32", SetLastError = true)]
        [ReliabilityContract(Consistency.WillNotCorruptState, Cer.MayFail)]
        internal extern static bool CloseHandle(IntPtr handle);

        public SafeThreadHandle(IntPtr handle) : this()
        {
            SetHandle(handle);
        }

        private SafeThreadHandle() : base(true)
        { }

        [ReliabilityContract(Consistency.WillNotCorruptState, Cer.MayFail)]
        override protected bool ReleaseHandle()
        {
            return CloseHandle(handle);
        }
    }

    public class WindowHelper
    {

        [DllImport("user32.dll")]
        private static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private const int SW_HIDE = 0;
        private const int SW_SHOW = 5;

        private IntPtr _hWnd = IntPtr.Zero;

        public WindowHelper(string windowTitle, int timeOutMillis = 10000)
        {
            RetryUntilSuccessOrTimeout(() =>
            {
                var targetProcess = Process.GetProcesses().FirstOrDefault(p => p.MainWindowTitle == windowTitle);
                if (targetProcess != null)
                {
                    _hWnd = targetProcess.MainWindowHandle;
                    return true;
                }
                else
                {
                    return false;
                }
            }, TimeSpan.FromMilliseconds(timeOutMillis), TimeSpan.FromMilliseconds(100));
        }

        public static bool RetryUntilSuccessOrTimeout(Func<bool> task, TimeSpan timeout, TimeSpan pause)
        {
            if (pause.TotalMilliseconds < 0)
            {
                throw new ArgumentException("pause must be >= 0 milliseconds");
            }
            var stopwatch = Stopwatch.StartNew();
            do
            {
                if (task()) { return true; }
                Thread.Sleep((int)pause.TotalMilliseconds);
            }
            while (stopwatch.Elapsed < timeout);
            return false;
        }

        public void HideConsole()
        {
            if (_hWnd != IntPtr.Zero)
            {
                ShowWindowAsync(_hWnd, SW_HIDE);
            }
        }

        public void ShowConsole(bool activate = true)
        {
            if (_hWnd != IntPtr.Zero)
            {
                ShowWindowAsync(_hWnd, SW_SHOW);

                if (activate)
                    SetForegroundWindow(_hWnd);
            }
        }
    }

    public class ReceivedData
    {
        public enum DataType
        {
            StdOut,
            StdError,
        }

        public string Data { get; private set; }
        public DataType Type { get; private set; }

        public ReceivedData(string data, DataType type)
        {
            Data = data;
            Type = type;
        }

        private ReceivedData() { }

        public ReceivedData Empty
        {
            get { return new ReceivedData(); }
        }
    }

    public class ProcessInfo : IDisposable
    {
        [Flags]
        public enum ThreadAccess : int
        {
            TERMINATE = (0x0001),
            SUSPEND_RESUME = (0x0002),
            GET_CONTEXT = (0x0008),
            SET_CONTEXT = (0x0010),
            SET_INFORMATION = (0x0020),
            QUERY_INFORMATION = (0x0040),
            SET_THREAD_TOKEN = (0x0080),
            IMPERSONATE = (0x0100),
            DIRECT_IMPERSONATION = (0x0200)
        }

        [DllImport("kernel32.dll")]
        static extern SafeThreadHandle OpenThread(ThreadAccess dwDesiredAccess, bool bInheritHandle, uint dwThreadId);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern bool CloseHandle(IntPtr handle);
        [DllImport("kernel32.dll")]
        static extern uint SuspendThread(SafeThreadHandle hThread);
        [DllImport("kernel32.dll")]
        static extern int ResumeThread(SafeThreadHandle hThread);

        private bool _disposed = false;

        private const int MAX_QUE = 512;
        private ProcessStartInfo _startInfo;
        private TaskCompletionSource<int> _eventHandled;
        private Process _process;

        public bool IsSuspended { get; private set; }

        public bool TryGetProcess(out Process process)
        {
            process = _process;
            return _process != null;
        }

        private BlockingCollection<ReceivedData> _receivedData = new BlockingCollection<ReceivedData>(new ConcurrentQueue<ReceivedData>(), MAX_QUE);

        public BlockingCollection<ReceivedData> ReceivedDataQueue
        {
            get { return _receivedData; }
        }
        public ProcessInfo(string filename, string argumentList, bool noRedirect = false)
        {
            _startInfo = new ProcessStartInfo
            {
                FileName = filename,
                Arguments = argumentList,
                UseShellExecute = false,
                RedirectStandardOutput = !noRedirect,
                RedirectStandardError = !noRedirect,
                RedirectStandardInput = !noRedirect
            };

            IsSuspended = false;
        }

        public void Suspend()
        {
            if (IsSuspended) { return; }

            Process process;
            if (TryGetProcess(out process))
                foreach (ProcessThread thread in process.Threads)
                {
                    using (var pOpenThread = OpenThread(ThreadAccess.SUSPEND_RESUME, false, (uint)thread.Id))
                    {
                        if (pOpenThread != null && (!pOpenThread.IsInvalid))
                            SuspendThread(pOpenThread);
                    }
                }
            IsSuspended = true;
        }

        public void Resume()
        {
            if (!IsSuspended) { return; }

            Process process;
            if (TryGetProcess(out process))
                foreach (ProcessThread thread in process.Threads)
                {
                    using (var pOpenThread = OpenThread(ThreadAccess.SUSPEND_RESUME, false, (uint)thread.Id))
                    {
                        if (pOpenThread != null && (!pOpenThread.IsInvalid))
                            ResumeThread(pOpenThread);
                    }
                }
            IsSuspended = false;
        }

        public async Task<int> Start()
        {

            if (_disposed) { throw new ObjectDisposedException("\"Dispose\" has already been executed."); }

            _eventHandled = new TaskCompletionSource<int>();

            using (_process = new Process { StartInfo = _startInfo, EnableRaisingEvents = true })
            {
                try
                {
                    if (!_startInfo.UseShellExecute)
                    {
                        _process.OutputDataReceived += (sender, e) =>
                        {
                            Task.Run(() =>
                            {
                                if (!String.IsNullOrEmpty(e.Data))
                                {
                                    _receivedData.TryAdd(
                                        new ReceivedData(
                                            e.Data,
                                            ReceivedData.DataType.StdOut
                                        ),
                                        System.Threading.Timeout.Infinite
                                    );
                                }
                            });
                        };
                        _process.ErrorDataReceived += (sender, e) =>
                        {
                            Task.Run(() =>
                            {
                                if (!String.IsNullOrEmpty(e.Data))
                                {
                                    _receivedData.TryAdd(
                                        new ReceivedData(
                                            e.Data,
                                            ReceivedData.DataType.StdError
                                        ),
                                        System.Threading.Timeout.Infinite
                                    );
                                }
                            });
                        };
                    }
                    _process.Exited += (sender, e) =>
                    {
                        _receivedData.CompleteAdding();
                        _eventHandled.TrySetResult(_process.ExitCode);
                    };
                    _process.Start();
                    if (!_startInfo.UseShellExecute)
                    {
                        _process.BeginErrorReadLine();
                        _process.BeginOutputReadLine();
                    }
                    //_process.PriorityClass = ProcessPriorityClass.High;
                }
                catch (Exception e)
                {
                    throw e;
                }
                await _eventHandled.Task;
            }

            _process = null;
            return _eventHandled.Task.Result;
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!this._disposed)
            {
                if (disposing)
                {
                    _receivedData.Dispose();
                }
                _disposed = true;
            }
        }
    }

    public class ThreadHelper
    {

        [FlagsAttribute]
        public enum EXECUTION_STATE : uint
        {
            ES_AWAYMODE_REQUIRED = 0x00000040,
            ES_CONTINUOUS = 0x80000000,
            ES_DISPLAY_REQUIRED = 0x00000002,
            ES_SYSTEM_REQUIRED = 0x00000001
            // Legacy flag, should not be used.
            // ES_USER_PRESENT = 0x00000004
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern EXECUTION_STATE SetThreadExecutionState(EXECUTION_STATE esFlags);

        public static void PreventSleep(EXECUTION_STATE state = EXECUTION_STATE.ES_SYSTEM_REQUIRED)
        {
            SetThreadExecutionState(EXECUTION_STATE.ES_CONTINUOUS | state);
        }

        public static void AllowSleep()
        {
            SetThreadExecutionState(EXECUTION_STATE.ES_CONTINUOUS);
        }
    }

    public class CommentOutText
    {
        private const string pattern1 = "//.*";
        private const string pattern2 = @" {2,}";
        private char[] pattern3 = "\r\n".ToCharArray();

        private List<string> _strings;
        public CommentOutText(string text)
        {
            var temp = Regex.Replace(text, pattern1, String.Empty);
            _strings = Regex.Replace(temp, pattern2, " ").Split(pattern3).ToList();
            _strings.RemoveAll(s => String.IsNullOrWhiteSpace(s));
        }

        public void Replace(string pattern, string replace, RegexOptions regexOptions = RegexOptions.None)
        {
            var temp = Regex.Replace(String.Join("\n", _strings), pattern, replace);
            _strings = temp.Split('\n').ToList();
        }

        public Match Match(string pattern)
        {
            return Regex.Match(String.Join(String.Empty, _strings), pattern);
        }

        public string GetText(string separator)
        {
            return String.Join(separator, _strings);
        }
    }


    public static class ConsoleHelper
    {
        public static void WriteLine(object message, int beforLines = 0, int afterLines = 0)
        {
            string format =
                String.Concat(Enumerable.Repeat(Environment.NewLine, beforLines)) +
                "{0}" + String.Concat(Enumerable.Repeat(Environment.NewLine, afterLines));

            string formattedMessage = String.Format(
                CultureInfo.CurrentUICulture,
                format,
                message
            );
            Console.WriteLine(formattedMessage);
        }
        public static void Log(object message, int beforLines = 0, int afterLines = 0)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            WriteLine(message, beforLines, afterLines);
            Console.ResetColor();
        }

        public static void Error(object message, int beforLines = 0, int afterLines = 0)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            WriteLine(message, beforLines, afterLines);
            Console.ResetColor();
        }

        public static void Info(object message, int beforLines = 0, int afterLines = 0)
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            WriteLine(message, beforLines, afterLines);
            Console.ResetColor();
        }
    }

    public static class FontHelper
    {
        public static FontFamily LoadFont(string fontName, string path, UriKind uriKind = UriKind.Absolute)
        {

            var fontUri = new Uri(path, UriKind.Absolute);

            // FontFamily: URI + #フォントファミリー名
            var fontFamily = new FontFamily(fontUri, String.Format("#{0}", fontName));

            return fontFamily;
        }
    } 

}