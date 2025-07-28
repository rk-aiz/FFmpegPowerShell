using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

namespace JobHelper
{
    enum JobObjectInfoType : int
    {
        CpuRateControlInformation = 15,
    }

    [Flags]
    enum CpuRateControlFlags : uint
    {
        Enable = 0x1,
        WEIGHT_BASED = 0x2,

        HardCap = 0x4,
        NOTIFY = 0x8,
        MIN_MAX_RATE = 0x10,
    }

    [StructLayout(LayoutKind.Sequential)]
    struct JOBOBJECT_CPU_RATE_CONTROL_INFORMATION
    {
        public CpuRateControlFlags ControlFlags;
        public uint CpuRate;  // 1/100% 単位。50% → 50 * 100 = 5000
    }

    static class NativeMethods
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr CreateJobObject(IntPtr lpJobAttributes, string name);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool SetInformationJobObject(
            IntPtr hJob,
            JobObjectInfoType infoType,
            ref JOBOBJECT_CPU_RATE_CONTROL_INFORMATION info,
            uint infoLength);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool AssignProcessToJobObject(
            IntPtr hJob,
            IntPtr hProcess);
    }

    public sealed class CpuLimitedJob : IDisposable
    {
        IntPtr _hJob;

        public CpuLimitedJob(int cpuPercentage)
        {
            // 1. ジョブ作成
            _hJob = NativeMethods.CreateJobObject(IntPtr.Zero, null);
            if (_hJob == IntPtr.Zero)
                ThrowLastError("CreateJobObject");

            // 2. CPUレート制御設定
            var info = new JOBOBJECT_CPU_RATE_CONTROL_INFORMATION
            {
                ControlFlags = CpuRateControlFlags.Enable | CpuRateControlFlags.WEIGHT_BASED, //CpuRateControlFlags.HardCap |
                CpuRate = (uint)(cpuPercentage * 100)
            };

            if (!NativeMethods.SetInformationJobObject(
                    _hJob,
                    JobObjectInfoType.CpuRateControlInformation,
                    ref info,
                    (uint)Marshal.SizeOf(info)))
            {
                ThrowLastError("SetInformationJobObject");
            }
        }

        public void AddProcess(Process process)
        {
            if (!NativeMethods.AssignProcessToJobObject(_hJob, process.Handle))
                ThrowLastError("AssignProcessToJobObject");
        }

        static void ThrowLastError(string name)
        {
            throw new Win32Exception(Marshal.GetLastWin32Error(), String.Format("{0} failed", name));
        }

        public void Dispose()
        {
            if (_hJob != IntPtr.Zero)
            {
                Marshal.FreeHGlobal(_hJob);
                _hJob = IntPtr.Zero;
            }
        }
    }


    public class ThrottledJobCpuController : IDisposable
    {
        // Win32 API 定義
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern IntPtr CreateJobObject(IntPtr lpJobAttributes, string lpName);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool SetInformationJobObject(
            IntPtr hJob,
            JOBOBJECTINFOCLASS infoClass,
            IntPtr lpJobObjectInfo,
            uint cbJobObjectInfoLength);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool AssignProcessToJobObject(IntPtr hJob, IntPtr hProcess);

        // CPU レート制御用構造体 (v2 API)
        [StructLayout(LayoutKind.Sequential)]
        struct JOBOBJECT_CPU_RATE_CONTROL_INFORMATION
        {
            public uint ControlFlags;
            public uint CpuRate;   // 1～10000 (1=0.01%, 10000=100%)
        }

        enum JOBOBJECTINFOCLASS : int
        {
            JobObjectCpuRateControlInformation = 15,
        }

        readonly IntPtr _hJob;
        readonly Timer _timer;
        readonly object _sync = new object();

        // クライアントがセットする「希望レート」
        // Interlocked.Exchange で書き換えられる
        int _targetRate;

        // 最後に適用したレート
        int _lastAppliedRate;

        // 最小更新間隔 (ms)
        const int MinIntervalMs = 1000;

        public ThrottledJobCpuController(int pollMs = 2000)
        {
            _hJob = CreateJobObject(IntPtr.Zero, null);
            if (_hJob == IntPtr.Zero) ThrowLastError("CreateJobObject");

            // タイマー開始：pollMs 毎にコールバック → 更新判定
            _timer = new Timer(UpdateIfNeeded, null, pollMs, pollMs);
        }

        public void AssignCurrentProcess()
        {
            using (var p = System.Diagnostics.Process.GetCurrentProcess())
            {
                if (!AssignProcessToJobObject(_hJob, p.Handle))
                    ThrowLastError("AssignProcessToJobObject");
            }
        }

        public void AssignProcess(Process process)
        {
            if (!AssignProcessToJobObject(_hJob, process.Handle))
                ThrowLastError("AssignProcessToJobObject");
        }

        /// <summary>
        /// クライアントから呼ぶ。頻繁に呼んでもOK。
        /// </summary>
        public void RequestRate(int rate)
        {
            if (rate < 1 || rate > 10000)
                throw new ArgumentOutOfRangeException("rate");

            // レースを避けつつ、最新値だけ保持
            Interlocked.Exchange(ref _targetRate, rate);
        }

        static int lastTick = 0;
        void UpdateIfNeeded(object state)
        {
            // ロックは不要。Interlocked で安全に読み出し
            int desired = Interlocked.CompareExchange(ref _targetRate, 0, 0);
            if (desired == _lastAppliedRate) return;

            // 最終更新から一定時間経っていなければスキップ
            // ※ 細かいロジックは要件に応じて調整
            lock (_sync)
            {
                // シンプルに時間計測
                // （DateTime.UtcNow または Stopwatch でも可）
                var now = Environment.TickCount;
                lastTick = 0;
                if (now - lastTick < MinIntervalMs)
                    return;
                lastTick = now;

                // Win32 API 呼び出し
                ApplyCpuRate(desired);
                _lastAppliedRate = desired;
            }
        }

        void ApplyCpuRate(int rate)
        {
            var info = new JOBOBJECT_CPU_RATE_CONTROL_INFORMATION
            {
                ControlFlags = (uint)CpuRateControlFlags.Enable | (uint)CpuRateControlFlags.WEIGHT_BASED,
                CpuRate = (uint)rate
            };
            int size = Marshal.SizeOf(info);
            IntPtr p = Marshal.AllocHGlobal(size);
            try
            {
                Marshal.StructureToPtr(info, p, false);

                if (!SetInformationJobObject(
                    _hJob,
                    JOBOBJECTINFOCLASS.JobObjectCpuRateControlInformation,
                    p,
                    (uint)size))
                {
                    ThrowLastError("SetInformationJobObject");
                }
            }
            finally
            {
                Marshal.FreeHGlobal(p);
            }
        }

        public void Dispose()
        {
            if (null != _timer) _timer.Dispose();
            // ジョブオブジェクトのクローズ処理
        }
        
        static void ThrowLastError(string name)
        {
            throw new Win32Exception(Marshal.GetLastWin32Error(), String.Format("{0} failed", name));
        }
    }

}