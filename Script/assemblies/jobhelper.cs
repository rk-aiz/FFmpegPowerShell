using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

namespace JobHelper
{
    internal static class NativeMethods
    {
        [Flags]
        internal enum CpuRateControlFlags : uint
        {
            Enable = 0x1,
            WEIGHT_BASED = 0x2,
            HardCap = 0x4,
            NOTIFY = 0x8,
            MIN_MAX_RATE = 0x10,
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct JOBOBJECT_CPU_RATE_CONTROL_INFORMATION
        {
            internal CpuRateControlFlags ControlFlags;
            internal uint CpuRate;  // CpuRateControlFlags.WEIGHT_BASED → 1 - 9, (HardCap → 1 - 10000)
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct JOBOBJECT_BASIC_ACCOUNTING_INFORMATION
        {
            internal long TotalUserTime; // 100-nanosecond ticks
            internal long TotalKernelTime; // 100-nanosecond ticks
            internal long ThisPeriodTotalUserTime;
            internal long ThisPeriodTotalKernelTime;
            internal uint TotalPageFaultCount;
            internal uint TotalProcesses;
            internal uint ActiveProcesses;
            internal uint TotalTerminatedProcesses;
        }

        internal enum JOBOBJECTINFOCLASS : int
        {
            JobObjectBasicAccountingInformation = 1,
            JobObjectCpuRateControlInformation = 15,
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern IntPtr CreateJobObject(IntPtr lpJobAttributes, string name);

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern bool SetInformationJobObject(
            IntPtr hJob,
            JOBOBJECTINFOCLASS infoType,
            ref JOBOBJECT_CPU_RATE_CONTROL_INFORMATION info,
            uint infoLength
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern bool QueryInformationJobObject(
            IntPtr hJob,
            JOBOBJECTINFOCLASS JobObjectInformationClass,
            out JOBOBJECT_BASIC_ACCOUNTING_INFORMATION lpJobObjectInformation,
            uint cbJobObjectInformationLength,
            IntPtr lpReturnLength
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern bool AssignProcessToJobObject(IntPtr hJob, IntPtr hProcess);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool CloseHandle(IntPtr hObject);
    }

    public class JobCpuWeightController : IDisposable
    {
        readonly IntPtr _hJob;
        readonly Timer _timer;
        readonly object _sync = new object();
                
        // For CPU Usage Calculation
        private readonly Stopwatch _wallClock = Stopwatch.StartNew();
        private long _lastQueryWallClockTicks;
        private long _lastQueryCpuTime;

        // クライアントがセットする「希望レート」
        // Interlocked.Exchange で書き換えられる
        private int _targetWeight;

        // 最後に適用したレート
        private int _lastAppliedWeight;

        // 最後に適用した時刻 (TickCount)
        private int _lastAppliedTick;

        // 最小更新間隔 (ms)
        private const int MinIntervalMs = 1000;

        public JobCpuWeightController(int pollMs = 500)
        {
            _hJob = NativeMethods.CreateJobObject(IntPtr.Zero, null);
            if (_hJob == IntPtr.Zero) ThrowLastError("CreateJobObject");

            // タイマー開始：pollMs 毎にコールバック → 更新判定
            _timer = new Timer(UpdateWeightIfNeeded, null, pollMs, pollMs);

            // Initialize CPU usage counters
            GetCpuUsage();
        }

        public void AssignCurrentProcess()
        {
            using (var p = Process.GetCurrentProcess())
            {
                if (!NativeMethods.AssignProcessToJobObject(_hJob, p.Handle))
                    ThrowLastError("AssignProcessToJobObject");
            }
        }

        public void AssignProcess(Process proc)
        {
            if (!NativeMethods.AssignProcessToJobObject(_hJob, proc.Handle))
                ThrowLastError("AssignProcessToJobObject");
        }

        /// <summary>
        /// クライアントから呼ぶ。頻繁に呼んでもOK。
        /// </summary>
        public void RequestWeight(int weight)
        {
            if (weight < 1 || weight > 9)
                throw new ArgumentOutOfRangeException("weight", "Weight must be between 1 and 9.");

            // レースを避けつつ、最新値だけ保持
            Interlocked.Exchange(ref _targetWeight, weight);
        }

        private void UpdateWeightIfNeeded(object state)
        {
            // Volatile.Read は、マルチスレッド環境で最新の値を確実に読み取ることを保証します。
            int desired = Volatile.Read(ref _targetWeight);

            // 更新不要な場合はすぐにリターン
            if (desired == 0 || desired == _lastAppliedWeight) return;

            lock (_sync)
            {
                // lockブロック内で再度チェックすることで、待機中に値が変わった場合に対応します (Double-checked locking)
                if (desired != Volatile.Read(ref _targetWeight) || desired == _lastAppliedWeight)
                {
                    return;
                }
                var now = Environment.TickCount;
                // uncheckedコンテキストで計算することで、TickCountのロールオーバー(一周して0に戻ること)を安全に扱います。
                if (unchecked(now - _lastAppliedTick) < MinIntervalMs)
                {
                    return;
                }

                // Win32 API 呼び出しで Weightを適用
                ApplyJobWeight(desired);

                _lastAppliedWeight = desired;
                _lastAppliedTick = now;
            }
        }

        public void ApplyJobWeight(int weight)
        {
            if (weight < 1 || weight > 9)
                throw new ArgumentOutOfRangeException("weight", "Weight must be between 1 and 9.");

            var info = new NativeMethods.JOBOBJECT_CPU_RATE_CONTROL_INFORMATION
            {
                ControlFlags = NativeMethods.CpuRateControlFlags.Enable | NativeMethods.CpuRateControlFlags.WEIGHT_BASED,
                CpuRate = (uint)weight
            };
            SetCpuRateControl(info);
        }

        /// <summary>
        /// Gets the average CPU usage of the job since the last call to this method.
        /// /// </summary>
        /// <returns>The CPU usage percentage. e.g., 50.0 means 50% of one core is being used.</returns>
        public double GetCpuUsage(bool normalized = false)
        {
            NativeMethods.JOBOBJECT_BASIC_ACCOUNTING_INFORMATION accountingInfo;
            if (!NativeMethods.QueryInformationJobObject(
                _hJob,
                NativeMethods.JOBOBJECTINFOCLASS.JobObjectBasicAccountingInformation,
                out accountingInfo,
                (uint)Marshal.SizeOf<NativeMethods.JOBOBJECT_BASIC_ACCOUNTING_INFORMATION>(),
                IntPtr.Zero))
            {
                ThrowLastError("QueryInformationJobObject");
            }

            long currentWallClockTicks = _wallClock.Elapsed.Ticks;
            long currentCpuTime = accountingInfo.TotalUserTime + accountingInfo.TotalKernelTime;

            long elapsedWallClock = currentWallClockTicks - _lastQueryWallClockTicks;
            long elapsedCpuTime = currentCpuTime - _lastQueryCpuTime;

            // Store current values for the next calculation
            _lastQueryWallClockTicks = currentWallClockTicks;
            _lastQueryCpuTime = currentCpuTime;

            if (elapsedWallClock <= 0)
            {
                return 0.0;
            }

            // Calculate the usage percentage
            double rawUsage = (double)elapsedCpuTime / elapsedWallClock;

            if (normalized)
            {
                return (rawUsage / Environment.ProcessorCount) * 100.0;
            }
            return rawUsage * 100.0;
        }

        private void SetCpuRateControl(NativeMethods.JOBOBJECT_CPU_RATE_CONTROL_INFORMATION info)
        {
            if (!NativeMethods.SetInformationJobObject(
                    _hJob,
                    NativeMethods.JOBOBJECTINFOCLASS.JobObjectCpuRateControlInformation,
                    ref info,
                    (uint)Marshal.SizeOf(info)))
            {
                ThrowLastError("SetInformationJobObject");
            }
        }

        public void Dispose()
        {
            if (null != _timer) _timer.Dispose();
            // ジョブオブジェクトのクローズ処理
            if (_hJob != IntPtr.Zero)
            {
                NativeMethods.CloseHandle(_hJob);
            }
        }

        static void ThrowLastError(string name)
        {
            throw new Win32Exception(Marshal.GetLastWin32Error(), String.Format("{0} failed", name));
        }
    }
}