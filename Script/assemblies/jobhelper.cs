using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;

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
        HardCap = 0x4,
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
                ControlFlags = CpuRateControlFlags.Enable | CpuRateControlFlags.HardCap,
                CpuRate      = (uint)(cpuPercentage * 100)
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
}