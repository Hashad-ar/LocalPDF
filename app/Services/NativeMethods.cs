using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace LocalPDF.Services
{
    internal static class NativeMethods
    {
        internal const uint JobObjectExtendedLimitInformation = 9;
        internal const uint JobObjectLimitKillOnJobClose = 0x00002000;
        private const int SwRestore = 9;
        private const int DwmwaUseImmersiveDarkMode = 20;

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        internal static extern IntPtr CreateJobObject(IntPtr securityAttributes, string name);

        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool SetInformationJobObject(IntPtr job, uint infoClass, IntPtr info, uint length);

        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);

        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool CloseHandle(IntPtr handle);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr FindWindow(string className, string windowName);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr window, int command);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr window);

        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr window, int attribute, ref int value, int size);

        internal static void FocusExistingWindow(string title)
        {
            var handle = FindWindow(null, title);
            if (handle == IntPtr.Zero) return;
            ShowWindow(handle, SwRestore);
            SetForegroundWindow(handle);
        }

        internal static void EnableDarkTitleBar(IntPtr handle)
        {
            var enabled = 1;
            DwmSetWindowAttribute(handle, DwmwaUseImmersiveDarkMode, ref enabled, sizeof(int));
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct IoCounters
        {
            internal ulong ReadOperationCount, WriteOperationCount, OtherOperationCount;
            internal ulong ReadTransferCount, WriteTransferCount, OtherTransferCount;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct BasicLimitInformation
        {
            internal long PerProcessUserTimeLimit, PerJobUserTimeLimit;
            internal uint LimitFlags;
            internal UIntPtr MinimumWorkingSetSize, MaximumWorkingSetSize;
            internal uint ActiveProcessLimit;
            internal UIntPtr Affinity;
            internal uint PriorityClass, SchedulingClass;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct ExtendedLimitInformation
        {
            internal BasicLimitInformation BasicLimitInformation;
            internal IoCounters IoInfo;
            internal UIntPtr ProcessMemoryLimit, JobMemoryLimit, PeakProcessMemoryUsed, PeakJobMemoryUsed;
        }
    }
}
