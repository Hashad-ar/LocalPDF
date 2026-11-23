using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace LocalPDF.Services
{
    internal sealed class WordJob : IDisposable
    {
        private IntPtr _handle;

        internal WordJob()
        {
            _handle = NativeMethods.CreateJobObject(IntPtr.Zero, null);
            if (_handle == IntPtr.Zero) throw new InvalidOperationException("Could not create a protected Word process group.");
            var info = new NativeMethods.ExtendedLimitInformation();
            info.BasicLimitInformation.LimitFlags = NativeMethods.JobObjectLimitKillOnJobClose;
            int size = Marshal.SizeOf(info);
            IntPtr pointer = Marshal.AllocHGlobal(size);
            try
            {
                Marshal.StructureToPtr(info, pointer, false);
                if (!NativeMethods.SetInformationJobObject(_handle, NativeMethods.JobObjectExtendedLimitInformation, pointer, (uint)size))
                    throw new InvalidOperationException("Could not configure the protected Word process group.");
            }
            finally { Marshal.FreeHGlobal(pointer); }
        }

        internal void Add(Process process)
        {
            if (!NativeMethods.AssignProcessToJobObject(_handle, process.Handle))
                throw new InvalidOperationException("Could not take ownership of the Word process.");
        }

        public void Dispose()
        {
            if (_handle == IntPtr.Zero) return;
            NativeMethods.CloseHandle(_handle);
            _handle = IntPtr.Zero;
        }
    }
}
