using System;
using System.IO;
using System.Linq;

namespace LocalPDF.Services
{
    internal sealed class TempWorkspace : IDisposable
    {
        internal static readonly string Root = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "LocalPDF");
        internal string Path { get; }

        internal TempWorkspace()
        {
            Directory.CreateDirectory(Root);
            Path = System.IO.Path.Combine(Root, Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path);
        }

        internal static void CleanupStale()
        {
            if (!Directory.Exists(Root)) return;
            foreach (var directory in Directory.EnumerateDirectories(Root))
            {
                try
                {
                    if (Directory.GetCreationTimeUtc(directory) < DateTime.UtcNow.AddHours(-12))
                        Directory.Delete(directory, true);
                }
                catch { }
            }
        }

        public void Dispose()
        {
            try { if (Directory.Exists(Path)) Directory.Delete(Path, true); } catch { }
        }
    }
}
