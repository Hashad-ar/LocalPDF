using System;
using System.Linq;
using LocalPDF.Services;

namespace LocalPDF
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            if (args.Contains("--word-helper", StringComparer.OrdinalIgnoreCase) ||
                args.Contains("--merge-helper", StringComparer.OrdinalIgnoreCase))
                return HelperHost.Run(args);

            var application = new App();
            application.InitializeComponent();
            return application.Run();
        }
    }
}
