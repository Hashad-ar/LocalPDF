using System;
using System.Threading;
using System.Windows;
using LocalPDF.Services;

namespace LocalPDF
{
    public partial class App : Application
    {
        private Mutex _singleInstance;

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            bool createdNew;
            _singleInstance = new Mutex(true, @"Local\LocalPDF.8B385FE4-2638-4C77-BAD8-4560D1A4D263", out createdNew);
            if (!createdNew)
            {
                NativeMethods.FocusExistingWindow("LocalPDF");
                Shutdown();
                return;
            }

            TempWorkspace.CleanupStale();
            var window = new MainWindow();
            MainWindow = window;
            window.Show();
        }

        protected override void OnExit(ExitEventArgs e)
        {
            _singleInstance?.Dispose();
            base.OnExit(e);
        }
    }
}
