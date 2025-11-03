using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace DemoApp
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            EnableHighDpiAwareness();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            SplashScreenManager.Show();

            DemoApp mainForm = null;
            try
            {
                mainForm = new DemoApp();
            }
            catch (Exception ex)
            {
                SplashScreenManager.Close();
                MessageBox.Show($"Failed to start Envista Turntable Demo:\n{ex}", "Startup Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            SplashScreenManager.Close();
            Application.Run(mainForm);
        }

        private static void EnableHighDpiAwareness()
        {
            try
            {
                // Try Per-Monitor V2 awareness first (requires Windows 10 Anniversary Update or later)
                if (!SetProcessDpiAwarenessContext(new IntPtr(-4))) // DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
                {
                    SetProcessDPIAware();
                }
            }
            catch (EntryPointNotFoundException)
            {
                SetProcessDPIAware();
            }
            catch (DllNotFoundException)
            {
                // Ignore; process will fall back to system DPI awareness
            }
        }

        private static class SplashScreenManager
        {
            private static Thread splashThread;
            private static LoadingForm splashForm;

            public static void Show()
            {
                if (splashThread != null)
                {
                    return;
                }

                splashThread = new Thread(() =>
                {
                    splashForm = new LoadingForm();
                    Application.Run(splashForm);
                })
                {
                    IsBackground = true
                };
                splashThread.SetApartmentState(ApartmentState.STA);
                splashThread.Start();
            }

            public static void Close()
            {
                if (splashForm == null)
                {
                    return;
                }

                try
                {
                    splashForm.Invoke(new Action(() => splashForm.Close()));
                    splashThread?.Join(500);
                }
                catch (InvalidOperationException)
                {
                    // Splash already closed
                }
                finally
                {
                    splashForm = null;
                    splashThread = null;
                }
            }
        }

        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        [DllImport("user32.dll")]
        private static extern bool SetProcessDpiAwarenessContext(IntPtr dpiContext);
    }
}
