using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;

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
            Application.Run(new DemoApp());
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

        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        [DllImport("user32.dll")]
        private static extern bool SetProcessDpiAwarenessContext(IntPtr dpiContext);
    }
}
