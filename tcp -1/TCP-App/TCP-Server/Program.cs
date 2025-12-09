using System;
using System.IO;
using System.Windows.Forms;

namespace TCP_Server
{
    static class Program
    {
        public static readonly string LogFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "TCPServer_Log.txt");

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                // Log application start
                LogToFile("Application starting...");
                
                // Set up global exception handlers
                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
                Application.ThreadException += (s, e) => HandleException(e.Exception);
                AppDomain.CurrentDomain.UnhandledException += (s, e) => 
                    HandleException((Exception)e.ExceptionObject);

                // Initialize and run the application
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                
                LogToFile("Creating main form...");
                var mainForm = new Form1();
                
                LogToFile("Starting application message loop...");
                Application.Run(mainForm);
                
                LogToFile("Application exiting normally.");
            }
            catch (Exception ex)
            {
                HandleException(ex);
            }
        }

        private static void HandleException(Exception ex)
        {
            string errorMessage = $"Unhandled exception: {ex}";
            LogToFile(errorMessage);
            
            // Show error to user
            MessageBox.Show(
                "An unexpected error occurred. Please check the log file for details.",
                "Application Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
                
            Environment.Exit(1);
        }
        
        private static void LogToFile(string message)
        {
            try
            {
                string logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}";
                File.AppendAllText(LogFilePath, logMessage + Environment.NewLine);
                Console.WriteLine(logMessage);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write to log file: {ex.Message}");
            }
        }
    }
}