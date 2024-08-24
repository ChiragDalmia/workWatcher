using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.IO;
using System.Collections.Concurrent;

namespace workWatcher
{
    public partial class Service1 : ServiceBase
    {
        // Windows API imports
        private static class NativeMethods
        {
            [DllImport("user32.dll")]
            public static extern IntPtr GetForegroundWindow();

            [DllImport("user32.dll")]
            public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

            [DllImport("user32.dll")]
            public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        }
        // Constants
        private const int MAX_TITLE_LENGTH = 256;
        private const string LOG_FILE_NAME = "activity_log.csv";
        private const int MIN_DURATION_SECONDS = 5;
        private const int CHECK_INTERVAL_MS = 1000;
        private const string LOG_DIRECTORY = "Logs";
        private const string LOG_FILE_PREFIX = "activity_log_";


        // Thread-safe queue for logging
        private static readonly ConcurrentQueue<string> logQueue = new ConcurrentQueue<string>();

        // Cancellation token to stop the tracking thread
        private CancellationTokenSource cancellationTokenSource;

        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            cancellationTokenSource = new CancellationTokenSource();
            Thread trackerThread = new Thread(() => RunActivityTracker(cancellationTokenSource.Token));
            trackerThread.IsBackground = true;
            trackerThread.Start();
        }

        protected override void OnStop()
        {
            cancellationTokenSource?.Cancel();
            // Allow time for the tracker and logger threads to finish
            Thread.Sleep(2000);
        }
        private void RunActivityTracker(CancellationToken cancellationToken)
        {
            InitializeLogFile();

            string lastWindowTitle = null;
            string lastProcessName = null;
            DateTime activityStartTime = DateTime.Now;

            // Start a separate thread for logging
            Thread loggerThread = new Thread(LoggerThreadWork);
            loggerThread.IsBackground = true;
            loggerThread.Start();

            while (!cancellationToken.IsCancellationRequested)
            {
                string currentWindowTitle = GetActiveWindowTitle();
                string currentProcessName = GetActiveProcessName();

                if (HasActivityChanged(currentWindowTitle, currentProcessName, lastWindowTitle, lastProcessName))
                {
                    LogActivityIfDurationMet(lastWindowTitle, lastProcessName, activityStartTime);

                    lastWindowTitle = currentWindowTitle;
                    lastProcessName = currentProcessName;
                    activityStartTime = DateTime.Now;
                }

                Thread.Sleep(CHECK_INTERVAL_MS);
            }

            // Log the final activity if it meets the duration threshold
            LogActivityIfDurationMet(lastWindowTitle, lastProcessName, activityStartTime);

            // Allow time for the logger thread to finish
            Thread.Sleep(1000);
        }
        private void InitializeLogFile()
        {
            string logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, LOG_DIRECTORY);
            if (!Directory.Exists(logDirectory))
            {
                Directory.CreateDirectory(logDirectory);
            }

            string logFilePath = GetCurrentLogFilePath();
            if (!File.Exists(logFilePath))
            {
                File.WriteAllText(logFilePath, "StartTimestamp,EndTimestamp,Duration,ProcessName,WindowTitle\n");
            }
        }

        private string GetCurrentLogFilePath()
        {
            string logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, LOG_DIRECTORY);
            string fileName = $"{LOG_FILE_PREFIX}{DateTime.Now:yyyyMMdd}.csv";
            return Path.Combine(logDirectory, fileName);
        }

        private bool HasActivityChanged(string currentWindowTitle, string currentProcessName, string lastWindowTitle, string lastProcessName)
        {
            return currentWindowTitle != lastWindowTitle || currentProcessName != lastProcessName;
        }

        private void LogActivityIfDurationMet(string windowTitle, string processName, DateTime startTime)
        {
            if (windowTitle != null && processName != null)
            {
                TimeSpan duration = DateTime.Now - startTime;
                if (duration.TotalSeconds >= MIN_DURATION_SECONDS)
                {
                    string logEntry = FormatLogEntry(windowTitle, processName, startTime, duration);
                    logQueue.Enqueue(logEntry);
                }
            }
        }

        private string GetActiveWindowTitle()
        {
            StringBuilder buff = new StringBuilder(MAX_TITLE_LENGTH);
            IntPtr handle = NativeMethods.GetForegroundWindow();

            if (NativeMethods.GetWindowText(handle, buff, MAX_TITLE_LENGTH) > 0)
            {
                return buff.ToString();
            }
            return null;
        }

        private string GetActiveProcessName()
        {
            try
            {
                IntPtr handle = NativeMethods.GetForegroundWindow();
                NativeMethods.GetWindowThreadProcessId(handle, out uint processId);
                using (Process process = Process.GetProcessById((int)processId))
                {
                    return process.ProcessName;
                }
            }
            catch (ArgumentException)
            {
                return "Unknown";
            }
        }

        private string FormatLogEntry(string windowTitle, string processName, DateTime startTime, TimeSpan duration)
        {
            return $"{startTime:yyyy-MM-dd HH:mm:ss}," +
                   $"{startTime.Add(duration):yyyy-MM-dd HH:mm:ss}," +
                   $"{duration:hh\\:mm\\:ss}," +
                   $"{EscapeCsvField(processName)}," +
                   $"{EscapeCsvField(windowTitle)}";
        }

        private string EscapeCsvField(string field)
        {
            if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
            {
                return $"\"{field.Replace("\"", "\"\"")}\"";
            }
            return field;
        }

        private void LoggerThreadWork()
        {
            while (true)
            {
                if (logQueue.TryDequeue(out string logEntry))
                {
                    try
                    {
                        string logFilePath = GetCurrentLogFilePath();
                        File.AppendAllText(logFilePath, logEntry + Environment.NewLine);
                    }
                    catch (IOException ex)
                    {
                        LogError(ex);
                    }
                }
                else
                {
                    Thread.Sleep(1000); // Sleep if queue is empty
                }
            }
        }
        private void LogError(Exception ex)
        {
            try
            {
                string errorLog = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"error_log_{DateTime.Now:yyyyMMdd}.txt");
                string errorMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}\n\n";
                File.AppendAllText(errorLog, errorMessage);
            }
            catch
            {
                // If we can't even log the error, just swallow it to avoid crashing
            }
        }
    }
}
