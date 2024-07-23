using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
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
        private const int MIN_CHECK_INTERVAL_MS = 1000;
        private const int MAX_CHECK_INTERVAL_MS = 5000;
        private const int LOG_BATCH_SIZE = 10;
        private const int LOG_ROTATION_SIZE_MB = 10;

        // Thread-safe queue for logging
        private static readonly ConcurrentQueue<string> logQueue = new ConcurrentQueue<string>();

        private Thread activityTrackerThread;
        private volatile bool shouldStop = false;
        private int currentCheckInterval = MIN_CHECK_INTERVAL_MS;
        private List<string> logBatch = new List<string>();
        private HashSet<string> recentWindowTitles = new HashSet<string>();

        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            EventLog.WriteEntry("WorkWatcher Service", "Service is starting...", EventLogEntryType.Information);

            activityTrackerThread = new Thread(RunActivityTracker);
            activityTrackerThread.IsBackground = true;
            activityTrackerThread.Start();

            EventLog.WriteEntry("WorkWatcher Service", "Service started successfully.", EventLogEntryType.Information);
        }


        protected override void OnStop()
        {
            EventLog.WriteEntry("WorkWatcher Service", "Service is stopping...", EventLogEntryType.Information);

            shouldStop = true;
            activityTrackerThread.Join(TimeSpan.FromSeconds(30)); // Wait for the thread to finish

            EventLog.WriteEntry("WorkWatcher Service", "Service stopped successfully.", EventLogEntryType.Information);
        }

        private void RunActivityTracker()
        {
            InitializeLogFile();

            string lastWindowTitle = null;
            string lastProcessName = null;
            DateTime activityStartTime = DateTime.Now;

            // Start a separate thread for logging
            Thread loggerThread = new Thread(LoggerThreadWork);
            loggerThread.IsBackground = true;
            loggerThread.Start();

            while (!shouldStop)
            {
                string currentWindowTitle = GetActiveWindowTitle();
                string currentProcessName = GetActiveProcessName();

                if (HasActivityChanged(currentWindowTitle, currentProcessName, lastWindowTitle, lastProcessName))
                {
                    LogActivityIfDurationMet(lastWindowTitle, lastProcessName, activityStartTime);

                    lastWindowTitle = currentWindowTitle;
                    lastProcessName = currentProcessName;
                    activityStartTime = DateTime.Now;
                    currentCheckInterval = MIN_CHECK_INTERVAL_MS;
                }
                else
                {
                    currentCheckInterval = Math.Min(currentCheckInterval * 2, MAX_CHECK_INTERVAL_MS);
                }

                Thread.Sleep(currentCheckInterval);
            }

            // Log the final activity if it meets the duration threshold
            LogActivityIfDurationMet(lastWindowTitle, lastProcessName, activityStartTime);

            // Allow time for the logger thread to finish
            Thread.Sleep(1000);
        }

        private void InitializeLogFile()
        {
            if (!File.Exists(LOG_FILE_NAME))
            {
                File.WriteAllText(LOG_FILE_NAME, "StartTimestamp,EndTimestamp,Duration,ProcessName,WindowTitle\n");
            }
        }

        private bool HasActivityChanged(string currentWindowTitle, string currentProcessName,
                                        string lastWindowTitle, string lastProcessName)
        {
            bool processChanged = currentProcessName != lastProcessName;
            bool titleChanged = !recentWindowTitles.Contains(currentWindowTitle);

            if (titleChanged)
            {
                recentWindowTitles.Clear();
                recentWindowTitles.Add(currentWindowTitle);
            }

            return processChanged || titleChanged;
        }

        private void LogActivityIfDurationMet(string windowTitle, string processName, DateTime startTime)
        {
            if (windowTitle != null && processName != null)
            {
                TimeSpan duration = DateTime.Now - startTime;
                if (duration.TotalSeconds >= MIN_DURATION_SECONDS)
                {
                    string logEntry = FormatLogEntry(windowTitle, processName, startTime, duration);
                    logBatch.Add(logEntry);
                    if (logBatch.Count >= LOG_BATCH_SIZE)
                    {
                        logQueue.Enqueue(string.Join(Environment.NewLine, logBatch));
                        logBatch.Clear();
                    }
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
            while (!shouldStop)
            {
                if (logQueue.TryDequeue(out string logEntry))
                {
                    try
                    {
                        RotateLogFileIfNeeded();
                        File.AppendAllText(LOG_FILE_NAME, logEntry + Environment.NewLine);
                    }
                    catch (IOException ex)
                    {
                        EventLog.WriteEntry("WorkWatcher Service", $"Error writing to log file: {ex.Message}", EventLogEntryType.Error);
                        LogError(ex);
                    }
                }
                else
                {
                    Thread.Sleep(1000); // Sleep if queue is empty
                }
            }
        }

        private void RotateLogFileIfNeeded()
        {
            var fileInfo = new FileInfo(LOG_FILE_NAME);
            if (fileInfo.Exists && fileInfo.Length > LOG_ROTATION_SIZE_MB * 1024 * 1024)
            {
                string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                string newFileName = $"activity_log_{timestamp}.csv";
                File.Move(LOG_FILE_NAME, newFileName);
                InitializeLogFile();
            }
        }

        private void LogError(Exception ex)
        {
            try
            {
                string errorLog = $"error_log_{DateTime.Now:yyyyMMdd}.txt";
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