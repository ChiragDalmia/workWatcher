using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.IO;
using System.Collections.Concurrent;

class Program
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

  // Thread-safe queue for logging
  private static readonly ConcurrentQueue<string> logQueue = new ConcurrentQueue<string>();

  static void Main(string[] args)
  {
    Console.WriteLine("Activity Tracking Started. Press 'Q' to quit.");

    try
    {
      RunActivityTracker();
    }
    catch (Exception ex)
    {
      Console.WriteLine($"An error occurred: {ex.Message}");
      LogError(ex);
    }
    finally
    {
      Console.WriteLine("Activity Tracking Stopped.");
    }
  }

  static void RunActivityTracker()
  {
    InitializeLogFile();

    string? lastWindowTitle = null;
    string? lastProcessName = null;
    DateTime activityStartTime = DateTime.Now;

    // Start a separate thread for logging
    Thread loggerThread = new Thread(LoggerThreadWork);
    loggerThread.IsBackground = true;
    loggerThread.Start();

    while (!Console.KeyAvailable || Console.ReadKey(true).Key != ConsoleKey.Q)
    {
      string? currentWindowTitle = GetActiveWindowTitle();
      string? currentProcessName = GetActiveProcessName();

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

  static void InitializeLogFile()
  {
    if (!File.Exists(LOG_FILE_NAME))
    {
      File.WriteAllText(LOG_FILE_NAME, "StartTimestamp,EndTimestamp,Duration,ProcessName,WindowTitle\n");
    }
  }

  static bool HasActivityChanged(string? currentWindowTitle, string? currentProcessName,
                                 string? lastWindowTitle, string? lastProcessName)
  {
    return currentWindowTitle != lastWindowTitle || currentProcessName != lastProcessName;
  }

  static void LogActivityIfDurationMet(string? windowTitle, string? processName, DateTime startTime)
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

  static string? GetActiveWindowTitle()
  {
    StringBuilder buff = new StringBuilder(MAX_TITLE_LENGTH);
    IntPtr handle = NativeMethods.GetForegroundWindow();

    if (NativeMethods.GetWindowText(handle, buff, MAX_TITLE_LENGTH) > 0)
    {
      return buff.ToString();
    }
    return null;
  }

  static string? GetActiveProcessName()
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

  static string FormatLogEntry(string windowTitle, string processName, DateTime startTime, TimeSpan duration)
  {
    return $"{startTime:yyyy-MM-dd HH:mm:ss}," +
           $"{startTime.Add(duration):yyyy-MM-dd HH:mm:ss}," +
           $"{duration:hh\\:mm\\:ss}," +
           $"{EscapeCsvField(processName)}," +
           $"{EscapeCsvField(windowTitle)}";
  }

  static string EscapeCsvField(string field)
  {
    if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
    {
      return $"\"{field.Replace("\"", "\"\"")}\"";
    }
    return field;
  }

  static void LoggerThreadWork()
  {
    while (true)
    {
      if (logQueue.TryDequeue(out string? logEntry))
      {
        try
        {
          File.AppendAllText(LOG_FILE_NAME, logEntry + Environment.NewLine);
        }
        catch (IOException ex)
        {
          Console.WriteLine($"Error writing to log file: {ex.Message}");
          LogError(ex);
        }
      }
      else
      {
        Thread.Sleep(1000); // Sleep if queue is empty
      }
    }
  }

  static void LogError(Exception ex)
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