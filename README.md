# Activity Tracker

This console application tracks user activity by logging active window titles and process names.

## Usage

1. Build the project using Visual Studio or the .NET CLI.
2. Run the `ActivityTracker.exe` directly, or use the `run_tracker.bat` to start it minimized.
3. To quit, either press 'Q' in the console window or close the window.

## Auto-start with Windows

To make the tracker start automatically when you log in:

1. Press Win+R, type `shell:startup`, and press Enter.
2. Create a shortcut to `run_tracker.bat` in this folder.

## Output

The activity log is saved as `activity_log.csv` in the same directory as the executable.
Error logs, if any, are saved as `error_log_YYYYMMDD.txt`.

## Note

This application uses Windows API calls and may require elevated permissions to access information about some windows and processes.
