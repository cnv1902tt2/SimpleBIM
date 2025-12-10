using System;
using System.IO;

namespace SimpleBIM.Update
{
    /// <summary>
    /// Logging utility cho update system
    /// </summary>
    public class UpdateLogger
    {
        private static readonly object _lockObject = new object();
        private static UpdateLogger _instance;
        public static UpdateLogger Instance => _instance ?? (_instance = new UpdateLogger());

        private readonly string _logFilePath;
        private readonly bool _loggingEnabled;

        private UpdateLogger()
        {
            var config = UpdateConfiguration.Instance.Settings.Logging;
            _logFilePath = config.LogFilePath;
            _loggingEnabled = UpdateConfiguration.Instance.Settings.Update.LogUpdateActivity;

            EnsureLogDirectory();
            CleanupOldLogs();
        }

        private void EnsureLogDirectory()
        {
            try
            {
                var directory = Path.GetDirectoryName(_logFilePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[UpdateLogger] Error creating log directory: {ex.Message}");
            }
        }

        public void LogInfo(string message)
        {
            Log("INFO", message);
        }

        public void LogWarning(string message)
        {
            Log("WARNING", message);
        }

        public void LogError(string message, Exception ex = null)
        {
            var fullMessage = ex != null ? $"{message}\nException: {ex.Message}\nStackTrace: {ex.StackTrace}" : message;
            Log("ERROR", fullMessage);
        }

        public void LogDebug(string message)
        {
            Log("DEBUG", message);
        }

        private void Log(string level, string message)
        {
            if (!_loggingEnabled) return;

            try
            {
                lock (_lockObject)
                {
                    var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                    var logEntry = $"[{timestamp}] [{level}] {message}";

                    File.AppendAllText(_logFilePath, logEntry + Environment.NewLine);
                    System.Diagnostics.Debug.WriteLine(logEntry);

                    // Check log file size
                    CheckLogFileSize();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[UpdateLogger] Failed to write log: {ex.Message}");
            }
        }

        private void CheckLogFileSize()
        {
            try
            {
                if (!File.Exists(_logFilePath)) return;

                var fileInfo = new FileInfo(_logFilePath);
                var maxSizeMB = UpdateConfiguration.Instance.Settings.Logging.MaxLogFileSizeMB;
                var maxSizeBytes = maxSizeMB * 1024 * 1024;

                if (fileInfo.Length > maxSizeBytes)
                {
                    // Rotate log file
                    var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    var archivePath = _logFilePath.Replace(".log", $"_{timestamp}.log");
                    File.Move(_logFilePath, archivePath);
                    System.Diagnostics.Debug.WriteLine($"[UpdateLogger] Log file rotated to: {archivePath}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[UpdateLogger] Error checking log size: {ex.Message}");
            }
        }

        private void CleanupOldLogs()
        {
            try
            {
                var directory = Path.GetDirectoryName(_logFilePath);
                if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory)) return;

                var retainDays = UpdateConfiguration.Instance.Settings.Logging.RetainLogDays;
                var cutoffDate = DateTime.Now.AddDays(-retainDays);

                var logFiles = Directory.GetFiles(directory, "*.log");
                foreach (var file in logFiles)
                {
                    var fileInfo = new FileInfo(file);
                    if (fileInfo.LastWriteTime < cutoffDate)
                    {
                        File.Delete(file);
                        System.Diagnostics.Debug.WriteLine($"[UpdateLogger] Deleted old log: {file}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[UpdateLogger] Error cleaning up logs: {ex.Message}");
            }
        }

        /// <summary>
        /// Log update activity với structured format
        /// </summary>
        public void LogUpdateActivity(string action, string version, string status, string details = "")
        {
            var message = $"UPDATE_ACTIVITY | Action: {action} | Version: {version} | Status: {status}";
            if (!string.IsNullOrEmpty(details))
            {
                message += $" | Details: {details}";
            }
            LogInfo(message);
        }
    }
}
