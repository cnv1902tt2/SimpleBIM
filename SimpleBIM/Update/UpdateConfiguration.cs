using Newtonsoft.Json;
using System;
using System.IO;

namespace SimpleBIM.Update
{
    /// <summary>
    /// Configuration manager cho update system
    /// </summary>
    public class UpdateConfiguration
    {
        private static readonly string ConfigPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "SimpleBIM", "update_config.json");

        private static UpdateConfiguration _instance;
        public static UpdateConfiguration Instance => _instance ?? (_instance = new UpdateConfiguration());

        public UpdateSettings Settings { get; private set; }

        private UpdateConfiguration()
        {
            LoadConfiguration();
        }

        private void LoadConfiguration()
        {
            try
            {
                if (File.Exists(ConfigPath))
                {
                    var json = File.ReadAllText(ConfigPath);
                    Settings = JsonConvert.DeserializeObject<UpdateSettings>(json);
                }
                else
                {
                    // Load default settings
                    Settings = UpdateSettings.Default;
                    SaveConfiguration();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[UpdateConfiguration] Error loading: {ex.Message}");
                Settings = UpdateSettings.Default;
            }
        }

        public void SaveConfiguration()
        {
            try
            {
                var directory = Path.GetDirectoryName(ConfigPath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var json = JsonConvert.SerializeObject(Settings, Formatting.Indented);
                File.WriteAllText(ConfigPath, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[UpdateConfiguration] Error saving: {ex.Message}");
            }
        }

        public void ReloadConfiguration()
        {
            LoadConfiguration();
        }
    }

    /// <summary>
    /// Update settings data structure
    /// </summary>
    public class UpdateSettings
    {
        public UpdateConfig Update { get; set; }
        public LoggingConfig Logging { get; set; }
        public AdvancedConfig Advanced { get; set; }

        public static UpdateSettings Default => new UpdateSettings
        {
            Update = new UpdateConfig
            {
                ApiBaseUrl = "https://apikeymanagement.onrender.com/updates/check",
                CheckIntervalHours = 24,
                EnableAutoUpdate = true,
                EnableAutoDownload = false,
                DownloadOnWiFiOnly = false,
                ShowNotificationOnStartup = true,
                LogUpdateActivity = true
            },
            Logging = new LoggingConfig
            {
                LogLevel = "Information",
                LogFilePath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "SimpleBIM", "Logs", "update.log"),
                MaxLogFileSizeMB = 10,
                RetainLogDays = 30
            },
            Advanced = new AdvancedConfig
            {
                ProxyEnabled = false,
                ProxyAddress = "",
                ProxyPort = 0,
                TimeoutSeconds = 300,
                MaxRetries = 3,
                BackupRetentionDays = 7
            }
        };
    }

    public class UpdateConfig
    {
        public string ApiBaseUrl { get; set; }
        public int CheckIntervalHours { get; set; }
        public bool EnableAutoUpdate { get; set; }
        public bool EnableAutoDownload { get; set; }
        public bool DownloadOnWiFiOnly { get; set; }
        public bool ShowNotificationOnStartup { get; set; }
        public bool LogUpdateActivity { get; set; }
    }

    public class LoggingConfig
    {
        public string LogLevel { get; set; }
        public string LogFilePath { get; set; }
        public int MaxLogFileSizeMB { get; set; }
        public int RetainLogDays { get; set; }
    }

    public class AdvancedConfig
    {
        public bool ProxyEnabled { get; set; }
        public string ProxyAddress { get; set; }
        public int ProxyPort { get; set; }
        public int TimeoutSeconds { get; set; }
        public int MaxRetries { get; set; }
        public int BackupRetentionDays { get; set; }
    }
}
