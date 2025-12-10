using System;
using System.IO;
using System.Reflection;
using Newtonsoft.Json;
using System.Windows;

namespace SimpleBIM.Update
{
    /// <summary>
    /// Quản lý version của add-in, so sánh semantic versioning, và cache version check
    /// </summary>
    public class VersionManager
    {
        private static readonly string CacheFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "SimpleBIM", "version_cache.json");

        private static VersionManager _instance;
        public static VersionManager Instance => _instance ?? (_instance = new VersionManager());

        public Version CurrentVersion { get; private set; }
        public VersionCache Cache { get; private set; }

        private VersionManager()
        {
            LoadCurrentVersion();
            LoadCache();
        }

        /// <summary>
        /// Đọc version hiện tại từ Assembly
        /// </summary>
        private void LoadCurrentVersion()
        {
            try
            {
                // Lấy đúng đường dẫn DLL mà Revit đang dùng (từ installer)
                string dllPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    "SimpleBIM", "Revit Addins", "SimpleBIM.dll");

                if (!File.Exists(dllPath))
                {
                    // Fallback: dùng executing assembly (chỉ khi chạy debug trực tiếp)
                    dllPath = Assembly.GetExecutingAssembly().Location;
                }

                var assemblyName = AssemblyName.GetAssemblyName(dllPath);
                CurrentVersion = assemblyName.Version;

                System.Diagnostics.Debug.WriteLine($"\n========== VERSION MANAGER DEBUG ==========");
                System.Diagnostics.Debug.WriteLine($"[VersionManager] Đang đọc version từ: {dllPath}");
                System.Diagnostics.Debug.WriteLine($"[VersionManager] Version thực tế: {CurrentVersion}");
                System.Diagnostics.Debug.WriteLine($"========================================\n");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[VersionManager] Không đọc được version từ file DLL: {ex.Message}");
                CurrentVersion = new Version("1.0.0");
            }
        }

        /// <summary>
        /// Đọc cache từ disk
        /// </summary>
        private void LoadCache()
        {
            try
            {
                if (File.Exists(CacheFilePath))
                {
                    var json = File.ReadAllText(CacheFilePath);
                    Cache = JsonConvert.DeserializeObject<VersionCache>(json);
                }
                else
                {
                    Cache = new VersionCache();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[VersionManager] Error loading cache: {ex.Message}");
                Cache = new VersionCache();
            }
        }

        /// <summary>
        /// Lưu cache ra disk
        /// </summary>
        public void SaveCache()
        {
            try
            {
                var directory = Path.GetDirectoryName(CacheFilePath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var json = JsonConvert.SerializeObject(Cache, Formatting.Indented);
                File.WriteAllText(CacheFilePath, json);
                System.Diagnostics.Debug.WriteLine($"[VersionManager] Cache saved");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[VersionManager] Error saving cache: {ex.Message}");
            }
        }

        /// <summary>
        /// So sánh 2 version theo semantic versioning
        /// </summary>
        /// <returns>-1 nếu v1 < v2, 0 nếu bằng nhau, 1 nếu v1 > v2</returns>
        public int CompareVersions(string v1, string v2)
        {
            try
            {
                var version1 = ParseVersion(v1);
                var version2 = ParseVersion(v2);
                return version1.CompareTo(version2);
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// Parse string thành Version object
        /// </summary>
        private Version ParseVersion(string versionString)
        {
            versionString = versionString.Trim().TrimStart('v', 'V');

            if (Version.TryParse(versionString, out Version version))
            {
                return version;
            }

            // Fallback: thêm .0 nếu thiếu
            if (!versionString.Contains("."))
            {
                versionString += ".0.0.0";
            }
            else
            {
                var parts = versionString.Split('.');
                while (parts.Length < 4)
                {
                    versionString += ".0";
                    parts = versionString.Split('.');
                }
            }

            return Version.TryParse(versionString, out version) ? version : new Version("1.0.0.0");
        }

        /// <summary>
        /// Kiểm tra có cần check update không (dựa vào cache)
        /// </summary>
        /// <param name="checkIntervalHours">Interval giữa các lần check (default: 24h)</param>
        public bool ShouldCheckForUpdates(int checkIntervalHours = 24)
        {
            // Nếu user đã skip version này
            if (!string.IsNullOrEmpty(Cache.SkippedVersion) &&
                Cache.SkippedVersion == CurrentVersion.ToString())
            {
                return false;
            }

            // Kiểm tra thời gian check gần nhất
            //var timeSinceLastCheck = DateTime.Now - Cache.LastCheckTime;
            //return timeSinceLastCheck.TotalHours >= checkIntervalHours;
            return true;
        }

        /// <summary>
        /// Update cache sau khi check version
        /// </summary>
        public void UpdateCache(string latestVersion, bool updateAvailable)
        {
            Cache.LastCheckTime = DateTime.Now;
            Cache.LatestKnownVersion = latestVersion;
            Cache.UpdateAvailable = updateAvailable;
            SaveCache();
        }

        /// <summary>
        /// User skip version này
        /// </summary>
        public void SkipVersion(string version)
        {
            Cache.SkippedVersion = version;
            SaveCache();
        }

        /// <summary>
        /// Reset skip version (khi user click "Check for updates" manually)
        /// </summary>
        public void ResetSkippedVersion()
        {
            Cache.SkippedVersion = null;
            SaveCache();
        }

        /// <summary>
        /// Force check bỏ qua cache
        /// </summary>
        public void ForceCheckNow()
        {
            Cache.LastCheckTime = DateTime.MinValue;
            SaveCache();
        }

        /// <summary>
        /// Kiểm tra version có hợp lệ không
        /// </summary>
        public bool IsValidVersion(string versionString)
        {
            if (string.IsNullOrWhiteSpace(versionString))
                return false;

            try
            {
                ParseVersion(versionString);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Get version string for display
        /// </summary>
        public string GetVersionString(bool includeRevision = false)
        {
            if (includeRevision)
            {
                return CurrentVersion.ToString();
            }
            else
            {
                return $"{CurrentVersion.Major}.{CurrentVersion.Minor}.{CurrentVersion.Build}";
            }
        }
    }

    /// <summary>
    /// Cache data structure
    /// </summary>
    public class VersionCache
    {
        public DateTime LastCheckTime { get; set; } = DateTime.MinValue;
        public string LatestKnownVersion { get; set; }
        public bool UpdateAvailable { get; set; }
        public string SkippedVersion { get; set; }
        public int CheckCount { get; set; }
        public DateTime LastUpdateDate { get; set; }
    }
}
