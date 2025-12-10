using Newtonsoft.Json;
using SimpleBIM.License;
using System;
using System.IO;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace SimpleBIM.Update
{
    /// <summary>
    /// Service xử lý kiểm tra và thực hiện update
    /// </summary>
    public class UpdateService
    {
        private static readonly string ApiBaseUrl = "https://apikeymanagement.onrender.com/updates/check";
        private static readonly string TempFolder = Path.Combine(Path.GetTempPath(), "SimpleBIM_Updates");
        private static readonly string BackupFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "SimpleBIM", "Backups");
        private static readonly HttpClient httpClient = new HttpClient { Timeout = TimeSpan.FromMinutes(10) };

        public event EventHandler<UpdateProgressEventArgs> ProgressChanged;

        private static UpdateService _instance;
        public static UpdateService Instance => _instance ?? (_instance = new UpdateService());

        private UpdateService()
        {
            EnsureFoldersExist();
        }

        private void EnsureFoldersExist()
        {
            try
            {
                if (!Directory.Exists(TempFolder))
                    Directory.CreateDirectory(TempFolder);
                if (!Directory.Exists(BackupFolder))
                    Directory.CreateDirectory(BackupFolder);
            }
            catch (Exception ex)
            {
                LogError($"Error creating folders: {ex.Message}");
            }
        }

        /// <summary>
        /// Kiểm tra có update mới không
        /// </summary>
        public async Task<UpdateCheckResult> CheckForUpdatesAsync(string revitVersion = "Unknown")
        {
            try
            {
                ReportProgress(UpdateStatus.Checking, 0, "Đang kiểm tra phiên bản mới...");

                var versionManager = VersionManager.Instance;
                var licenseManager = LicenseManager.Instance;

                var request = new UpdateCheckRequest
                {
                    Product = "SimpleBIM",
                    CurrentVersion = versionManager.CurrentVersion.ToString(),
                    RevitVersion = revitVersion,
                    MachineHash = licenseManager.GetMachineHash(),
                    OS = Environment.OSVersion.VersionString
                };

                var jsonContent = JsonConvert.SerializeObject(request);
                var httpContent = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                // ✅ DEBUG: Log chi tiết
                LogInfo($"\n========== UPDATE CHECK DEBUG ==========");
                LogInfo($"📦 Product: {request.Product}");
                LogInfo($"🔢 Current Version (From Assembly): {request.CurrentVersion}");
                LogInfo($"💻 Revit Version: {request.RevitVersion}");
                LogInfo($"📕 Machine Hash: {request.MachineHash}");
                LogInfo($"🔧 OS: {request.OS}");
                LogInfo($"🌐 API Endpoint: {ApiBaseUrl}");
                LogInfo($"📋 Full JSON Request:\n{jsonContent}");
                LogInfo($"======================================\n");

                LogInfo($"Checking for updates... Current version: {request.CurrentVersion}");

                var response = await httpClient.PostAsync(ApiBaseUrl, httpContent);

                if (!response.IsSuccessStatusCode)
                {
                    return new UpdateCheckResult
                    {
                        Success = false,
                        ErrorMessage = $"HTTP Error: {response.StatusCode}"
                    };
                }

                var responseString = await response.Content.ReadAsStringAsync();

                // ✅ DEBUG: Log API response
                LogInfo($"\n========== API RESPONSE ==========");
                LogInfo($"Status Code: {response.StatusCode}");
                LogInfo($"Response JSON:\n{responseString}");
                LogInfo($"================================\n");

                var updateResponse = JsonConvert.DeserializeObject<UpdateCheckResponse>(responseString);

                var updateInfo = new UpdateInfo
                {
                    UpdateAvailable = updateResponse.UpdateAvailable,
                    LatestVersion = updateResponse.LatestVersion,
                    MinimumRequiredVersion = updateResponse.MinimumRequiredVersion,
                    ReleaseDate = DateTime.Parse(updateResponse.ReleaseDate),
                    ReleaseNotes = updateResponse.ReleaseNotes,
                    DownloadUrl = updateResponse.DownloadUrl,
                    FileSize = updateResponse.FileSize,
                    ChecksumSHA256 = updateResponse.ChecksumSHA256,
                    UpdateType = ParseUpdateType(updateResponse.UpdateType),
                    ForceUpdate = updateResponse.ForceUpdate,
                    NotificationMessage = updateResponse.NotificationMessage
                };

                // Update cache
                versionManager.UpdateCache(updateInfo.LatestVersion, updateInfo.UpdateAvailable);

                // ✅ DEBUG: Log comparison result
                LogInfo($"\n========== VERSION COMPARISON ==========");
                LogInfo($"Current Version Sent: {request.CurrentVersion}");
                LogInfo($"Latest Version Received: {updateInfo.LatestVersion}");
                LogInfo($"Update Available: {updateInfo.UpdateAvailable}");
                LogInfo($"Force Update: {updateInfo.ForceUpdate}");
                LogInfo($"Update Type: {updateInfo.UpdateType}");
                LogInfo($"Download URL: {updateInfo.DownloadUrl}");
                LogInfo($"Notification Message: {updateInfo.NotificationMessage}");
                LogInfo($"========================================\n");

                LogInfo($"Update check completed. Available: {updateInfo.UpdateAvailable}, Latest: {updateInfo.LatestVersion}");

                ReportProgress(UpdateStatus.Idle, 100, "Kiểm tra hoàn tất");

                return new UpdateCheckResult
                {
                    Success = true,
                    UpdateInfo = updateInfo
                };
            }
            catch (HttpRequestException ex)
            {
                LogError($"Network error: {ex.Message}");
                return new UpdateCheckResult
                {
                    Success = false,
                    ErrorMessage = "Không thể kết nối đến server. Vui lòng kiểm tra kết nối internet."
                };
            }
            catch (Exception ex)
            {
                LogError($"Error checking for updates: {ex.Message}");
                return new UpdateCheckResult
                {
                    Success = false,
                    ErrorMessage = $"Lỗi: {ex.Message}"
                };
            }
        }

        /// <summary>
        /// Download update package
        /// </summary>
        public async Task<string> DownloadUpdateAsync(UpdateInfo updateInfo)
        {
            try
            {
                ReportProgress(UpdateStatus.Downloading, 0, "Đang tải xuống update...");

                var fileName = $"SimpleBIM_v{updateInfo.LatestVersion}.zip";
                var downloadPath = Path.Combine(TempFolder, fileName);

                // Xóa file cũ nếu có
                if (File.Exists(downloadPath))
                    File.Delete(downloadPath);

                LogInfo($"Downloading update from: {updateInfo.DownloadUrl}");

                using (var response = await httpClient.GetAsync(updateInfo.DownloadUrl, HttpCompletionOption.ResponseHeadersRead))
                {
                    response.EnsureSuccessStatusCode();

                    var totalBytes = response.Content.Headers.ContentLength ?? updateInfo.FileSize;
                    var buffer = new byte[8192];
                    var bytesRead = 0L;

                    using (var contentStream = await response.Content.ReadAsStreamAsync())
                    using (var fileStream = new FileStream(downloadPath, FileMode.Create, FileAccess.Write, FileShare.None, 8192, true))
                    {
                        int read;
                        while ((read = await contentStream.ReadAsync(buffer, 0, buffer.Length)) > 0)
                        {
                            await fileStream.WriteAsync(buffer, 0, read);
                            bytesRead += read;

                            var progress = totalBytes > 0 ? (int)((bytesRead * 100) / totalBytes) : 0;
                            ReportProgress(UpdateStatus.Downloading, progress,
                                $"Đang tải xuống... {FormatBytes(bytesRead)} / {FormatBytes(totalBytes)}",
                                bytesRead, totalBytes);
                        }
                    }
                }

                LogInfo($"Download completed: {downloadPath}");
                return downloadPath;
            }
            catch (Exception ex)
            {
                LogError($"Error downloading update: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Verify checksum của file download
        /// </summary>
        public bool VerifyUpdateIntegrity(string filePath, string expectedChecksum)
        {
            try
            {
                ReportProgress(UpdateStatus.Verifying, 50, "Đang xác minh tính toàn vẹn của file...");

                if (!File.Exists(filePath))
                {
                    LogError("File không tồn tại để verify");
                    return false;
                }

                using (var sha256 = SHA256.Create())
                using (var stream = File.OpenRead(filePath))
                {
                    var hash = sha256.ComputeHash(stream);
                    var hashString = BitConverter.ToString(hash).Replace("-", "").ToLower();

                    LogInfo($"Computed hash: {hashString}");
                    LogInfo($"Expected hash: {expectedChecksum}");

                    var verified = hashString.Equals(expectedChecksum, StringComparison.OrdinalIgnoreCase);

                    ReportProgress(UpdateStatus.Verifying, 100,
                        verified ? "Xác minh thành công" : "Xác minh thất bại");

                    return verified;
                }
            }
            catch (Exception ex)
            {
                LogError($"Error verifying integrity: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Apply update - sử dụng InstantUpdateService
        /// </summary>
        public async Task<bool> ApplyUpdateAsync(string updatePackagePath)
        {
            try
            {
                ReportProgress(UpdateStatus.Installing, 10, "Đang chuẩn bị cài đặt...");

                // 1. Backup current version
                var currentDllPath = GetCurrentDllPath();
                if (string.IsNullOrEmpty(currentDllPath) || !File.Exists(currentDllPath))
                {
                    LogError($"Current DLL not found: {currentDllPath}");
                    return false;
                }

                var version = VersionManager.Instance.CurrentVersion.ToString();
                var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var backupFileName = $"SimpleBIM_{version}_{timestamp}.dll";
                var backupPath = Path.Combine(BackupFolder, backupFileName);

                // Tạo backup
                Directory.CreateDirectory(BackupFolder);
                File.Copy(currentDllPath, backupPath, true);
                LogInfo($"Backup created: {backupPath}");

                ReportProgress(UpdateStatus.Installing, 30, "Đang cài đặt...");

                // 2. Sử dụng InstantUpdateService để update NGAY
                var instantUpdater = InstantUpdateService.Instance;
                var success = await instantUpdater.ApplyUpdateInstantly(
                    updatePackagePath,
                    currentDllPath,
                    backupPath
                );

                if (success)
                {
                    ReportProgress(UpdateStatus.Completed, 100, "Cài đặt hoàn tất! Vui lòng khởi động lại Revit.");
                    LogInfo("✅ Update completed successfully");

                    // Cleanup old backups
                    CleanupOldBackups();

                    return true;
                }
                else
                {
                    ReportProgress(UpdateStatus.Failed, 0, "Cài đặt thất bại");
                    LogError("Update failed");
                    return false;
                }
            }
            catch (Exception ex)
            {
                LogError($"Error applying update: {ex.Message}");
                ReportProgress(UpdateStatus.Failed, 0, $"Lỗi: {ex.Message}");
                return false;
            }
        }

        // ✅ Helper: Kiểm tra file có đang được sử dụng không
        private bool IsFileInUse(string filePath)
        {
            try
            {
                using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    return false;
                }
            }
            catch (IOException)
            {
                return true;
            }
        }

        /// <summary>
        /// Rollback về version trước
        /// </summary>
        public bool RollbackToPreviousVersion()
        {
            try
            {
                var backupFiles = Directory.GetFiles(BackupFolder, "SimpleBIM_*.dll");
                if (backupFiles.Length == 0)
                {
                    LogError("No backup found for rollback");
                    return false;
                }

                // Lấy backup mới nhất
                Array.Sort(backupFiles);
                var latestBackup = backupFiles[backupFiles.Length - 1];

                var currentDllPath = GetCurrentDllPath();
                if (!string.IsNullOrEmpty(currentDllPath))
                {
                    File.Copy(latestBackup, currentDllPath, true);
                    LogInfo($"Rolled back to: {latestBackup}");
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                LogError($"Error rolling back: {ex.Message}");
                return false;
            }
        }

        #region Helper Methods

        private string GetCurrentDllPath()
        {
            // Đường dẫn DLL theo installer convention
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                "SimpleBIM", "Revit Addins", "SimpleBIM.dll");
        }

        private void BackupCurrentVersion(string dllPath)
        {
            try
            {
                var version = VersionManager.Instance.CurrentVersion.ToString();
                var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var backupFileName = $"SimpleBIM_{version}_{timestamp}.dll";
                var backupPath = Path.Combine(BackupFolder, backupFileName);

                File.Copy(dllPath, backupPath, true);
                LogInfo($"Backup created: {backupPath}");
            }
            catch (Exception ex)
            {
                LogError($"Error creating backup: {ex.Message}");
            }
        }

        private void CleanupOldBackups()
        {
            try
            {
                var backupFiles = Directory.GetFiles(BackupFolder, "SimpleBIM_*.dll");
                if (backupFiles.Length <= 5) return; // Giữ 5 backups gần nhất

                Array.Sort(backupFiles);
                for (int i = 0; i < backupFiles.Length - 5; i++)
                {
                    File.Delete(backupFiles[i]);
                    LogInfo($"Deleted old backup: {backupFiles[i]}");
                }
            }
            catch (Exception ex)
            {
                LogError($"Error cleaning up backups: {ex.Message}");
            }
        }

        private UpdateType ParseUpdateType(string type)
        {
            if (Enum.TryParse<UpdateType>(type, true, out var result))
                return result;
            return UpdateType.Optional;
        }

        private string FormatBytes(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB" };
            double len = bytes;
            int order = 0;
            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len = len / 1024;
            }
            return $"{len:0.##} {sizes[order]}";
        }

        private void ReportProgress(UpdateStatus status, int percentage, string message,
            long bytesDownloaded = 0, long totalBytes = 0)
        {
            ProgressChanged?.Invoke(this, new UpdateProgressEventArgs
            {
                Status = status,
                ProgressPercentage = percentage,
                Message = message,
                BytesDownloaded = bytesDownloaded,
                TotalBytes = totalBytes
            });
        }

        private void LogInfo(string message)
        {
            System.Diagnostics.Debug.WriteLine($"[UpdateService] {message}");

            // ✅ Ghi vào file log
            try
            {
                var logDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "SimpleBIM", "Logs");

                if (!Directory.Exists(logDir))
                    Directory.CreateDirectory(logDir);

                var logFile = Path.Combine(logDir, "update_debug.log");
                var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}\n";

                File.AppendAllText(logFile, logEntry);
            }
            catch { }
        }

        private void LogError(string message)
        {
            System.Diagnostics.Debug.WriteLine($"[UpdateService] ERROR: {message}");

            // ✅ Ghi vào file log
            try
            {
                var logDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "SimpleBIM", "Logs");

                if (!Directory.Exists(logDir))
                    Directory.CreateDirectory(logDir);

                var logFile = Path.Combine(logDir, "update_debug.log");
                var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ERROR: {message}\n";

                File.AppendAllText(logFile, logEntry);
            }
            catch { }
        }

        #endregion
    }
}