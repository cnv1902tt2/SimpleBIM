using Autodesk.Revit.UI;
using SimpleBIM.License;
using SimpleBIM.Update;
using System;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace SimpleBIM
{
    public class App : IExternalApplication
    {
        private static LicenseWindow _licenseWindow;
        private static UIControlledApplication _uiApp;
        private static Dispatcher _mainDispatcher;

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                _uiApp = application;

                // ✅ Lưu dispatcher hiện tại
                _mainDispatcher = Dispatcher.CurrentDispatcher;

                string tabName = "SimpleBIM";
                bool isLicensed = LicenseManager.Instance.ValidateOffline();

                Ribbon.RibbonManager.CreateRibbon(application, tabName, isLicensed);

                if (!isLicensed)
                {
                    Task.Delay(600).ContinueWith(_ =>
                    {
                        _mainDispatcher?.BeginInvoke(new Action(() =>
                        {
                            ShowLicenseWindow();
                        }));
                    });
                }

                // ✅ AUTO-UPDATE CHECK (Async, không block UI)
                _ = Task.Run(() => CheckForUpdatesAsync(application));

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("SimpleBIM Error", ex.ToString());
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application) => Result.Succeeded;

        // Dùng chung cho tất cả panel
        public static void LoadIconToButton(PushButtonData btn, string baseName)
        {
            var large = LoadSingleIcon(baseName + "32") ?? LoadSingleIcon(baseName);
            var small = LoadSingleIcon(baseName + "16") ?? LoadSingleIcon(baseName) ?? large;
            if (large != null) btn.LargeImage = large;
            if (small != null) btn.Image = small;
        }

        public static BitmapImage LoadSingleIcon(string exactNameNoExtension)
        {
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                var resourceName = assembly.GetManifestResourceNames()
                    .FirstOrDefault(r => r.EndsWith("." + exactNameNoExtension + ".png", StringComparison.OrdinalIgnoreCase));
                if (string.IsNullOrEmpty(resourceName)) return null;

                using (var stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream == null) return null;
                    var bmp = new BitmapImage();
                    bmp.BeginInit();
                    bmp.StreamSource = stream;
                    bmp.CacheOption = BitmapCacheOption.OnLoad;
                    bmp.EndInit();
                    bmp.Freeze();
                    return bmp;
                }
            }
            catch { return null; }
        }

        public static void ShowLicenseWindow()
        {
            try
            {
                if (_mainDispatcher == null)
                {
                    _mainDispatcher = Dispatcher.CurrentDispatcher;
                }

                _mainDispatcher.Invoke(() =>
                {
                    if (_licenseWindow == null || !_licenseWindow.IsLoaded)
                    {
                        _licenseWindow = new License.LicenseWindow();
                        _licenseWindow.Show();
                    }
                    else
                    {
                        _licenseWindow.Activate();
                    }
                });
            }
            catch (Exception ex)
            {
                TaskDialog.Show("License Error", "Không thể mở cửa sổ license:\n" + ex.Message);
            }
        }

        public static bool ValidateLicenseForCommand()
        {
            if (!LicenseManager.Instance.ValidateOffline())
            {
                ShowLicenseWindow();
                return false;
            }
            return true;
        }

        public static string GetRevitVersion()
        {
            return _uiApp?.ControlledApplication?.VersionNumber ?? "Unknown";
        }

        /// <summary>
        /// Kiểm tra update async khi startup (không block UI)
        /// </summary>
        private async Task CheckForUpdatesAsync(UIControlledApplication app)
        {
            try
            {
                // Đợi Revit hoàn tất khởi động (8 giây để đảm bảo UI thread sẵn sàng)
                await Task.Delay(8000);

                var versionManager = VersionManager.Instance;
                var updateService = UpdateService.Instance;

                // Kiểm tra có nên check update không (dựa vào cache)
                if (!versionManager.ShouldCheckForUpdates(checkIntervalHours: 24))
                {
                    LogDebug("[App] Skip update check (too soon or version skipped)");
                    return;
                }

                LogDebug("[App] Checking for updates...");

                var revitVersion = GetRevitVersion();
                var result = await updateService.CheckForUpdatesAsync(revitVersion);

                if (result.Success && result.UpdateInfo != null && result.UpdateInfo.UpdateAvailable)
                {
                    LogDebug($"[App] Update available: {result.UpdateInfo.LatestVersion}");

                    // ✅ Hiển thị notification window trên UI thread với retry logic
                    await ShowUpdateNotificationWithRetry(result.UpdateInfo);
                }
                else if (result.Success)
                {
                    LogDebug("[App] No updates available");
                }
                else
                {
                    LogDebug($"[App] Update check failed: {result.ErrorMessage}");
                }
            }
            catch (Exception ex)
            {
                // Silent fail - không làm gián đoạn Revit
                LogDebug($"[App] Update check error: {ex.Message}");
            }
        }

        /// <summary>
        /// Hiển thị notification window với retry logic để đảm bảo thành công
        /// ✅ SIMPLIFIED: No more multiple message boxes, just show the window
        /// </summary>
        private async Task ShowUpdateNotificationWithRetry(UpdateInfo updateInfo, int retryCount = 10)
        {
            for (int attempt = 1; attempt <= retryCount; attempt++)
            {
                try
                {
                    LogDebug($"[App] Attempting to show update window (attempt {attempt}/{retryCount})");

                    // ✅ Lấy dispatcher hiện tại hoặc tạo mới
                    var dispatcher = _mainDispatcher ?? Dispatcher.CurrentDispatcher;

                    if (dispatcher == null)
                    {
                        LogDebug("[App] No dispatcher available, waiting before retry...");
                        await Task.Delay(2000);
                        continue;
                    }

                    // ✅ Tạo và hiển thị window trực tiếp trên dispatcher
                    var windowShown = false;
                    Exception lastException = null;

                    dispatcher.Invoke(() =>
                    {
                        try
                        {
                            var updateWindow = new UpdateNotificationWindow(updateInfo)
                            {
                                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                                Topmost = true,
                                ShowInTaskbar = true,
                                Owner = null // ✅ Không set Owner để window độc lập
                            };

                            updateWindow.Show();
                            updateWindow.Activate();
                            updateWindow.Focus();

                            windowShown = true;
                            LogDebug("[App] ✅ Update notification window shown successfully");
                        }
                        catch (Exception ex)
                        {
                            lastException = ex;
                        }
                    });

                    if (windowShown)
                    {
                        return; // ✅ Thành công, thoát
                    }

                    if (lastException != null)
                    {
                        throw lastException;
                    }
                }
                catch (Exception ex)
                {
                    LogDebug($"[App] Attempt {attempt} failed: {ex.Message}");

                    if (attempt < retryCount)
                    {
                        await Task.Delay(2000);
                    }
                    else
                    {
                        LogDebug($"[App] ❌ Failed to show update window after {retryCount} attempts");

                        // ✅ Fallback: Hiển thị TaskDialog (last resort only)
                        try
                        {
                            TaskDialog.Show(
                                "SimpleBIM - Cập Nhật Mới",
                                $"Phiên bản mới {updateInfo.LatestVersion} đã có sẵn!\n\n" +
                                $"Vui lòng vào menu SimpleBIM > Check for Updates để cập nhật."
                            );
                        }
                        catch { }
                    }
                }
            }
        }

        /// <summary>
        /// Manual check for updates (từ ribbon button)
        /// </summary>
        public static async void ManualCheckForUpdates()
        {
            try
            {
                var versionManager = VersionManager.Instance;
                var updateService = UpdateService.Instance;

                // Force check bỏ qua cache (reset 24h interval)
                versionManager.ForceCheckNow();
                versionManager.ResetSkippedVersion();

                LogDebug("[App] Manual update check initiated by user");

                var revitVersion = GetRevitVersion();
                var result = await updateService.CheckForUpdatesAsync(revitVersion);

                if (result.Success && result.UpdateInfo != null && result.UpdateInfo.UpdateAvailable)
                {
                    var dispatcher = _mainDispatcher ?? Dispatcher.CurrentDispatcher;

                    dispatcher.Invoke(() =>
                    {
                        try
                        {
                            var updateWindow = new UpdateNotificationWindow(result.UpdateInfo)
                            {
                                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                                Topmost = true,
                                ShowInTaskbar = true,
                                Owner = null
                            };
                            updateWindow.Show();
                            updateWindow.Activate();
                            updateWindow.Focus();
                        }
                        catch (Exception ex)
                        {
                            TaskDialog.Show("SimpleBIM Update", $"Lỗi hiển thị thông báo:\n{ex.Message}");
                        }
                    });
                }
                else if (result.Success)
                {
                    TaskDialog.Show("SimpleBIM Update",
                        $"Bạn đang sử dụng phiên bản mới nhất: {versionManager.GetVersionString()}");
                }
                else
                {
                    TaskDialog.Show("SimpleBIM Update",
                        $"Không thể kiểm tra cập nhật:\n{result.ErrorMessage}");
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("SimpleBIM Error", $"Lỗi kiểm tra cập nhật:\n{ex.Message}");
            }
        }

        /// <summary>
        /// Helper method for debug logging
        /// </summary>
        private static void LogDebug(string message)
        {
            System.Diagnostics.Debug.WriteLine(message);

            // ✅ Ghi vào file log
            try
            {
                var logDir = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "SimpleBIM", "Logs");

                if (!System.IO.Directory.Exists(logDir))
                    System.IO.Directory.CreateDirectory(logDir);

                var logFile = System.IO.Path.Combine(logDir, "app_startup.log");
                var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}\n";

                System.IO.File.AppendAllText(logFile, logEntry);
            }
            catch { }
        }
    }
}