using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Imaging;
using System.Diagnostics;

namespace SimpleBIM.Update
{
    public partial class UpdateNotificationWindow : Window
    {
        private readonly UpdateInfo _updateInfo;
        private readonly UpdateService _updateService;
        private bool _isUpdating = false;
        private string _downloadedUpdatePath = null; // Track downloaded file path


        public UpdateNotificationWindow(UpdateInfo updateInfo)
        {
            InitializeComponent();
            _updateInfo = updateInfo;
            _updateService = UpdateService.Instance;

            // Subscribe to progress events
            _updateService.ProgressChanged += OnUpdateProgress;

            LoadUpdateInfo();
        }



        /// <summary>
        /// ✅ Load icon khi window đã loaded (tránh lỗi XAML parsing)
        /// </summary>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // Thử load icon từ embedded resource
                var icon = App.LoadSingleIcon("license32") ?? App.LoadSingleIcon("license");
                if (icon != null)
                {
                    this.Icon = icon;
                }
                else
                {
                    // Nếu không có icon, dùng icon mặc định của Windows
                    System.Diagnostics.Debug.WriteLine("[UpdateNotificationWindow] No icon found, using default");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[UpdateNotificationWindow] Failed to load icon: {ex.Message}");
            }
        }

        private void LoadUpdateInfo()
        {
            try
            {
                CurrentVersionText.Text = VersionManager.Instance.GetVersionString();
                LatestVersionText.Text = _updateInfo.LatestVersion;
                ReleaseDateText.Text = $"Ngày phát hành: {_updateInfo.ReleaseDate:dd/MM/yyyy}";
                ReleaseNotesText.Text = _updateInfo.ReleaseNotes ?? "Không có thông tin chi tiết.";
                FileSizeText.Text = FormatFileSize(_updateInfo.FileSize);

                // Customize title based on update type
                if (_updateInfo.ForceUpdate || _updateInfo.UpdateType == UpdateType.Mandatory)
                {
                    TitleText.Text = "⚠️ CẬP NHẬT BẮT BUỘC - Vui lòng cập nhật để tiếp tục sử dụng";
                    SkipButton.Visibility = Visibility.Collapsed;
                    RemindLaterButton.Visibility = Visibility.Collapsed;
                }
                else if (_updateInfo.UpdateType == UpdateType.Recommended)
                {
                    TitleText.Text = "Khuyến khích cập nhật để có trải nghiệm tốt nhất";
                }
                else
                {
                    TitleText.Text = "Phiên bản mới của SimpleBIM đã sẵn sàng";
                }

                // Custom notification message
                if (!string.IsNullOrEmpty(_updateInfo.NotificationMessage))
                {
                    TitleText.Text = _updateInfo.NotificationMessage;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[UpdateNotificationWindow] Error loading info: {ex.Message}");
            }
        }

        private async void UpdateNowButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isUpdating) return;

            try
            {
                _isUpdating = true;
                DisableButtons();
                ShowProgressBar();

                // 1. Download
                _downloadedUpdatePath = await _updateService.DownloadUpdateAsync(_updateInfo);

                // 2. Verify
                if (!_updateService.VerifyUpdateIntegrity(_downloadedUpdatePath, _updateInfo.ChecksumSHA256))
                {
                    LogError("Checksum verification failed");

                    // Delete corrupted file
                    DeleteDownloadedFiles();

                    MessageBox.Show(
                        "Xác minh file thất bại. File tải xuống có thể bị hỏng.\nVui lòng thử lại sau.",
                        "Lỗi Xác Minh",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);

                    EnableButtons();
                    HideProgressBar();
                    _isUpdating = false;
                    return;
                }

                // ✅ 3. DOWNLOAD COMPLETE - Switch to post-download UI
                LogInfo("Download and verification completed successfully");
                ShowPostDownloadUI();
            }
            catch (Exception ex)
            {
                LogError($"Error downloading update: {ex.Message}");

                // Clean up on error
                DeleteDownloadedFiles();

                MessageBox.Show(
                    $"Lỗi trong quá trình tải xuống:\n\n{ex.Message}\n\n" +
                    "Vui lòng thử lại hoặc tải xuống installer thủ công.",
                    "Lỗi",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                EnableButtons();
                HideProgressBar();
            }
            finally
            {
                _isUpdating = false;
            }
        }

        /// <summary>
        /// ✅ NEW: Handle "Update Later" button - Delete downloaded files and close
        /// </summary>
        private void UpdateLaterButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LogInfo("User chose to update later - cleaning up downloaded files");

                // Delete downloaded update files
                DeleteDownloadedFiles();

                LogInfo("Update files deleted successfully");

                // Close window without further action
                this.Close();
            }
            catch (Exception ex)
            {
                LogError($"Error during cleanup: {ex.Message}");

                // Still close the window even if cleanup fails
                this.Close();
            }
        }

        /// <summary>
        /// ✅ NEW: Handle "Close Revit & Update" button - Force kill Revit and replace DLL
        /// </summary>
        private async void CloseRevitButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LogInfo("User clicked 'Close Revit & Update' - initiating forced update");

                // Disable buttons to prevent double-click
                DisableButtons();

                // ✅ STEP 1: Create update script that will run AFTER Revit closes
                var updateScript = CreateForceUpdateScript();

                if (string.IsNullOrEmpty(updateScript))
                {
                    throw new Exception("Failed to create update script");
                }

                LogInfo($"Update script created: {updateScript}");

                // ✅ STEP 2: Launch the update script in background
                var scriptLaunched = LaunchUpdateScript(updateScript);

                if (!scriptLaunched)
                {
                    throw new Exception("Failed to launch update script");
                }

                LogInfo("Update script launched successfully");

                // ✅ STEP 3: Give script time to initialize
                await Task.Delay(1000);

                // ✅ STEP 4: Force kill Revit process
                LogInfo("Attempting to force-close Revit...");
                ForceKillRevit();

                // Note: Revit is now killed, this code won't execute
                // The PowerShell script will handle DLL replacement and show final notification
            }
            catch (Exception ex)
            {
                LogError($"Error during forced update: {ex.Message}");

                MessageBox.Show(
                    $"Lỗi khi đóng Revit:\n\n{ex.Message}\n\n" +
                    "Vui lòng đóng Revit thủ công và chạy lại update.",
                    "Lỗi",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                EnableButtons();
            }
        }

        private void RemindLaterButton_Click(object sender, RoutedEventArgs e)
        {
            // Reset last check time để check lại sau
            VersionManager.Instance.ForceCheckNow();
            this.Close();
        }

        private void SkipButton_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show(
                $"Bạn có chắc muốn bỏ qua phiên bản {_updateInfo.LatestVersion}?\n\n" +
                "Bạn sẽ không nhận được thông báo cho phiên bản này nữa.\n" +
                "Bạn vẫn sẽ được thông báo về các phiên bản mới hơn.",
                "Xác Nhận Bỏ Qua",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                VersionManager.Instance.SkipVersion(_updateInfo.LatestVersion);
                this.Close();
            }
        }

        private void OnUpdateProgress(object sender, UpdateProgressEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                ProgressText.Text = e.Message;
                ProgressBar.Value = e.ProgressPercentage;

                if (e.Status == UpdateStatus.Failed || e.Status == UpdateStatus.Cancelled)
                {
                    EnableButtons();
                    HideProgressBar();
                }
            });
        }

        /// <summary>
        /// ✅ NEW: Switch UI to post-download state
        /// </summary>
        private void ShowPostDownloadUI()
        {
            // Hide progress bar
            HideProgressBar();

            // Show post-download notification panel
            PostDownloadPanel.Visibility = Visibility.Visible;

            // Hide pre-download buttons
            PreDownloadButtons.Visibility = Visibility.Collapsed;

            // Show post-download buttons
            PostDownloadButtons.Visibility = Visibility.Visible;

            LogInfo("UI switched to post-download state");
        }

        /// <summary>
        /// ✅ NEW: Delete all downloaded update files
        /// </summary>
        private void DeleteDownloadedFiles()
        {
            try
            {
                if (string.IsNullOrEmpty(_downloadedUpdatePath))
                {
                    LogInfo("No downloaded file to delete");
                    return;
                }

                if (File.Exists(_downloadedUpdatePath))
                {
                    File.Delete(_downloadedUpdatePath);
                    LogInfo($"Deleted downloaded file: {_downloadedUpdatePath}");
                }

                // Also try to delete the temp extract folder if it exists
                var tempFolder = Path.Combine(Path.GetTempPath(), "SimpleBIM_Updates");
                if (Directory.Exists(tempFolder))
                {
                    try
                    {
                        Directory.Delete(tempFolder, true);
                        LogInfo($"Deleted temp folder: {tempFolder}");
                    }
                    catch (Exception ex)
                    {
                        LogError($"Could not delete temp folder: {ex.Message}");
                        // Not critical, continue
                    }
                }

                _downloadedUpdatePath = null;
            }
            catch (Exception ex)
            {
                LogError($"Error deleting downloaded files: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// ✅ NEW: Create PowerShell script for forced update
        /// </summary>
        private string CreateForceUpdateScript()
        {
            try
            {
                var updateFolder = Path.Combine(Path.GetTempPath(), "SimpleBIM_Updates", "ForceUpdate");
                Directory.CreateDirectory(updateFolder);

                var scriptPath = Path.Combine(updateFolder, "force_update.ps1");

                // Extract update package to get new DLL
                var extractPath = Path.Combine(Path.GetTempPath(), "SimpleBIM_Updates", "Extract");
                if (Directory.Exists(extractPath))
                    Directory.Delete(extractPath, true);

                Directory.CreateDirectory(extractPath);
                System.IO.Compression.ZipFile.ExtractToDirectory(_downloadedUpdatePath, extractPath);

                // Find new DLL
                var dllFiles = Directory.GetFiles(extractPath, "SimpleBIM.dll", SearchOption.AllDirectories);
                if (dllFiles.Length == 0)
                {
                    throw new Exception("SimpleBIM.dll not found in update package");
                }

                var newDllPath = dllFiles[0];
                var targetDllPath = GetTargetDllPath();

                // Create PowerShell script
                var script = $@"
# SimpleBIM Force Update Script
$ErrorActionPreference = 'Stop'

# Set window properties
$Host.UI.RawUI.WindowTitle = '🔄 SimpleBIM - Đang Cập Nhật...'
$Host.UI.RawUI.BackgroundColor = 'DarkGreen'
$Host.UI.RawUI.ForegroundColor = 'White'
Clear-Host

Write-Host '═══════════════════════════════════════════════════════════' -ForegroundColor Cyan
Write-Host '           SimpleBIM - AUTOMATIC UPDATE                   ' -ForegroundColor Yellow
Write-Host '═══════════════════════════════════════════════════════════' -ForegroundColor Cyan
Write-Host ''

try {{
    # Step 1: Wait a moment for Revit to fully close
    Write-Host '[1/5] Waiting for Revit to close completely...' -ForegroundColor Cyan
    Start-Sleep -Seconds 3
    
    # Verify Revit is closed
    $revitProcess = Get-Process -Name 'Revit' -ErrorAction SilentlyContinue
    if ($revitProcess) {{
        Write-Host '      ⚠️  Revit process still running, waiting...' -ForegroundColor Yellow
        Start-Sleep -Seconds 2
    }}
    Write-Host '      ✓ Revit closed' -ForegroundColor Green
    Write-Host ''

    # Step 2: Backup old DLL
    Write-Host '[2/5] Creating backup...' -ForegroundColor Cyan
    $targetDll = '{targetDllPath.Replace("\\", "\\\\").Replace("'", "''")}'
    if (Test-Path $targetDll) {{
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $backupPath = $targetDll + '.backup_' + $timestamp
        Copy-Item $targetDll $backupPath -Force
        Write-Host ""      ✓ Backup created: $backupPath"" -ForegroundColor Green
    }}
    Write-Host ''

    # Step 3: Wait for file unlock
    Write-Host '[3/5] Ensuring file is unlocked...' -ForegroundColor Cyan
    $maxRetries = 10
    $unlocked = $false
    
    for ($i = 1; $i -le $maxRetries; $i++) {{
        try {{
            if (Test-Path $targetDll) {{
                $stream = [System.IO.File]::Open($targetDll, 'Open', 'ReadWrite', 'None')
                $stream.Close()
                $stream.Dispose()
            }}
            $unlocked = $true
            Write-Host '      ✓ File unlocked' -ForegroundColor Green
            break
        }} catch {{
            Write-Host ""      ⏳ Retry $i/$maxRetries..."" -ForegroundColor Yellow
            Start-Sleep -Seconds 1
        }}
    }}
    
    if (-not $unlocked) {{
        throw 'File still locked after {0} retries' -f $maxRetries
    }}
    Write-Host ''

    # Step 4: Replace DLL
    Write-Host '[4/5] Installing new version...' -ForegroundColor Cyan
    $newDll = '{newDllPath.Replace("\\", "\\\\").Replace("'", "''")}'
    
    if (Test-Path $targetDll) {{
        Remove-Item $targetDll -Force
    }}
    
    Copy-Item $newDll $targetDll -Force
    
    # Verify installation
    if (Test-Path $targetDll) {{
        $fileInfo = Get-Item $targetDll
        $versionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($targetDll)
        Write-Host ""      ✓ Installed: $($fileInfo.Length) bytes"" -ForegroundColor Green
        Write-Host ""      ✓ Version: $($versionInfo.FileVersion)"" -ForegroundColor Green
    }} else {{
        throw 'Installation failed - target DLL not found'
    }}
    Write-Host ''

    # Step 5: Cleanup
    Write-Host '[5/5] Cleaning up...' -ForegroundColor Cyan
    $extractFolder = '{extractPath.Replace("\\", "\\\\").Replace("'", "''")}'
    if (Test-Path $extractFolder) {{
        Remove-Item $extractFolder -Recurse -Force -ErrorAction SilentlyContinue
    }}
    
    $zipFile = '{_downloadedUpdatePath.Replace("\\", "\\\\").Replace("'", "''")}'
    if (Test-Path $zipFile) {{
        Remove-Item $zipFile -Force -ErrorAction SilentlyContinue
    }}
    Write-Host '      ✓ Temporary files cleaned' -ForegroundColor Green
    Write-Host ''

    # SUCCESS
    Write-Host '═══════════════════════════════════════════════════════════' -ForegroundColor Green
    Write-Host '          ✅ CẬP NHẬT HOÀN TẤT THÀNH CÔNG!               ' -ForegroundColor Green -BackgroundColor Black
    Write-Host '═══════════════════════════════════════════════════════════' -ForegroundColor Green
    Write-Host ''
    Write-Host '🎉 SimpleBIM đã được cập nhật lên phiên bản {_updateInfo.LatestVersion}' -ForegroundColor Cyan
    Write-Host ''
    Write-Host '📌 Đang tự động khởi động lại Revit...' -ForegroundColor Yellow
    Write-Host ''

    # ✅ AUTO-RESTART REVIT
    Start-Sleep -Seconds 2
    
    # Find Revit executable
    $revitPaths = @(
        'C:\Program Files\Autodesk\Revit 2024\Revit.exe',
        'C:\Program Files\Autodesk\Revit 2025\Revit.exe',
        'C:\Program Files\Autodesk\Revit 2023\Revit.exe',
        'C:\Program Files\Autodesk\Revit 2022\Revit.exe'
    )
    
    $revitExe = $null
    foreach ($path in $revitPaths) {{
        if (Test-Path $path) {{
            $revitExe = $path
            break
        }}
    }}
    
    if ($revitExe) {{
        Write-Host ""✓ Tìm thấy Revit: $revitExe"" -ForegroundColor Green
        Write-Host '✓ Đang khởi động Revit...' -ForegroundColor Green
        Start-Process $revitExe
        Write-Host '✓ Revit đã được khởi động!' -ForegroundColor Green
    }} else {{
        Write-Host '⚠️  Không tìm thấy Revit, vui lòng mở thủ công' -ForegroundColor Yellow
    }}
    
    Write-Host ''
    Write-Host 'Cửa sổ này sẽ tự động đóng sau 5 giây...' -ForegroundColor DarkGray
    Write-Host ''

    # Auto-close countdown (reduced to 5 seconds since Revit is starting)
    for ($i = 5; $i -gt 0; $i--) {{
        Write-Host ""`r   Đóng sau $i giây...  "" -NoNewline -ForegroundColor Yellow
        Start-Sleep -Seconds 1
    }}
    
    exit 0

}} catch {{
    Write-Host ''
    Write-Host '═══════════════════════════════════════════════════════════' -ForegroundColor Red
    Write-Host '                    ❌ LỖI CẬP NHẬT                        ' -ForegroundColor Red
    Write-Host '═══════════════════════════════════════════════════════════' -ForegroundColor Red
    Write-Host ''
    Write-Host 'Chi tiết lỗi:' -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Yellow
    Write-Host ''
    Write-Host 'Nhấn phím bất kỳ để đóng...' -ForegroundColor DarkGray
    pause
    exit 1
}}
";

                File.WriteAllText(scriptPath, script, System.Text.Encoding.UTF8);
                LogInfo($"Force update script created: {scriptPath}");

                return scriptPath;
            }
            catch (Exception ex)
            {
                LogError($"Error creating force update script: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// ✅ NEW: Launch update script in background
        /// </summary>
        private bool LaunchUpdateScript(string scriptPath)
        {
            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-ExecutionPolicy Bypass -NoProfile -WindowStyle Normal -File \"{scriptPath}\"",
                    UseShellExecute = true,
                    CreateNoWindow = false
                };

                var process = Process.Start(startInfo);

                if (process != null)
                {
                    LogInfo($"Update script launched with PID: {process.Id}");
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                LogError($"Error launching update script: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// ✅ NEW: Force kill all Revit processes
        /// </summary>
        private void ForceKillRevit()
        {
            try
            {
                var scriptContent = @"
Get-Process -Name 'Revit' -ErrorAction SilentlyContinue | Stop-Process -Force
";

                var scriptPath = Path.Combine(Path.GetTempPath(), "kill_revit.ps1");
                File.WriteAllText(scriptPath, scriptContent);

                var startInfo = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File \"{scriptPath}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                var process = Process.Start(startInfo);
                process?.WaitForExit(5000); // Wait max 5 seconds

                LogInfo("Revit force-kill command executed");

                // Note: This line may not execute if Revit kills this process too
            }
            catch (Exception ex)
            {
                LogError($"Error killing Revit: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Get target DLL path
        /// </summary>
        private string GetTargetDllPath()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                "SimpleBIM", "Revit Addins", "SimpleBIM.dll");
        }

        private void ShowProgressBar()
        {
            ProgressPanel.Visibility = Visibility.Visible;
        }

        private void HideProgressBar()
        {
            ProgressPanel.Visibility = Visibility.Collapsed;
        }

        private void DisableButtons()
        {
            UpdateNowButton.IsEnabled = false;
            RemindLaterButton.IsEnabled = false;
            SkipButton.IsEnabled = false;
        }

        private void EnableButtons()
        {
            UpdateNowButton.IsEnabled = true;
            RemindLaterButton.IsEnabled = true;
            SkipButton.IsEnabled = true;
        }

        private string FormatFileSize(long bytes)
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

        /// <summary>
        /// ✅ NEW: Log info message
        /// </summary>
        private void LogInfo(string message)
        {
            System.Diagnostics.Debug.WriteLine($"[UpdateNotificationWindow] {message}");

            try
            {
                var logDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "SimpleBIM", "Logs");

                if (!Directory.Exists(logDir))
                    Directory.CreateDirectory(logDir);

                var logFile = Path.Combine(logDir, "update_window.log");
                var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] INFO: {message}\n";

                File.AppendAllText(logFile, logEntry);
            }
            catch { }
        }

        /// <summary>
        /// ✅ NEW: Log error message
        /// </summary>
        private void LogError(string message)
        {
            System.Diagnostics.Debug.WriteLine($"[UpdateNotificationWindow] ERROR: {message}");

            try
            {
                var logDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "SimpleBIM", "Logs");

                if (!Directory.Exists(logDir))
                    Directory.CreateDirectory(logDir);

                var logFile = Path.Combine(logDir, "update_window.log");
                var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ERROR: {message}\n";

                File.AppendAllText(logFile, logEntry);
            }
            catch { }
        }

        protected override void OnClosed(EventArgs e)
        {
            // Unsubscribe from events
            _updateService.ProgressChanged -= OnUpdateProgress;
            base.OnClosed(e);
        }

    }
}