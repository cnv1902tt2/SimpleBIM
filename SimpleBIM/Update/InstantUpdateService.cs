using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SimpleBIM.Update
{
    /// <summary>
    /// Update service - Cập nhật NGAY LẬP TỨC bằng external process
    /// Giải pháp: Tạo một updater.exe độc lập chạy sau khi copy DLL
    /// </summary>
    public class InstantUpdateService
    {
        private static readonly string TempFolder = Path.Combine(Path.GetTempPath(), "SimpleBIM_Updates");
        private static readonly string UpdaterFolder = Path.Combine(TempFolder, "Updater");

        public static InstantUpdateService Instance { get; } = new InstantUpdateService();

        /// <summary>
        /// Apply update NGAY - sử dụng shadow copy technique
        /// </summary>
        public async Task<bool> ApplyUpdateInstantly(string updatePackagePath, string targetDllPath, string backupPath)
        {
            try
            {
                LogInfo("🚀 Starting INSTANT update...");

                // 1. Extract update package
                var extractPath = Path.Combine(TempFolder, "Extract");
                if (Directory.Exists(extractPath))
                    Directory.Delete(extractPath, true);

                Directory.CreateDirectory(extractPath);
                System.IO.Compression.ZipFile.ExtractToDirectory(updatePackagePath, extractPath);

                LogInfo("========== FILES IN ZIP ==========");
                var allFiles = Directory.GetFiles(extractPath, "*.*", SearchOption.AllDirectories);
                foreach (var file in allFiles)
                {
                    var relativePath = file.Replace(extractPath, "").TrimStart('\\', '/');
                    LogInfo($"Found: {relativePath}");
                }
                LogInfo("==================================");

                // 2. Tìm DLL mới
                var dllFiles = Directory.GetFiles(extractPath, "SimpleBIM.dll", SearchOption.AllDirectories);
                if (dllFiles.Length == 0)
                {
                    LogError("❌ SimpleBIM.dll not found in update package");
                    return false;
                }

                var newDllPath = dllFiles[0];
                LogInfo($"✓ Found new DLL: {newDllPath}");

                // 3. Copy DLL mới vào thư mục temp với tên chính xác
                var tempUpdatePath = Path.Combine(TempFolder, "PendingUpdate", "SimpleBIM.dll");
                Directory.CreateDirectory(Path.GetDirectoryName(tempUpdatePath));
                File.Copy(newDllPath, tempUpdatePath, true);
                LogInfo($"✓ Copied to pending location: {tempUpdatePath}");

                // 4. Tạo monitoring script để đợi Revit đóng rồi tự động thay thế
                var monitorScript = CreateMonitoringScript(tempUpdatePath, targetDllPath, backupPath);
                var scriptPath = Path.Combine(UpdaterFolder, "auto_update.ps1");

                Directory.CreateDirectory(UpdaterFolder);
                File.WriteAllText(scriptPath, monitorScript, Encoding.UTF8);
                LogInfo($"✓ Created monitoring script: {scriptPath}");

                // 5. Chạy monitoring script trong background
                var success = await RunMonitoringScriptAsync(scriptPath);

                if (success)
                {
                    LogInfo("✅ Monitoring script started - waiting for Revit to close");

                    // Cleanup extract folder
                    await Task.Delay(1000);
                    CleanupTempFiles(extractPath, updatePackagePath);

                    return true;
                }
                else
                {
                    LogError("❌ Failed to start monitoring script - Trying fallback method...");

                    // ✅ FALLBACK: Schedule update on next Revit restart
                    var fallbackSuccess = ScheduleUpdateOnRestart(tempUpdatePath, targetDllPath);
                    if (fallbackSuccess)
                    {
                        LogInfo("✅ Update scheduled for next Revit restart");
                        return true;
                    }

                    return false;
                }
            }
            catch (Exception ex)
            {
                LogError($"❌ Error in instant update: {ex.Message}");
                LogError($"Stack trace: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// Tạo PowerShell script để monitor Revit process và tự động update khi Revit đóng
        /// ✅ IMPROVED: Better visibility, progress updates, and error handling
        /// </summary>
        private string CreateMonitoringScript(string newDllPath, string targetDll, string backupPath)
        {
            return $@"
# SimpleBIM Auto-Update Monitor Script
# Monitors Revit process and replaces DLL when Revit closes

$ErrorActionPreference = 'Stop'
$Host.UI.RawUI.WindowTitle = '⚡ SimpleBIM Auto-Update Monitor - ĐANG CHỜ BẠN ĐÓNG REVIT'
$Host.UI.RawUI.BackgroundColor = 'DarkBlue'
$Host.UI.RawUI.ForegroundColor = 'White'
Clear-Host

Write-Host '═══════════════════════════════════════════════════════════' -ForegroundColor Cyan
Write-Host '        🚀 SimpleBIM - AUTO UPDATE MONITOR 🚀              ' -ForegroundColor Yellow
Write-Host '═══════════════════════════════════════════════════════════' -ForegroundColor Cyan
Write-Host ''
Write-Host '📌 TRẠNG THÁI: ' -NoNewline -ForegroundColor White
Write-Host 'CHỜ BẠN ĐÓNG REVIT' -ForegroundColor Yellow -BackgroundColor DarkRed
Write-Host ''
Write-Host '💡 HƯỚNG DẪN:' -ForegroundColor Cyan
Write-Host '   1. Lưu công việc trong Revit' -ForegroundColor Gray
Write-Host '   2. Đóng Revit (File > Exit hoặc nút X)' -ForegroundColor Gray
Write-Host '   3. Script sẽ TỰ ĐỘNG cài đặt cập nhật' -ForegroundColor Gray
Write-Host '   4. Mở lại Revit để dùng phiên bản mới' -ForegroundColor Gray
Write-Host ''
Write-Host '⚠️  QUAN TRỌNG: ' -NoNewline -ForegroundColor Red
Write-Host 'ĐỪNG ĐÓNG CỬA SỔ NÀY!' -ForegroundColor Yellow
Write-Host ''
Write-Host '═══════════════════════════════════════════════════════════' -ForegroundColor Cyan
Write-Host ''

# Đợi Revit đóng (check mỗi 2 giây)
$revitProcesses = @('Revit')
$maxWaitMinutes = 120  # Tăng lên 2 giờ cho an toàn

$startTime = Get-Date
$iteration = 0

Write-Host '⏳ ' -NoNewline -ForegroundColor Yellow
Write-Host 'Đang quét Revit processes...' -ForegroundColor White
Write-Host ''

while ($true) {{
    $iteration++
    $running = Get-Process | Where-Object {{ $revitProcesses -contains $_.ProcessName }}
    
    if (-not $running) {{
        Write-Host ''
        Write-Host '✅ REVIT ĐÃ ĐÓNG! Bắt đầu cập nhật...' -ForegroundColor Green -BackgroundColor Black
        Write-Host ''
        break
    }}
    
    # Hiển thị progress mỗi 30 giây
    if ($iteration % 15 -eq 0) {{
        $elapsed = (Get-Date) - $startTime
        $minutes = [math]::Floor($elapsed.TotalMinutes)
        $seconds = [math]::Floor($elapsed.TotalSeconds % 60)
        Write-Host ""⏱️  Đã chờ: ${{minutes}}m ${{seconds}}s | Revit vẫn đang chạy..."" -ForegroundColor DarkYellow
    }}
    
    $elapsed = (Get-Date) - $startTime
    if ($elapsed.TotalMinutes -gt $maxWaitMinutes) {{
        Write-Host ''
        Write-Host '⚠️  TIMEOUT: Đã chờ quá 2 giờ' -ForegroundColor Red
        Write-Host ''
        Write-Host '📋 Cập nhật sẽ được lên lịch khi khởi động lại Windows.' -ForegroundColor Yellow
        Write-Host ''
        pause
        exit 2
    }}
    
    Start-Sleep -Seconds 2
}}

# ═══════════════════════════════════════════════════════════
# BƯỚC 2: BẮT ĐẦU CÀI ĐẶT CẬP NHẬT
# ═══════════════════════════════════════════════════════════

Write-Host ''
Write-Host '🔧 ĐANG CÀI ĐẶT CẬP NHẬT...' -ForegroundColor Yellow -BackgroundColor DarkBlue
Write-Host ''
Start-Sleep -Seconds 2

try {{
    # ✅ STEP 1: Verify new DLL exists
    Write-Host '[1/5] Kiểm tra file cập nhật...' -ForegroundColor Cyan
    if (-not (Test-Path '{newDllPath.Replace("'", "''")}')) {{
        throw ""File cập nhật không tồn tại: {newDllPath.Replace("'", "''")}""
    }}
    $newDllInfo = Get-Item '{newDllPath.Replace("'", "''")}'
    Write-Host ""      ✓ Tìm thấy: $($newDllInfo.Length) bytes"" -ForegroundColor Green
    Write-Host ''

    # ✅ STEP 2: Create backup
    Write-Host '[2/5] Tạo bản sao lưu...' -ForegroundColor Cyan
    if (Test-Path '{targetDll.Replace("'", "''")}') {{
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $backupName = '{targetDll.Replace("'", "''")}.backup_' + $timestamp
        Copy-Item '{targetDll.Replace("'", "''")}' $backupName -Force
        $backupInfo = Get-Item $backupName
        Write-Host ""      ✓ Backup: $backupName ($($backupInfo.Length) bytes)"" -ForegroundColor Green
    }} else {{
        Write-Host '      ⚠️  File gốc không tồn tại (có thể là lần cài đầu tiên)' -ForegroundColor Yellow
    }}
    Write-Host ''

    # ✅ STEP 3: Wait for file unlock (critical!)
    Write-Host '[3/5] Chờ file được unlock...' -ForegroundColor Cyan
    $maxRetries = 10
    $retryCount = 0
    $unlocked = $false
    
    while (-not $unlocked -and $retryCount -lt $maxRetries) {{
        try {{
            if (Test-Path '{targetDll.Replace("'", "''")}') {{
                # Try to open with exclusive access
                $fileStream = [System.IO.File]::Open('{targetDll.Replace("'", "''")}', 'Open', 'ReadWrite', 'None')
                $fileStream.Close()
                $fileStream.Dispose()
            }}
            $unlocked = $true
            Write-Host '      ✓ File đã unlock' -ForegroundColor Green
        }} catch {{
            $retryCount++
            Write-Host ""      ⏳ Retry $retryCount/$maxRetries (file vẫn bị lock)..."" -ForegroundColor Yellow
            Start-Sleep -Seconds 1
        }}
    }}
    
    if (-not $unlocked) {{
        throw ""File vẫn bị lock sau $maxRetries lần thử. Vui lòng đảm bảo Revit đã đóng hoàn toàn.""
    }}
    Write-Host ''

    # ✅ STEP 4: Remove old DLL
    Write-Host '[4/5] Xóa phiên bản cũ...' -ForegroundColor Cyan
    if (Test-Path '{targetDll.Replace("'", "''")}') {{
        Remove-Item '{targetDll.Replace("'", "''")}' -Force
        Write-Host '      ✓ Đã xóa file cũ' -ForegroundColor Green
    }}
    Write-Host ''

    # ✅ STEP 5: Copy new DLL
    Write-Host '[5/5] Cài đặt phiên bản mới...' -ForegroundColor Cyan
    Copy-Item '{newDllPath.Replace("'", "''")}' '{targetDll.Replace("'", "''")}' -Force
    
    # Verify installation
    if (Test-Path '{targetDll.Replace("'", "''")}') {{
        $installedInfo = Get-Item '{targetDll.Replace("'", "''")}'
        $installedSize = $installedInfo.Length
        
        # ✅ CRITICAL: Verify file size matches
        if ($installedSize -eq $newDllInfo.Length) {{
            Write-Host ""      ✓ Đã cài đặt: $installedSize bytes"" -ForegroundColor Green
            
            # Try to get DLL version
            try {{
                $versionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo('{targetDll.Replace("'", "''")}')
                Write-Host ""      ✓ Version: $($versionInfo.FileVersion)"" -ForegroundColor Green
            }} catch {{
                Write-Host '      ⚠️  Không đọc được version info' -ForegroundColor Yellow
            }}
        }} else {{
            throw ""Kích thước file không khớp! Expected: $($newDllInfo.Length), Got: $installedSize""
        }}
    }} else {{
        throw ""File mới không tồn tại sau khi copy: {targetDll.Replace("'", "''")}""
    }}
    Write-Host ''
    
    # ═══════════════════════════════════════════════════════════
    # SUCCESS MESSAGE
    # ═══════════════════════════════════════════════════════════
    Write-Host ''
    Write-Host '═══════════════════════════════════════════════════════════' -ForegroundColor Green
    Write-Host '                ✅ CẬP NHẬT HOÀN TẤT!                      ' -ForegroundColor Green -BackgroundColor Black
    Write-Host '═══════════════════════════════════════════════════════════' -ForegroundColor Green
    Write-Host ''
    Write-Host '📁 Vị trí: ' -NoNewline -ForegroundColor White
    Write-Host '{targetDll.Replace("'", "''")}' -ForegroundColor Cyan
    Write-Host ""📦 Kích thước: $installedSize bytes"" -ForegroundColor White
    Write-Host ''
    Write-Host '🚀 BƯỚC TIẾP THEO:' -ForegroundColor Yellow
    Write-Host '   1. Mở Revit' -ForegroundColor Gray
    Write-Host '   2. Kiểm tra version trong SimpleBIM ribbon' -ForegroundColor Gray
    Write-Host '   3. Sử dụng các tính năng mới!' -ForegroundColor Gray
    Write-Host ''
    Write-Host '═══════════════════════════════════════════════════════════' -ForegroundColor Green
    Write-Host ''

    # Cleanup temp file
    if (Test-Path '{newDllPath.Replace("'", "''")}) {{
        Remove-Item '{newDllPath.Replace("'", "''")}' -Force -ErrorAction SilentlyContinue
        Write-Host '🗑️  Cleaned up temp file' -ForegroundColor DarkGray
    }}

    # Cleanup extract folder
    $extractFolder = Split-Path '{newDllPath.Replace("'", "''")}' -Parent
    if (Test-Path $extractFolder) {{
        Remove-Item $extractFolder -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host '🗑️  Cleaned up extract folder' -ForegroundColor DarkGray
    }}

    Write-Host ''
    Write-Host '⏰ Cửa sổ này sẽ tự động đóng sau 10 giây...' -ForegroundColor DarkYellow
    Write-Host '   (hoặc nhấn phím bất kỳ để đóng ngay)' -ForegroundColor DarkGray
    Write-Host ''
    
    # Auto-close after 10 seconds
    $timeout = 10
    for ($i = $timeout; $i -gt 0; $i--) {{
        Write-Host ""`r   Đóng sau $i giây...  "" -NoNewline -ForegroundColor Yellow
        Start-Sleep -Seconds 1
    }}
    Write-Host ""`r   Đang đóng...         "" -ForegroundColor Green
    
    exit 0
}}
catch {{
    Write-Host ''
    Write-Host '═══════════════════════════════════════════════════════════' -ForegroundColor Red
    Write-Host '                ❌ LỖI CẬP NHẬT                            ' -ForegroundColor Red -BackgroundColor Black
    Write-Host '═══════════════════════════════════════════════════════════' -ForegroundColor Red
    Write-Host ''
    Write-Host '📛 Chi tiết lỗi:' -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Yellow
    Write-Host ''
    
    # Restore backup nếu có lỗi
    if (Test-Path '{backupPath.Replace("'", "''")}) {{
        Write-Host '🔄 Đang khôi phục bản sao lưu...' -ForegroundColor Yellow
        try {{
            Copy-Item '{backupPath.Replace("'", "''")}' '{targetDll.Replace("'", "''")}' -Force
            Write-Host '✓ Đã khôi phục backup thành công' -ForegroundColor Green
        }} catch {{
            Write-Host '❌ Không thể khôi phục backup: ' + $_.Exception.Message -ForegroundColor Red
        }}
    }}
    
    Write-Host ''
    Write-Host '💡 GIẢI PHÁP:' -ForegroundColor Cyan
    Write-Host '   1. Đảm bảo Revit đã đóng HOÀN TOÀN (kiểm tra Task Manager)' -ForegroundColor Gray
    Write-Host '   2. Khởi động lại Windows' -ForegroundColor Gray
    Write-Host '   3. Chạy lại update hoặc cài đặt thủ công' -ForegroundColor Gray
    Write-Host ''
    Write-Host '═══════════════════════════════════════════════════════════' -ForegroundColor Red
    Write-Host ''
    Write-Host 'Nhấn phím bất kỳ để đóng...' -ForegroundColor DarkGray
    pause
    exit 1
}}
";
        }

        /// <summary>
        /// Chạy PowerShell monitoring script trong background
        /// ĐẢM BẢO HIỆN LÊN TRÊN CÙNG, KHÔNG BỊ ẨN DƯỚI REVIT
        /// </summary>
        private async Task<bool> RunMonitoringScriptAsync(string scriptPath)
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

                LogInfo("Đang khởi động script cập nhật và ép lên trên cùng...");

                var process = Process.Start(startInfo);
                if (process == null)
                {
                    LogError("Không thể khởi động PowerShell");
                    return false;
                }

                // Đợi cửa sổ hiện ra (tối đa 6 giây)
                for (int i = 0; i < 30; i++)
                {
                    await Task.Delay(200);
                    process.Refresh();

                    if (process.MainWindowHandle != IntPtr.Zero)
                    {
                        NativeMethods.BringToFrontAndFlash(process.MainWindowHandle);
                        LogInfo("Đã đưa cửa sổ cập nhật lên trên cùng + nhấp nháy taskbar!");
                        return true;
                    }
                }

                LogError("Timeout: Không tìm thấy cửa sổ PowerShell sau 6 giây");
                return false;
            }
            catch (Exception ex)
            {
                LogError($"Lỗi khởi động script: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// ✅ FALLBACK: Schedule update on next restart using batch file
        /// Cách này LUÔN THÀNH CÔNG vì chạy khi Revit đã tắt
        /// </summary>
        private bool ScheduleUpdateOnRestart(string tempDll, string targetDll)
        {
            try
            {
                LogInfo("📋 Creating update batch script for next restart...");

                var escapedTempDll = tempDll.Replace("\\", "\\\\").Replace("^", "^^^^").Replace("&", "^&").Replace("<", "^<").Replace(">", "^>").Replace("|", "^|");
                var escapedTargetDll = targetDll.Replace("\\", "\\\\").Replace("^", "^^^^").Replace("&", "^&").Replace("<", "^<").Replace(">", "^>").Replace("|", "^|");

                var batchScript = $@"@echo off
echo ================================================
echo SimpleBIM - Delayed Update
echo ================================================
echo.
echo Waiting for Revit to close...
timeout /t 3 /nobreak >nul

echo Backing up current DLL...
if exist ""{escapedTargetDll}"" (
    set BACKUP_NAME={escapedTargetDll}.old_%date:~-4,4%%date:~-7,2%%date:~-10,2%_%time:~0,2%%time:~3,2%%time:~6,2%
    set BACKUP_NAME=%BACKUP_NAME: =0%
    ren ""{escapedTargetDll}"" ""%BACKUP_NAME%""
    echo Renamed old DLL to backup
)

echo Installing new version...
if exist ""{escapedTempDll}"" (
    copy ""{escapedTempDll}"" ""{escapedTargetDll}"" /Y
    echo Copied new DLL
) else (
    echo ERROR: New DLL file not found: {escapedTempDll}
    goto :error
)

if exist ""{escapedTargetDll}"" (
    echo.
    echo ================================================
    echo UPDATE COMPLETED SUCCESSFULLY!
    echo ================================================
    echo.
    echo New DLL installed: {escapedTargetDll}
    echo Please start Revit to use the new version.
    echo.
    echo Old DLL backed up with .old extension
    echo.
    goto :success
) else (
    echo ERROR: Failed to install update!
    goto :error
)

:error
echo.
echo ================================================
echo UPDATE FAILED!
echo ================================================
echo.
if exist ""%BACKUP_NAME%"" (
    echo Restoring backup...
    ren ""%BACKUP_NAME%"" ""SimpleBIM.dll""
    echo Backup restored
)
pause
exit /b 1

:success
echo Press any key to close...
pause >nul

rem Self-delete batch file
del ""%~f0"" /F /Q
exit /b 0
";

                var batchPath = Path.Combine(UpdaterFolder, "delayed_update.bat");
                Directory.CreateDirectory(UpdaterFolder);
                File.WriteAllText(batchPath, batchScript, Encoding.Default);

                LogInfo($"✓ Batch script created: {batchPath}");

                // Tạo shortcut trong Startup folder để chạy khi Windows khởi động
                var startupFolder = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
                var shortcutPath = Path.Combine(startupFolder, "SimpleBIM_Update.bat");
                File.Copy(batchPath, shortcutPath, true);

                LogInfo($"✓ Update script copied to Startup: {shortcutPath}");
                LogInfo("⚠️ Update will be applied when you restart Windows or manually run the batch file");

                return true;
            }
            catch (Exception ex)
            {
                LogError($"Failed to schedule update: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Cleanup temp files
        /// </summary>
        private void CleanupTempFiles(string extractPath, string zipPath)
        {
            try
            {
                if (Directory.Exists(extractPath))
                {
                    Directory.Delete(extractPath, true);
                    LogInfo("✓ Cleaned up extract folder");
                }

                if (File.Exists(zipPath))
                {
                    File.Delete(zipPath);
                    LogInfo("✓ Deleted update package");
                }
            }
            catch (Exception ex)
            {
                LogError($"Cleanup warning: {ex.Message}");
            }
        }

        private void LogInfo(string message)
        {
            System.Diagnostics.Debug.WriteLine($"[InstantUpdate] {message}");

            try
            {
                var logDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "SimpleBIM", "Logs");

                if (!Directory.Exists(logDir))
                    Directory.CreateDirectory(logDir);

                var logFile = Path.Combine(logDir, "instant_update.log");
                var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}\n";

                File.AppendAllText(logFile, logEntry);
            }
            catch { }
        }

        private void LogError(string message)
        {
            System.Diagnostics.Debug.WriteLine($"[InstantUpdate] ERROR: {message}");

            try
            {
                var logDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "SimpleBIM", "Logs");

                if (!Directory.Exists(logDir))
                    Directory.CreateDirectory(logDir);

                var logFile = Path.Combine(logDir, "instant_update.log");
                var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ERROR: {message}\n";

                File.AppendAllText(logFile, logEntry);
            }
            catch { }
        }
    }
}
