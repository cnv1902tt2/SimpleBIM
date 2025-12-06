using System;
using System.IO;
using System.Linq;
using System.Reflection;
using IWshRuntimeLibrary;

namespace SimpleBIM.Installer
{
    internal class Program
    {
        private const string AddinName = "SimpleBIM.AS.tab";
        private const string CompanyName = "SimpleBIM";
        private static readonly string[] RevitVersions = { "2025", "2024", "2023", "2022", "2021", "2020", "2019", "2018" };

        static void Main(string[] args)
        {
            Console.Title = "SimpleBIM AS Tab - Installer";
            Console.OutputEncoding = System.Text.Encoding.UTF8;

            try
            {
                if (args.Length > 0)
                {
                    ProcessCommandLine(args);
                }
                else
                {
                    ShowMainMenu();
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"\n✗ LỖI: {ex.Message}");
                Console.ResetColor();
                Console.WriteLine("\nNhấn phím bất kỳ để thoát...");
                Console.ReadKey();
            }
        }

        static void ProcessCommandLine(string[] args)
        {
            if (args.Contains("/install") || args.Contains("-i"))
            {
                Install(true);
            }
            else if (args.Contains("/uninstall") || args.Contains("-u"))
            {
                Uninstall(true);
            }
            else if (args.Contains("/silent") || args.Contains("-s"))
            {
                InstallSilent();
            }
            else
            {
                Console.WriteLine("Tham số không hợp lệ!");
                Console.WriteLine("Sử dụng: /install, /uninstall, hoặc /silent");
                Environment.Exit(1);
            }
        }

        static void ShowMainMenu()
        {
            while (true)
            {
                Console.Clear();
                Console.WriteLine("╔══════════════════════════════════════════════════╗");
                Console.WriteLine("║          SIMPLEBIM AS TAB - INSTALLER           ║");
                Console.WriteLine("╠══════════════════════════════════════════════════╣");
                Console.WriteLine("║                                                  ║");
                Console.WriteLine("║   1. 📥 CÀI ĐẶT ADD-IN                          ║");
                Console.WriteLine("║   2. 🗑️  GỠ CÀI ĐẶT                             ║");
                Console.WriteLine("║   3. ℹ️  THÔNG TIN CÀI ĐẶT                      ║");
                Console.WriteLine("║   4. 🚪 THOÁT                                   ║");
                Console.WriteLine("║                                                  ║");
                Console.WriteLine("╚══════════════════════════════════════════════════╝");
                Console.Write("\nChọn tùy chọn [1-4]: ");

                var key = Console.ReadKey();
                Console.WriteLine();

                switch (key.KeyChar)
                {
                    case '1':
                        Install(false);
                        break;
                    case '2':
                        Uninstall(false);
                        break;
                    case '3':
                        ShowInstallInfo();
                        break;
                    case '4':
                        Environment.Exit(0);
                        break;
                    default:
                        Console.WriteLine("Lựa chọn không hợp lệ!");
                        Pause();
                        break;
                }
            }
        }

        static void Install(bool silentMode)
        {
            try
            {
                if (!silentMode) Console.Clear();
                Console.WriteLine("🔄 Đang cài đặt SimpleBIM AS Tab...\n");

                // Lấy đường dẫn thư mục hiện tại (nơi chứa installer)
                string currentDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string addinSource = Path.Combine(currentDir, "SimpleBIM.AS.tab.addin");
                string dllSource = Path.Combine(currentDir, "SimpleBIM.AS.tab.dll");

                // === THÊM PHẦN NÀY: TẠO FILE .ADDIN NẾU CHƯA CÓ ===
                if (!System.IO.File.Exists(addinSource))
                {
                    Console.WriteLine("📝 Tự động tạo file cấu hình add-in...");
                    CreateInitialAddinFile(addinSource);
                    Console.WriteLine($"✓ Đã tạo file: {addinSource}");
                }

                // Kiểm tra file tồn tại
                if (!System.IO.File.Exists(addinSource) || !System.IO.File.Exists(dllSource))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("❌ KHÔNG TÌM THẤY FILE ADD-IN!");
                    Console.ResetColor();
                    Console.WriteLine("\nVui lòng đặt các file sau cùng thư mục với installer:");
                    Console.WriteLine("✓ SimpleBIM.AS.tab.addin");
                    Console.WriteLine("✓ SimpleBIM.AS.tab.dll");

                    if (!silentMode) Pause();
                    return;
                }

                // Tạo thư mục chung cho DLL
                string commonPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    CompanyName,
                    "Revit Addins"
                );

                Directory.CreateDirectory(commonPath);
                Console.WriteLine($"✓ Tạo thư mục chung: {commonPath}");

                // Copy DLL
                string dllDest = Path.Combine(commonPath, "SimpleBIM.AS.tab.dll");
                System.IO.File.Copy(dllSource, dllDest, true);
                Console.WriteLine($"✓ Sao chép DLL: {dllDest}");

                // Cài đặt cho các phiên bản Revit
                int installedCount = 0;
                foreach (string version in RevitVersions)
                {
                    string revitAddinFolder = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                        "Autodesk",
                        "Revit",
                        "Addins",
                        version
                    );

                    if (Directory.Exists(revitAddinFolder))
                    {
                        CreateAddinFile(revitAddinFolder, dllDest);
                        installedCount++;
                        Console.WriteLine($"✓ Cài đặt cho Revit {version}");
                    }
                }

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("\n══════════════════════════════════════════════════");
                Console.WriteLine("            ✅ CÀI ĐẶT THÀNH CÔNG!");
                Console.WriteLine("══════════════════════════════════════════════════");
                Console.ResetColor();
                Console.WriteLine($"\n• Đã cài đặt cho {installedCount} phiên bản Revit");
                Console.WriteLine("• Khởi động lại Revit để sử dụng add-in");
                Console.WriteLine("• Shortcut gỡ cài đặt đã được tạo trên Desktop");

                // Ghi registry để hỗ trợ gỡ cài đặt qua Control Panel
                WriteRegistryInfo(currentDir);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"\n❌ LỖI TRONG QUÁ TRÌNH CÀI ĐẶT:");
                Console.WriteLine($"   {ex.Message}");
                Console.ResetColor();
            }

            if (!silentMode) Pause();
        }
        static void CreateInitialAddinFile(string filePath)
        {
            // Lấy GUID từ DLL hoặc dùng mặc định
            string guid = ExtractGuidFromDll() ?? "{12345678-1234-1234-1234-123456789ABC}";

            string content = $@"<?xml version=""1.0"" encoding=""utf-8"" standalone=""no""?>
            <RevitAddIns>
              <AddIn Type=""Application"">
                <Name>SimpleBIM AS Tab</Name>
                <Assembly>REPLACE_WITH_INSTALL_PATH</Assembly>
                <AddInId>{guid}</AddInId>
                <FullClassName>SimpleBIM.AS.tab.App</FullClassName>
                <VendorId>SIMPLEBIM</VendorId>
                <VendorDescription>SimpleBIM Add-in for Revit</VendorDescription>
              </AddIn>
            </RevitAddIns>";

            System.IO.File.WriteAllText(filePath, content, System.Text.Encoding.UTF8);
        }

        static string ExtractGuidFromDll()
        {
            try
            {
                string currentDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string dllPath = Path.Combine(currentDir, "SimpleBIM.AS.tab.dll");

                if (System.IO.File.Exists(dllPath))
                {
                    var assembly = Assembly.LoadFrom(dllPath);
                    var guidAttr = assembly.GetCustomAttribute<System.Runtime.InteropServices.GuidAttribute>();
                    if (guidAttr?.Value != null)
                    {
                        return "{" + guidAttr.Value.ToUpper() + "}";
                    }
                }
            }
            catch { }
            return null;
        }
        static void InstallSilent()
        {
            try
            {
                Install(true);
                Environment.Exit(0);
            }
            catch
            {
                Environment.Exit(1);
            }
        }

        static string GetAddinGuid()
        {
            string currentDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string addinFilePath = Path.Combine(currentDir, "SimpleBIM.AS.tab.addin");

            if (System.IO.File.Exists(addinFilePath))
            {
                try
                {
                    string content = System.IO.File.ReadAllText(addinFilePath);

                    // Tìm GUID trong file
                    int start = content.IndexOf("<AddInId>");
                    int end = content.IndexOf("</AddInId>");

                    if (start > 0 && end > start)
                    {
                        string guid = content.Substring(start + 9, end - start - 9).Trim();
                        return guid;
                    }
                }
                catch { }
            }

            // Fallback: GUID mặc định
            return "{12345678-1234-1234-1234-123456789ABC}";
        }
        static void CreateAddinFile(string targetFolder, string dllPath)
        {
            string addinPath = Path.Combine(targetFolder, "SimpleBIM.AS.tab.addin");
            string guidString = GetAddinGuid();
            // Đọc nội dung file .addin gốc
            string addinContent = @"<?xml version=""1.0"" encoding=""utf-8"" standalone=""no""?>
            <RevitAddIns>
              <AddIn Type=""Application"">
                <Name>SimpleBIM AS Tab</Name>
                <Assembly>" + dllPath.Replace("\\", "\\\\") + @"</Assembly>
                <AddInId>" + guidString + @"</AddInId>
                <FullClassName>SimpleBIM.AS.tab.App</FullClassName>
                <VendorId>SIMPLEBIM</VendorId>
                <VendorDescription>SimpleBIM Add-in for Revit</VendorDescription>
              </AddIn>
            </RevitAddIns>";

            System.IO.File.WriteAllText(addinPath, addinContent, System.Text.Encoding.UTF8);
        }

        static void Uninstall(bool silentMode)
        {
            try
            {
                if (!silentMode) Console.Clear();
                Console.WriteLine("🔄 Đang gỡ cài đặt SimpleBIM AS Tab...\n");

                int removedCount = 0;

                // Xóa file .addin từ tất cả Revit versions
                foreach (string version in RevitVersions)
                {
                    string addinPath = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                        "Autodesk",
                        "Revit",
                        "Addins",
                        version,
                        "SimpleBIM.AS.tab.addin"
                    );

                    if (System.IO.File.Exists(addinPath))
                    {
                        System.IO.File.Delete(addinPath);
                        removedCount++;
                        Console.WriteLine($"✓ Đã xóa: Revit {version}");
                    }
                }

                // Xóa thư mục chứa DLL
                string commonPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    CompanyName
                );

                if (Directory.Exists(commonPath))
                {
                    Directory.Delete(commonPath, true);
                    Console.WriteLine($"✓ Đã xóa thư mục: {commonPath}");
                }
                // 3. XÓA HOÀN TOÀN LICENSE CŨ (đây là phần bạn cần!)
                string licenseFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "SimpleBIM"
                );
                if (Directory.Exists(licenseFolder))
                {
                    try
                    {
                        Directory.Delete(licenseFolder, true);
                        Console.WriteLine($"Đã xóa dữ liệu license cũ: {licenseFolder}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Không thể xóa license folder (có thể đang dùng): {ex.Message}");
                    }
                }
                // Xóa shortcut
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string shortcutPath = Path.Combine(desktopPath, "Uninstall SimpleBIM.lnk");
                if (System.IO.File.Exists(shortcutPath))
                {
                    System.IO.File.Delete(shortcutPath);
                    Console.WriteLine("✓ Đã xóa shortcut trên Desktop");
                }

                // Xóa registry
                DeleteRegistryInfo();

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("\n══════════════════════════════════════════════════");
                Console.WriteLine("            ✅ GỠ CÀI ĐẶT THÀNH CÔNG!");
                Console.WriteLine("══════════════════════════════════════════════════");
                Console.ResetColor();
                Console.WriteLine($"\n• Đã xóa từ {removedCount} phiên bản Revit");
                Console.WriteLine("• Add-in đã được gỡ hoàn toàn");
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"\n❌ LỖI: {ex.Message}");
                Console.ResetColor();
            }

            if (!silentMode) Pause();
        }

        static void ShowInstallInfo()
        {
            Console.Clear();
            Console.WriteLine("╔══════════════════════════════════════════════════╗");
            Console.WriteLine("║            THÔNG TIN CÀI ĐẶT                     ║");
            Console.WriteLine("╠══════════════════════════════════════════════════╣\n");

            Console.WriteLine("📂 Vị trí file DLL:");
            string commonPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                CompanyName,
                "Revit Addins",
                "SimpleBIM.AS.tab.dll"
            );
            Console.WriteLine($"   {commonPath}");
            Console.WriteLine($"   Tồn tại: {(System.IO.File.Exists(commonPath) ? "✓ CÓ" : "✗ KHÔNG")}");

            Console.WriteLine("\n📁 Các phiên bản Revit đã cài đặt:");
            int count = 0;
            foreach (string version in RevitVersions)
            {
                string addinPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "Autodesk",
                    "Revit",
                    "Addins",
                    version,
                    "SimpleBIM.AS.tab.addin"
                );

                if (System.IO.File.Exists(addinPath))
                {
                    Console.WriteLine($"   ✓ Revit {version}");
                    count++;
                }
            }

            if (count == 0)
            {
                Console.WriteLine("   ✗ Chưa cài đặt cho phiên bản nào");
            }

            Console.WriteLine($"\n📊 Tổng số phiên bản đã cài: {count}/{RevitVersions.Length}");
            Pause();
        }

        static void WriteRegistryInfo(string installPath)
        {
            try
            {
                Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(
                    @"Software\SimpleBIM\AS Tab"
                );
                key.SetValue("InstallPath", installPath);
                key.SetValue("InstallDate", DateTime.Now.ToString("yyyy-MM-dd"));
                key.SetValue("Version", "1.0.0");
                key.Close();
            }
            catch { }
        }

        static void DeleteRegistryInfo()
        {
            try
            {
                Microsoft.Win32.Registry.CurrentUser.DeleteSubKeyTree(@"Software\SimpleBIM", false);
            }
            catch { }
        }

        static void Pause()
        {
            Console.WriteLine("\n──────────────────────────────────────────");
            Console.Write("Nhấn phím bất kỳ để tiếp tục...");
            Console.ReadKey();
        }
    }
}