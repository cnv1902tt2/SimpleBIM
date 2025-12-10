using System;
using System.IO;
using System.Linq;
using System.Reflection;

namespace SimpleBIM.Installer
{
    internal class Program
    {
        private const string AddinName = "SimpleBIM";
        private const string CompanyName = "SimpleBIM";
        private static readonly string[] RevitVersions = { "2025", "2024", "2023", "2022", "2021", "2020", "2019", "2018" };

        static void Main(string[] args)
        {
            Console.Title = "SimpleBIM - Installer";
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
                Console.WriteLine("║          SIMPLEBIM - INSTALLER                  ║");
                Console.WriteLine("╠══════════════════════════════════════════════════╣");
                Console.WriteLine("║                                                  ║");
                Console.WriteLine("║   1. 📥 CÀI ĐẶT ADD-IN                          ║");
                Console.WriteLine("║   2. 🗑️  GỠ CÀI ĐẶT                             ║");
                Console.WriteLine("║   3. ℹ️  THÔNG TIN CÀI ĐẶT                      ║");
                Console.WriteLine("║   4. 🔄 TẠO GUID MỚI                            ║");
                Console.WriteLine("║   5. 🚪 THOÁT                                   ║");
                Console.WriteLine("║                                                  ║");
                Console.WriteLine("╚══════════════════════════════════════════════════╝");
                Console.Write("\nChọn tùy chọn [1-5]: ");

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
                        RegenerateGuid();
                        break;
                    case '5':
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
                Console.WriteLine("🔄 Đang cài đặt SimpleBIM...\n");

                // Lấy đường dẫn thư mục hiện tại (nơi chứa installer)
                string currentDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string addinSource = Path.Combine(currentDir, "SimpleBIM.addin");
                string dllSource = Path.Combine(currentDir, "SimpleBIM.dll");

                // Tạo file .addin nếu chưa có hoặc cần tạo mới
                if (!System.IO.File.Exists(addinSource))
                {
                    Console.WriteLine("📝 Tự động tạo file cấu hình add-in với GUID mới...");
                    CreateInitialAddinFile(addinSource, true);
                    Console.WriteLine($"✓ Đã tạo file: {addinSource}");
                }

                // Kiểm tra file tồn tại
                if (!System.IO.File.Exists(dllSource))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("❌ KHÔNG TÌM THẤY FILE SimpleBIM.dll!");
                    Console.ResetColor();
                    Console.WriteLine("\nVui lòng đặt SimpleBIM.dll cùng thư mục với installer.");

                    if (!silentMode) Pause();
                    return;
                }

                // Tạo thư mục chung cho DLL và .addin
                string commonPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    CompanyName,
                    "Revit Addins"
                );

                Directory.CreateDirectory(commonPath);
                Console.WriteLine($"✓ Tạo thư mục chung: {commonPath}");

                // Copy DLL
                string dllDest = Path.Combine(commonPath, "SimpleBIM.dll");
                System.IO.File.Copy(dllSource, dllDest, true);
                Console.WriteLine($"✓ Sao chép DLL: {dllDest}");

                // Copy file .addin vào thư mục chung
                string addinDest = Path.Combine(commonPath, "SimpleBIM.addin");
                UpdateAddinFile(addinSource, addinDest, dllDest);
                Console.WriteLine($"✓ Sao chép .addin: {addinDest}");

                // Cài đặt cho các phiên bản Revit (trỏ đến file .addin trong thư mục chung)
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
                        CreateRevitAddinFile(revitAddinFolder, addinDest);
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
                Console.WriteLine($"• File DLL và .addin nằm tại: {commonPath}");
                Console.WriteLine("• Khởi động lại Revit để sử dụng add-in");

                // Ghi registry
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

        static void UpdateAddinFile(string sourcePath, string destPath, string dllPath)
        {
            // Đọc nội dung file .addin
            string content = System.IO.File.ReadAllText(sourcePath, System.Text.Encoding.UTF8);

            // Cập nhật đường dẫn Assembly trong file .addin
            // Tìm và thay thế đường dẫn Assembly
            int assemblyStart = content.IndexOf("<Assembly>");
            int assemblyEnd = content.IndexOf("</Assembly>");

            if (assemblyStart > 0 && assemblyEnd > assemblyStart)
            {
                string oldAssemblyPath = content.Substring(assemblyStart + 10, assemblyEnd - assemblyStart - 10);

                // Escape đường dẫn cho XML
                string escapedPath = dllPath.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");

                content = content.Replace(oldAssemblyPath, escapedPath);
                System.IO.File.WriteAllText(destPath, content, System.Text.Encoding.UTF8);
            }
            else
            {
                // Nếu không tìm thấy thẻ Assembly, tạo file mới
                string guidString = GetAddinGuid();

                string addinContent = $@"<?xml version=""1.0"" encoding=""utf-8"" standalone=""no""?>
<RevitAddIns>
  <AddIn Type=""Application"">
    <Name>SimpleBIM</Name>
    <Assembly>{dllPath}</Assembly>
    <AddInId>{guidString}</AddInId>
    <FullClassName>SimpleBIM.App</FullClassName>
    <VendorId>SIMPLEBIM</VendorId>
    <VendorDescription>SimpleBIM Add-in for Revit</VendorDescription>
  </AddIn>
</RevitAddIns>";

                System.IO.File.WriteAllText(destPath, addinContent, System.Text.Encoding.UTF8);
            }
        }

        static void CreateRevitAddinFile(string targetFolder, string commonAddinPath)
        {
            string revitAddinPath = Path.Combine(targetFolder, "SimpleBIM.addin");

            // Đọc nội dung từ file .addin chung và copy vào folder Revit
            if (System.IO.File.Exists(commonAddinPath))
            {
                System.IO.File.Copy(commonAddinPath, revitAddinPath, true);
            }
        }

        static void RegenerateGuid()
        {
            Console.Clear();
            Console.WriteLine("╔══════════════════════════════════════════════════╗");
            Console.WriteLine("║            TẠO GUID MỚI                          ║");
            Console.WriteLine("╚══════════════════════════════════════════════════╝\n");

            try
            {
                string currentDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string addinSource = Path.Combine(currentDir, "SimpleBIM.addin");

                // Xóa file .addin cũ nếu có
                if (System.IO.File.Exists(addinSource))
                {
                    Console.WriteLine("🗑️  Xóa file .addin cũ...");
                    System.IO.File.Delete(addinSource);
                }

                // Xóa file .addin trong thư mục chung
                string commonAddinPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    CompanyName,
                    "Revit Addins",
                    "SimpleBIM.addin"
                );

                if (System.IO.File.Exists(commonAddinPath))
                {
                    System.IO.File.Delete(commonAddinPath);
                    Console.WriteLine("✓ Đã xóa file .addin trong thư mục chung");
                }

                // Xóa các file .addin đã cài đặt trong Revit
                Console.WriteLine("🔄 Xóa các file .addin cũ từ Revit...");
                foreach (string version in RevitVersions)
                {
                    string addinPath = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                        "Autodesk",
                        "Revit",
                        "Addins",
                        version,
                        "SimpleBIM.addin"
                    );

                    if (System.IO.File.Exists(addinPath))
                    {
                        System.IO.File.Delete(addinPath);
                        Console.WriteLine($"   ✓ Đã xóa: Revit {version}");
                    }
                }

                // Tạo file mới với GUID mới
                Console.WriteLine("\n✨ Tạo file .addin mới với GUID mới...");
                CreateInitialAddinFile(addinSource, true);

                string newGuid = GetAddinGuid();
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"\n✅ ĐÃ TẠO GUID MỚI THÀNH CÔNG!");
                Console.ResetColor();
                Console.WriteLine($"\nGUID mới: {newGuid}");
                Console.WriteLine($"File mới: {addinSource}");
                Console.WriteLine("\n⚠️  Vui lòng CÀI ĐẶT LẠI add-in (chọn option 1)");
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"\n❌ LỖI: {ex.Message}");
                Console.ResetColor();
            }

            Pause();
        }

        static void CreateInitialAddinFile(string filePath, bool forceNewGuid = false)
        {
            // Tạo GUID hoàn toàn mới
            string guid = "{" + Guid.NewGuid().ToString().ToUpper() + "}";

            string content = $@"<?xml version=""1.0"" encoding=""utf-8"" standalone=""no""?>
<RevitAddIns>
  <AddIn Type=""Application"">
    <Name>SimpleBIM</Name>
    <Assembly>REPLACE_WITH_INSTALL_PATH</Assembly>
    <AddInId>{guid}</AddInId>
    <FullClassName>SimpleBIM.App</FullClassName>
    <VendorId>SIMPLEBIM</VendorId>
    <VendorDescription>SimpleBIM Add-in for Revit</VendorDescription>
  </AddIn>
</RevitAddIns>";

            System.IO.File.WriteAllText(filePath, content, System.Text.Encoding.UTF8);
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
            string addinFilePath = Path.Combine(currentDir, "SimpleBIM.addin");

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

            // Tạo GUID mới nếu không tìm thấy
            return "{" + Guid.NewGuid().ToString().ToUpper() + "}";
        }

        static void Uninstall(bool silentMode)
        {
            try
            {
                if (!silentMode) Console.Clear();
                Console.WriteLine("🔄 Đang gỡ cài đặt SimpleBIM...\n");

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
                        "SimpleBIM.addin"
                    );

                    if (System.IO.File.Exists(addinPath))
                    {
                        System.IO.File.Delete(addinPath);
                        removedCount++;
                        Console.WriteLine($"✓ Đã xóa: Revit {version}");
                    }
                }

                // Xóa thư mục chứa DLL và .addin
                string commonPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    CompanyName
                );

                if (Directory.Exists(commonPath))
                {
                    Directory.Delete(commonPath, true);
                    Console.WriteLine($"✓ Đã xóa thư mục: {commonPath}");
                }

                // Xóa license cũ
                string licenseFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "SimpleBIM"
                );
                if (Directory.Exists(licenseFolder))
                {
                    try
                    {
                        Directory.Delete(licenseFolder, true);
                        Console.WriteLine($"✓ Đã xóa dữ liệu license: {licenseFolder}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"⚠️  Không thể xóa license folder: {ex.Message}");
                    }
                }

                // Xóa file .addin trong thư mục installer
                string currentDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string addinSource = Path.Combine(currentDir, "SimpleBIM.addin");
                if (System.IO.File.Exists(addinSource))
                {
                    System.IO.File.Delete(addinSource);
                    Console.WriteLine("✓ Đã xóa file .addin trong thư mục installer");
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
                Console.WriteLine("• File .addin cũ đã được xóa");
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

            // Hiển thị GUID hiện tại
            string currentGuid = GetAddinGuid();
            Console.WriteLine("🔑 GUID hiện tại:");
            Console.WriteLine($"   {currentGuid}\n");

            Console.WriteLine("📂 Vị trí file DLL và .addin:");
            string commonPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                CompanyName,
                "Revit Addins"
            );

            string dllPath = Path.Combine(commonPath, "SimpleBIM.dll");
            string addinPath = Path.Combine(commonPath, "SimpleBIM.addin");

            Console.WriteLine($"   DLL: {dllPath}");
            Console.WriteLine($"   Tồn tại: {(System.IO.File.Exists(dllPath) ? "✓ CÓ" : "✗ KHÔNG")}");
            Console.WriteLine($"\n   .addin: {addinPath}");
            Console.WriteLine($"   Tồn tại: {(System.IO.File.Exists(addinPath) ? "✓ CÓ" : "✗ KHÔNG")}");

            Console.WriteLine("\n📁 Các phiên bản Revit đã cài đặt:");
            int count = 0;
            foreach (string version in RevitVersions)
            {
                string revitAddinPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "Autodesk",
                    "Revit",
                    "Addins",
                    version,
                    "SimpleBIM.addin"
                );

                if (System.IO.File.Exists(revitAddinPath))
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
                    @"Software\SimpleBIM\"
                );
                key.SetValue("InstallPath", installPath);
                key.SetValue("InstallDate", DateTime.Now.ToString("yyyy-MM-dd"));
                key.SetValue("Version", "1.0.0");
                key.SetValue("GUID", GetAddinGuid());
                key.SetValue("CommonAddinPath",
                    Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                        CompanyName,
                        "Revit Addins",
                        "SimpleBIM.addin"
                    )
                );
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