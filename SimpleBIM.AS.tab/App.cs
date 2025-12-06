using Autodesk.Revit.UI;
using System;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Imaging;

namespace SimpleBIM.AS.tab
{
    public class App : IExternalApplication
    {
        private static License.LicenseManager _licenseManager;
        private static License.LicenseWindow _licenseWindow;

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                string tabName = "SimpleBIM.AS";
                _licenseManager = License.LicenseManager.Instance;

                bool isLicensed = false;

                // Chỉ kiểm tra online 1 lần duy nhất khi khởi động, nếu có mạng
                if (IsNetworkAvailable())
                {
                    // Nếu có key cũ → thử validate online 1 lần
                    if (_licenseManager.CurrentLicense != null && !string.IsNullOrEmpty(_licenseManager.CurrentLicense.Key))
                    {
                        try
                        {
                            var result = Task.Run(() =>
                                _licenseManager.ValidateOnlineAsync(_licenseManager.CurrentLicense.Key, "2025")
                            ).GetAwaiter().GetResult();

                            isLicensed = result.valid;
                        }
                        catch
                        {
                            // Nếu lỗi mạng hoặc server → vẫn cho dùng offline
                            isLicensed = _licenseManager.ValidateOffline();
                        }
                    }
                    else
                    {
                        // Chưa từng có key → không check online
                        isLicensed = _licenseManager.ValidateOffline();
                    }
                }
                else
                {
                    // Không có mạng → chỉ check offline
                    isLicensed = _licenseManager.ValidateOffline();
                }

                // Dựa vào kết quả cuối cùng để quyết định hiển thị
                if (isLicensed)
                {
                    CreateFullRibbon(application, tabName);
                }
                else
                {
                    CreateMinimalRibbon(application, tabName);

                    // Tự động mở form nhập key
                    Task.Delay(800).ContinueWith(_ =>
                    {
                        Application.Current?.Dispatcher.Invoke(ShowLicenseWindow);
                    }, TaskScheduler.FromCurrentSynchronizationContext());
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("SimpleBIM Error", ex.ToString());
                return Result.Failed;
            }
        }

        private bool IsNetworkAvailable()
        {
            try
            {
                using (var client = new System.Net.WebClient())
                using (client.OpenRead("http://www.google.com"))
                    return true;
            }
            catch
            {
                return false;
            }
        }

        public Result OnShutdown(UIControlledApplication application) => Result.Succeeded;

        private void CreateFullRibbon(UIControlledApplication app, string tabName)
        {
            app.CreateRibbonTab(tabName);
            CreateArchitectureModelingPanel(app, tabName);
            CreateFinishArchitecture1Panel(app, tabName);
            CreateFinishArchitecture2Panel(app, tabName);
            CreateStructureModelingPanel(app, tabName);
        }

        private void CreateMinimalRibbon(UIControlledApplication app, string tabName)
        {
            try
            {
                app.CreateRibbonTab(tabName);
                var panel = app.CreateRibbonPanel(tabName, "License");

                var btnData = new PushButtonData("ActivateLicense", "Activate\nLicense",
                    Assembly.GetExecutingAssembly().Location,
                    "SimpleBIM.AS.tab.ShowLicenseCommand")
                {
                    ToolTip = "Kích hoạt bản quyền SimpleBIM",
                    LongDescription = "Nhấn để nhập license key và mở khóa toàn bộ chức năng add-in"
                };

                LoadIconToButton(btnData, "license"); // tự động tìm license32.png + license16.png

                panel.AddItem(btnData);
            }
            catch { }
        }

        // ====================== ICON LOADER SI ĐẸP HOÀN HẢO ======================
        private void LoadIconToButton(PushButtonData btn, string baseName)
        {
            var large = LoadSingleIcon(baseName + "32") ?? LoadSingleIcon(baseName);
            var small = LoadSingleIcon(baseName + "16") ?? LoadSingleIcon(baseName) ?? large;

            if (large != null) btn.LargeImage = large;
            if (small != null) btn.Image = small;
        }

        private BitmapImage LoadSingleIcon(string exactNameNoExtension)
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
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
        // =====================================================================

        private void AddButton(RibbonPanel panel, string name, string text, string className, string iconBaseName)
        {
            var btn = new PushButtonData(name, text,
                Assembly.GetExecutingAssembly().Location,
                "SimpleBIM.AS.tab.Commands." + className) // nếu bạn tạo thư mục Commands thì để vậy, không thì xóa ".Commands"
            {
                ToolTip = text.Replace("\n", " ")
            };

            LoadIconToButton(btn, iconBaseName);
            panel.AddItem(btn);
        }

        private void CreateArchitectureModelingPanel(UIControlledApplication app, string tabName)
        {
            var p = app.CreateRibbonPanel(tabName, "ARCHITECTURE MODELING");
            AddButton(p, "FinishSillDoor", "Finish Sill\nDoor", "FinishSillDoor", "finishsilldoor");
            AddButton(p, "WallsFromCAD", "Walls from\nCAD", "WallsFromCAD", "wallsfromcad");
        }

        private void CreateFinishArchitecture1Panel(UIControlledApplication app, string tabName)
        {
            var p = app.CreateRibbonPanel(tabName, "FINISH ARCHITECTURE 1");
            AddButton(p, "FinishRoomsCeiling", "Finish Rooms\nCeiling", "FinishRoomsCeiling", "finishroomsceiling");
            AddButton(p, "FinishRoomsFloor", "Finish Rooms\nFloor", "FinishRoomsFloor", "finishroomsfloor");
            AddButton(p, "FinishRoomsWall", "Finish Rooms\nWall", "FinishRoomsWall", "finishroomswall");
        }

        private void CreateFinishArchitecture2Panel(UIControlledApplication app, string tabName)
        {
            var p = app.CreateRibbonPanel(tabName, "FINISH ARCHITECTURE 2");
            AddButton(p, "FinishFaceFloors", "Finish Face\nFloors", "FinishFaceFloors", "finishfacefloors");
            AddButton(p, "FinishRoof", "Finish Roof", "FinishRoof", "finishroof");
            AddButton(p, "TryWallFace", "Try Wall\nFace", "TryWallFace", "trywallface");
        }

        private void CreateStructureModelingPanel(UIControlledApplication app, string tabName)
        {
            var p = app.CreateRibbonPanel(tabName, "STRUCTURE MODELING");
            AddButton(p, "AdaptiveFromCSV", "Adaptive From\nCSV", "AdaptiveFromCSV", "adaptivefromcsv");

            // Nút Try Xago
            var btnXago = new PushButtonData("TryXago", "Try Xago",
                Assembly.GetExecutingAssembly().Location,
                "SimpleBIM.AS.tab.Commands.TryXago")
            {
                ToolTip = "Create beam system from sloped face"
            };
            LoadIconToButton(btnXago, "tryxago");
            p.AddItem(btnXago);
        }

        // ====================== LICENSE WINDOW HELPER ======================
        public static void ShowLicenseWindow()
        {
            try
            {
                if (Application.Current == null)
                    new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };

                Application.Current.Dispatcher.Invoke(() =>
                {
                    if (_licenseWindow == null || _licenseWindow.IsLoaded == false)
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
            if (_licenseManager == null)
                _licenseManager = License.LicenseManager.Instance;
            if (!_licenseManager.ValidateOffline())
            {
                ShowLicenseWindow();
                return false;
            }
            return true;
        }
    }
}