using Autodesk.Revit.UI;
using SimpleBIM.License;
using System;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Imaging;

namespace SimpleBIM
{
    public class App : IExternalApplication
    {
        private static LicenseWindow _licenseWindow;
        private static UIControlledApplication _uiApp;
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                _uiApp = application;
                string tabName = "SimpleBIM";
                bool isLicensed = LicenseManager.Instance.ValidateOffline();

                Ribbon.RibbonManager.CreateRibbon(application, tabName, isLicensed);

                if (!isLicensed)
                {
                    Task.Delay(600).ContinueWith(_ =>
                        Application.Current?.Dispatcher.Invoke(ShowLicenseWindow));
                }

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
                if (Application.Current == null)
                {
                    new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown }.Run();
                }

                Application.Current.Dispatcher.Invoke(() =>
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
    }
}