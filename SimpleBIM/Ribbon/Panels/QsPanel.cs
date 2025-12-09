using Autodesk.Revit.UI;
using System.Reflection;

namespace SimpleBIM.Ribbon.Panels
{
    public static class QsPanel
    {
        private static readonly string DllPath = Assembly.GetExecutingAssembly().Location;

        public static void Create(UIControlledApplication app, string tabName, bool isLicensed)
        {
            if (isLicensed)
            {
                // ✅ CHỈ TẠO CÁC PANEL KHI ĐÃ KÍCH HOẠT
                CreateNamingPanel(app, tabName);
                CreateDataCleanupPanel(app, tabName);
                CreateVNNormPanel(app, tabName);
            }
            else
            {
                // ✅ CHƯA KÍCH HOẠT: CHỈ HIỂN THỊ PANEL LICENSE
                var licensePanel = app.CreateRibbonPanel(tabName, "LICENSE");
                AddActivateLicenseButton(licensePanel);
            }
        }

        private static void AddButton(RibbonPanel panel, string name, string text, string fullClassName, string icon)
        {
            var btn = new PushButtonData(name, text, DllPath, fullClassName)
            {
                ToolTip = text.Replace("\n", " ")
            };
            App.LoadIconToButton(btn, icon);
            panel.AddItem(btn);
        }

        private static void AddActivateLicenseButton(RibbonPanel panel)
        {
            var btn = new PushButtonData("ActivateLicense", "Activate\nLicense", DllPath, "SimpleBIM.ShowLicenseCommand")
            {
                ToolTip = "Kích hoạt bản quyền SimpleBIM",
                LongDescription = "Nhấn để nhập license key và mở khóa toàn bộ chức năng add-in"
            };
            App.LoadIconToButton(btn, "license");
            panel.AddItem(btn);
        }

        private static void CreateNamingPanel(UIControlledApplication app, string tabName)
        {
            var p = app.CreateRibbonPanel(tabName, "NAMING");
            AddButton(p, "FamilyNamesExportImport", "Family Names\nExport/Import", "SimpleBIM.Commands.Qs.FamilyNamesExportImport", "familynamesexportimport");
            AddButton(p, "MaterialNamesExportImport", "Material Names\nExport/Import", "SimpleBIM.Commands.Qs.MaterialNamesExportImport", "materialnamesexportimport");
            AddButton(p, "TypeNamesExportImport", "Type Names\nExport/Import", "SimpleBIM.Commands.Qs.TypeNamesExportImport", "typenamesexportimport");
        }

        private static void CreateDataCleanupPanel(UIControlledApplication app, string tabName)
        {
            var p = app.CreateRibbonPanel(tabName, "DATA CLEANUP");
            AddButton(p, "DeleteDuplicates", "Delete\nDuplicates", "SimpleBIM.Commands.Qs.DeleteDuplicates", "deleteduplicates");
            AddButton(p, "ExportSchedules", "Export\nSchedules", "SimpleBIM.Commands.Qs.ExportSchedules", "exportschedules");
        }

        private static void CreateVNNormPanel(UIControlledApplication app, string tabName)
        {
            var p = app.CreateRibbonPanel(tabName, "VN NORM CODES");
            AddButton(p, "VNNormCodes", "VN Norm\nCodes", "SimpleBIM.Commands.Qs.VNNormCodes", "vnnormcodes");
            AddButton(p, "TinhNangTimKiem", "Tính Năng\nTìm Kiếm", "SimpleBIM.Commands.Qs.TinhNangTimKiem", "tinhnangtimnkiem");
        }
    }
}
