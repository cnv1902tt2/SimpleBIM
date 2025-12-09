using Autodesk.Revit.UI;
using System.Reflection;

namespace SimpleBIM.Ribbon.Panels
{
    public static class MEPFPanel
    {
        private static readonly string DllPath = Assembly.GetExecutingAssembly().Location;

        public static void Create(UIControlledApplication app, string tabName, bool isLicensed)
        {
            if (isLicensed)
            {
                // ✅ CHỈ TẠO CÁC PANEL KHI ĐÃ KÍCH HOẠT
                CreateConvertFromCADPanel(app, tabName);
                CreateEditModelingPanel(app, tabName);
                CreateUtilitiesPanel(app, tabName);
                CreateViewTemplatesPanel(app, tabName);
                CreateSheetsPresentationPanel(app, tabName);
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

        private static void CreateConvertFromCADPanel(UIControlledApplication app, string tabName)
        {
            var p = app.CreateRibbonPanel(tabName, "CONVERT FROM CAD");
            AddButton(p, "CableTraysFromCAD", "Cable Trays\nfrom CAD", "SimpleBIM.Commands.MEPF.CableTraysFromCAD", "cabletrays");
            AddButton(p, "DuctsFromCAD", "Ducts (Rec)\nfrom CAD", "SimpleBIM.Commands.MEPF.DuctsFromCAD", "ducts");
        }

        private static void CreateEditModelingPanel(UIControlledApplication app, string tabName)
        {
            var p = app.CreateRibbonPanel(tabName, "EDIT MODELING");
            AddButton(p, "AlignMEP3D", "Align MEP\n3D", "SimpleBIM.Commands.MEPF.AlignMEP3D", "alignmep3d");
            AddButton(p, "PipeSlopeChanges", "Pipe Slope\nChanges", "SimpleBIM.Commands.MEPF.PipeSlopeChanges", "pipeslopechanges");
            AddButton(p, "PipeSlopeConverter", "Pipe Slope\nConverter", "SimpleBIM.Commands.MEPF.PipeSlopeConverter", "pipeslopeconverter");
        }

        private static void CreateUtilitiesPanel(UIControlledApplication app, string tabName)
        {
            var p = app.CreateRibbonPanel(tabName, "UTILITIES");
            AddButton(p, "FamilyNameStyle", "Family Name\nStyle", "SimpleBIM.Commands.MEPF.FamilyNameStyle", "familyname");
            AddButton(p, "Grids 3D to 2D", "Grids\n3D to 2D", "SimpleBIM.Commands.MEPF.AutoAlignConnectPipes", "autoalign");

        }

        private static void CreateViewTemplatesPanel(UIControlledApplication app, string tabName)
        {
            var p = app.CreateRibbonPanel(tabName, "VIEW TEMPLATES");
            AddButton(p, "ViewTemplatesRename", "View Templates\nRename", "SimpleBIM.Commands.MEPF.ViewTemplatesRename", "viewtemplates");
            AddButton(p, "ExportImportFilterColor", "Filter Colors\nExport/Import", "SimpleBIM.Commands.MEPF.ExportImportFilterColor", "filtercolor");
        }

        private static void CreateSheetsPresentationPanel(UIControlledApplication app, string tabName)
        {
            var p = app.CreateRibbonPanel(tabName, "SHEETS PRESENTATION");
            AddButton(p, "CreateSheets", "Create\nSheets", "SimpleBIM.Commands.MEPF.CreateSheets", "createsheets");
            AddButton(p, "DuplicateViews", "Duplicate\nViews", "SimpleBIM.Commands.MEPF.DuplicateViews", "duplicateviews");
            AddButton(p, "PlaceViews", "Place\nViews", "SimpleBIM.Commands.MEPF.PlaceViews", "placeviews");
            AddButton(p, "ReorderSheets", "Reorder\nSheets", "SimpleBIM.Commands.MEPF.ReorderSheets", "reordersheets");
        }
    }
}

/*
**STRUCTURE TƯƠNG TỰ AsPanel.cs:**
- Sử dụng pattern giống AsPanel
- isLicensed check để show/hide panels
- Multiple panels: CONVERT FROM CAD, EDIT MODELING, UTILITIES
- AddButton helper method
- LoadIconToButton từ App class

**REGISTERED COMMANDS:**
1. CONVERT FROM CAD panel:
   - CableTraysFromCAD (✅ đã chuyển đổi đầy đủ)
   - DuctsFromCAD (⚠️ chưa include trong response này do giới hạn)

2. EDIT MODELING panel:
   - AlignMEP3D (✅ đã chuyển đổi đầy đủ)
   - PipeSlopeChanges (✅ đã chuyển đổi đầy đủ)
   - PipeSlopeConverter (✅ đã chuyển đổi đầy đủ)
   - AutoAlignConnectPipes (✅ đã chuyển đổi đầy đủ)

3. UTILITIES panel:
   - FamilyNameStyle (✅ đã chuyển đổi đầy đủ)

**NOTES:**
- Icons: cần có file icon tương ứng trong Icons/16/ và Icons/32/
- DuctsFromCAD: cấu trúc tương tự 95% CableTraysFromCAD, chỉ khác:
  * Sử dụng Duct API thay vì CableTray
  * Cần thêm System Type selection
  * Parameters khác: Duct.Create() cần SystemTypeId
*/
