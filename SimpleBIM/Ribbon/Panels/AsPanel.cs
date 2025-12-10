using Autodesk.Revit.UI;
using System.Reflection;

namespace SimpleBIM.Ribbon.Panels
{
    public static class AsPanel
    {
        private static readonly string DllPath = Assembly.GetExecutingAssembly().Location;

        public static void Create(UIControlledApplication app, string tabName, bool isLicensed)
        {
            if (isLicensed)
            {
                // ✅ CHỈ TẠO CÁC PANEL KHI ĐÃ KÍCH HOẠT
                CreateArchitectureModelingPanel(app, tabName);
                CreateFinishArchitecture1Panel(app, tabName);
                CreateFinishArchitecture2Panel(app, tabName);
                CreateStructureModelingPanel(app, tabName);
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

        private static void CreateArchitectureModelingPanel(UIControlledApplication app, string tabName)
        {
            var p = app.CreateRibbonPanel(tabName, "ARCHITECTURE MODELING");
            AddButton(p, "FinishSillDoor", "Finish Sill\nDoor", "SimpleBIM.Commands.As.FinishSillDoor", "finishsilldoor");
            AddButton(p, "WallsFromCAD", "Walls from\nCAD", "SimpleBIM.Commands.As.WallsFromCAD", "wallsfromcad");
        }

        private static void CreateFinishArchitecture1Panel(UIControlledApplication app, string tabName)
        {
            var p = app.CreateRibbonPanel(tabName, "FINISH ARCHITECTURE 1");
            AddButton(p, "FinishRoomsCeiling", "Finish Rooms\nCeiling", "SimpleBIM.Commands.As.FinishRoomsCeiling", "finishroomsceiling");
            AddButton(p, "FinishRoomsFloor", "Finish Rooms\nFloor", "SimpleBIM.Commands.As.FinishRoomsFloor", "finishroomsfloor");
            AddButton(p, "FinishRoomsWall", "Finish Rooms\nWall", "SimpleBIM.Commands.As.FinishRoomsWall", "finishroomswall");
        }

        private static void CreateFinishArchitecture2Panel(UIControlledApplication app, string tabName)
        {
            var p = app.CreateRibbonPanel(tabName, "FINISH ARCHITECTURE 2");
            AddButton(p, "FinishFaceFloors", "Finish Face\nFloors", "SimpleBIM.Commands.As.FinishFaceFloors", "finishfacefloors");
            AddButton(p, "FinishRoof", "Finish Roof", "SimpleBIM.Commands.As.FinishRoof", "finishroof");
            AddButton(p, "TryWallFace", "Try Wall\nFace", "SimpleBIM.Commands.As.TryWallFace", "trywallface");
        }

        private static void CreateStructureModelingPanel(UIControlledApplication app, string tabName)
        {
            var p = app.CreateRibbonPanel(tabName, "STRUCTURE MODELING");
            AddButton(p, "AdaptiveFromCSV", "Adaptive From\nCSV", "SimpleBIM.Commands.As.AdaptiveFromCSV", "adaptivefromcsv");
            AddButton(p, "TryXago", "Try Xago", "SimpleBIM.Commands.As.TryXago", "tryxago");
        }
    }
}