using Autodesk.Revit.UI;

namespace SimpleBIM.Ribbon
{
    public static class RibbonManager
    {
        public static void CreateRibbon(UIControlledApplication app, string baseTabName, bool isLicensed)
        {
            // Tạo tab AS
            string asTabName = "SIMPLEBIM.AS";
            try { app.CreateRibbonTab(asTabName); } catch { }
            Panels.AsPanel.Create(app, asTabName, isLicensed);

            // ✅ TẠO TAB MEPF (đã chuyển đổi từ Python)
            string mepfTabName = "SIMPLEBIM.MEPF";
            try { app.CreateRibbonTab(mepfTabName); } catch { }
            Panels.MEPFPanel.Create(app, mepfTabName, isLicensed);

            // ✅ TẠO TAB QS (chuyển đổi từ Python - NAMING + DATA CLEANUP)
            string qsTabName = "SIMPLEBIM.QS";
            try { app.CreateRibbonTab(qsTabName); } catch { }
            Panels.QsPanel.Create(app, qsTabName, isLicensed);

            // Sau này thêm tab BS:
            // string bsTabName = "BS";
            // try { app.CreateRibbonTab(bsTabName); } catch { }
            // Panels.BsPanel.Create(app, bsTabName, isLicensed);

            // Sau này thêm tab CS:
            // string csTabName = "CS";
            // try { app.CreateRibbonTab(csTabName); } catch { }
            // Panels.CsPanel.Create(app, csTabName, isLicensed);
        }
    }
}