using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using System;
using System.Windows; // thêm cái này

namespace SimpleBIM.AS.tab
{
    [Transaction(TransactionMode.Manual)]
    public class ShowLicenseCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // ĐẢM BẢO có WPF Application instance trước khi mở Window
                if (Application.Current == null)
                {
                    // Tạo một Application instance tạm nếu chưa có
                    new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };
                }

                // Bây giờ mới được mở window
                var licenseWindow = new License.LicenseWindow();
                licenseWindow.ShowDialog();   // ShowDialog() là an toàn nhất trong command

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message + "\n\n" + ex.StackTrace;
                return Result.Failed;
            }
        }
    }
}