using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SimpleBIM.Commands.MEPF
{
    /// <summary>
    /// Family Name to Uppercase Converter
    /// Chuyển đổi tên Family sang chữ hoa
    /// Converted from Python to C#
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class FamilyNameStyle : IExternalCommand
    {
        private Document _doc;

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            _doc = commandData.Application.ActiveUIDocument.Document;

            try
            {
                return ShowMainMenu();
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Error", $"Error: {ex.Message}");
                return Result.Failed;
            }
        }

        private Result ShowMainMenu()
        {
            // Simple menu using TaskDialog
            TaskDialog td = new TaskDialog("Family Name Converter");
            td.MainInstruction = "Chọn chức năng:";
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "🔤 Chuyển tên FAMILY sang CHỮ HOA");
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "📊 Xem thống kê tên Family");
            td.CommonButtons = TaskDialogCommonButtons.Close;

            TaskDialogResult result = td.Show();

            if (result == TaskDialogResult.CommandLink1)
            {
                return ConvertFamilyNames();
            }
            else if (result == TaskDialogResult.CommandLink2)
            {
                return GetFamilyStatistics();
            }

            return Result.Cancelled;
        }

        private Result ConvertFamilyNames()
        {
            System.Diagnostics.Debug.WriteLine("Bắt đầu chuyển đổi tên Family...");

            // Lấy tất cả families
            List<Family> families = GetAllFamilies();
            if (families.Count == 0)
            {
                TaskDialog.Show("Thông báo", "Không tìm thấy Family nào trong model!");
                return Result.Cancelled;
            }

            System.Diagnostics.Debug.WriteLine($"Tìm thấy {families.Count} families trong model");

            // Chuẩn bị danh sách thay đổi
            List<FamilyChange> familyChanges = PreviewFamilyChanges(families);
            System.Diagnostics.Debug.WriteLine($"Tìm thấy {familyChanges.Count} family cần đổi tên");

            // Hiển thị preview
            if (!ShowPreview(familyChanges))
            {
                return Result.Cancelled;
            }

            // Thực hiện thay đổi
            using (Transaction t = new Transaction(_doc, "Chuyển tên Family sang CHỮ HOA"))
            {
                t.Start();
                var (successCount, errorCount) = ApplyFamilyChanges(familyChanges);
                t.Commit();

                // Báo cáo kết quả
                string resultMsg = "🎉 HOÀN THÀNH CHUYỂN ĐỔI TÊN FAMILY!\n\n";
                resultMsg += $"✅ Thành công: {successCount} families\n";
                if (errorCount > 0)
                {
                    resultMsg += $"❌ Lỗi/Bỏ qua: {errorCount} families\n";
                }
                resultMsg += "\n💾 Hãy Save file để lưu thay đổi!";

                TaskDialog.Show("Kết quả", resultMsg);
                System.Diagnostics.Debug.WriteLine($"KẾT QUẢ: {successCount} thành công, {errorCount} lỗi");
            }

            return Result.Succeeded;
        }

        private Result GetFamilyStatistics()
        {
            List<Family> families = GetAllFamilies();

            if (families.Count == 0)
            {
                TaskDialog.Show("Thông báo", "Không tìm thấy Family nào trong model!");
                return Result.Cancelled;
            }

            // Phân tích Family names
            var stats = new
            {
                total = families.Count,
                upper = families.Count(f => f.Name.ToUpper() == f.Name && f.Name.Any(char.IsLetter)),
                lower = families.Count(f => f.Name.ToLower() == f.Name && f.Name.Any(char.IsLetter)),
                title = families.Count(f => System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(f.Name.ToLower()) == f.Name && f.Name.Any(char.IsLetter)),
                mixed = 0
            };

            int needChange = stats.total - stats.upper;

            // Hiển thị thống kê
            string report = "📊 THỐNG KÊ TÊN FAMILY\n\n";
            report += $"🏠 FAMILIES ({stats.total} total):\n";
            report += $"  • CHỮ HOA: {stats.upper}\n";
            report += $"  • chữ thường: {stats.lower}\n";
            report += $"  • Title Case: {stats.title}\n";
            report += $"  • Hỗn hợp: {stats.total - stats.upper - stats.lower - stats.title}\n\n";

            if (needChange > 0)
            {
                report += $"🔄 Cần chuyển sang CHỮ HOA: {needChange} families";
            }
            else
            {
                report += "✅ Tất cả Family names đã là CHỮ HOA!";
            }

            TaskDialog.Show("Thống kê", report);
            return Result.Succeeded;
        }

        private List<Family> GetAllFamilies()
        {
            FilteredElementCollector collector = new FilteredElementCollector(_doc).OfClass(typeof(Family));
            return collector.Cast<Family>().ToList();
        }

        private List<FamilyChange> PreviewFamilyChanges(List<Family> families)
        {
            List<FamilyChange> changes = new List<FamilyChange>();

            foreach (Family family in families)
            {
                string oldName = family.Name;
                string newName = oldName.ToUpper();

                if (oldName != newName)
                {
                    changes.Add(new FamilyChange
                    {
                        Element = family,
                        OldName = oldName,
                        NewName = newName
                    });
                }
            }

            return changes;
        }

        private bool ShowPreview(List<FamilyChange> familyChanges)
        {
            if (familyChanges.Count == 0)
            {
                TaskDialog.Show("Thông báo", "Không có tên Family nào cần thay đổi!\nTất cả đã là chữ hoa.");
                return false;
            }

            string preview = $"🔄 SẼ CHUYỂN {familyChanges.Count} FAMILY NAMES SANG CHỮ HOA:\n\n";

            // Hiển thị tối đa 15 family names
            for (int i = 0; i < Math.Min(15, familyChanges.Count); i++)
            {
                preview += $"{i + 1}. {familyChanges[i].OldName} → {familyChanges[i].NewName}\n";
            }

            if (familyChanges.Count > 15)
            {
                preview += $"   ... và {familyChanges.Count - 15} family khác\n";
            }

            preview += "\n⚠️ LƯU Ý: Thao tác này KHÔNG THỂ HOÀN TÁC!";
            preview += "\n💾 Hãy backup file trước khi tiếp tục!";

            TaskDialogResult result = TaskDialog.Show(
                "Xác nhận",
                preview + "\n\nTiếp tục thực hiện?",
                TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

            return result == TaskDialogResult.Yes;
        }

        private (int, int) ApplyFamilyChanges(List<FamilyChange> familyChanges)
        {
            int successCount = 0;
            int errorCount = 0;

            foreach (FamilyChange change in familyChanges)
            {
                try
                {
                    Family family = change.Element;
                    string newName = change.NewName;

                    // Kiểm tra tên mới có trùng không
                    if (IsFamilyNameExists(newName, family.Id))
                    {
                        System.Diagnostics.Debug.WriteLine($"⚠️  Tên Family '{newName}' đã tồn tại, bỏ qua");
                        errorCount++;
                        continue;
                    }

                    family.Name = newName;
                    successCount++;
                    System.Diagnostics.Debug.WriteLine($"✅ Family: {change.OldName} → {newName}");
                }
                catch (Exception e)
                {
                    errorCount++;
                    System.Diagnostics.Debug.WriteLine($"❌ Lỗi đổi tên Family '{change.OldName}': {e.Message}");
                }
            }

            return (successCount, errorCount);
        }

        private bool IsFamilyNameExists(string name, ElementId excludeId)
        {
            List<Family> families = GetAllFamilies();
            foreach (Family family in families)
            {
                if (family.Name == name && family.Id != excludeId)
                {
                    return true;
                }
            }
            return false;
        }

        private class FamilyChange
        {
            public Family Element { get; set; }
            public string OldName { get; set; }
            public string NewName { get; set; }
        }
    }
}

/* 
**PYREVIT → C# CONVERSIONS APPLIED:**
1. `forms.alert()` → `TaskDialog.Show()`
2. `forms.SelectFromList.show()` → `TaskDialog` với `CommandLink`
3. `print()` → `System.Diagnostics.Debug.WriteLine()`
4. Python class attributes → C# class with properties
5. `with revit.Transaction()` → `using (Transaction t = new Transaction())`
6. Python string methods `.upper()` → C# `.ToUpper()`

**THAM KHẢO TỪ Commands/As/:**
- IExternalCommand structure
- FilteredElementCollector patterns
- Transaction handling

**IMPORTANT NOTES:**
- Chuyển đổi tên Family sang uppercase
- Kiểm tra duplicate names trước khi rename
- Preview changes trước khi apply
- Statistics về family name formats
*/
