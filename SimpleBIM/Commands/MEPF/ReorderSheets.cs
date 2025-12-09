using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SimpleBIM.Commands.MEPF
{
    /// <summary>
    /// Sheet Reorder Tool - Xử lý trùng số sheet
    /// FULL CONVERSION (300 lines Python)
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ReorderSheets : IExternalCommand
    {
        private UIDocument _uidoc;
        private Document _doc;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            _uidoc = commandData.Application.ActiveUIDocument;
            _doc = _uidoc.Document;

            if (_doc.IsFamilyDocument)
            {
                TaskDialog.Show("Lỗi", "Script chỉ chạy trong Project!");
                return Result.Failed;
            }

            try
            {
                return ExecuteReorder();
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Lỗi", $"Lỗi thực thi: {ex.Message}");
                return Result.Failed;
            }
        }

        private Result ExecuteReorder()
        {
            // Step 1: Get all sheets
            List<SheetData> sheets = GetSheetsByDiscipline();
            if (sheets.Count == 0)
            {
                TaskDialog.Show("Thông báo", "Không tìm thấy sheets nào trong project!");
                return Result.Cancelled;
            }

            // Step 2: Check duplicates
            var allNumbers = sheets.Select(s => s.CurrentNumber).ToList();
            var duplicates = allNumbers.GroupBy(n => n).Where(g => g.Count() > 1).Select(g => g.Key).ToList();

            if (duplicates.Count > 0)
            {
                TaskDialog warningDialog = new TaskDialog("Cảnh báo");
                warningDialog.MainContent = $"⚠️ CẢNH BÁO: Có {duplicates.Count} số sheet bị trùng:\n{string.Join(", ", duplicates)}\n\nTool sẽ tự động xử lý.";
                warningDialog.CommonButtons = TaskDialogCommonButtons.Ok;
                warningDialog.Show();
            }

            // Step 3: Group by discipline and renumber
            Dictionary<string, List<SheetData>> disciplineGroups = RenumberSheetsByDiscipline(sheets);

            // Step 4: Create new numbering
            List<SheetChangeData> changes = ApplyNewNumbering(disciplineGroups);

            // Step 5: Preview
            string preview = ShowPreviewChanges(changes);

            // Step 6: Confirm
            TaskDialog confirmDialog = new TaskDialog("Xác nhận");
            confirmDialog.MainInstruction = "CÁC THAY ĐỔI SẼ ĐƯỢC ÁP DỤNG:";
            confirmDialog.MainContent = preview;
            confirmDialog.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;

            if (confirmDialog.Show() != TaskDialogResult.Yes)
                return Result.Cancelled;

            // Step 7: Apply changes (2 steps to avoid duplicates)
            return ApplyChangesSafely(changes, disciplineGroups);
        }

        // =============================================================================
        // GET SHEETS
        // =============================================================================
        private List<SheetData> GetSheetsByDiscipline()
        {
            List<SheetData> sheets = new List<SheetData>();
            FilteredElementCollector collector = new FilteredElementCollector(_doc).OfClass(typeof(ViewSheet));

            foreach (ViewSheet sheet in collector)
            {
                if (sheet.IsPlaceholder)
                    continue;

                string sheetNumber = sheet.SheetNumber;
                string sheetName = sheet.Name;

                // Parse prefix, number, suffix
                Match match = Regex.Match(sheetNumber, @"^([A-Za-z]*)(\d+)([A-Z]?)");
                if (match.Success)
                {
                    string prefix = match.Groups[1].Value;
                    int number = int.Parse(match.Groups[2].Value);
                    string suffix = match.Groups[3].Value;

                    sheets.Add(new SheetData
                    {
                        Sheet = sheet,
                        CurrentNumber = sheetNumber,
                        Name = sheetName,
                        Prefix = prefix,
                        OriginalNumber = number,
                        Suffix = suffix
                    });
                }
                else
                {
                    sheets.Add(new SheetData
                    {
                        Sheet = sheet,
                        CurrentNumber = sheetNumber,
                        Name = sheetName,
                        Prefix = "",
                        OriginalNumber = 0,
                        Suffix = ""
                    });
                }
            }

            return sheets;
        }

        // =============================================================================
        // RENUMBER BY DISCIPLINE
        // =============================================================================
        private Dictionary<string, List<SheetData>> RenumberSheetsByDiscipline(List<SheetData> sheets)
        {
            Dictionary<string, List<SheetData>> disciplineGroups = new Dictionary<string, List<SheetData>>();

            // Group by prefix
            foreach (var sheet in sheets)
            {
                string prefix = sheet.Prefix;
                if (!disciplineGroups.ContainsKey(prefix))
                    disciplineGroups[prefix] = new List<SheetData>();
                
                disciplineGroups[prefix].Add(sheet);
            }

            // Sort each group
            foreach (var kvp in disciplineGroups)
            {
                kvp.Value.Sort((a, b) =>
                {
                    int cmp = a.OriginalNumber.CompareTo(b.OriginalNumber);
                    if (cmp != 0) return cmp;
                    return string.Compare(a.Suffix, b.Suffix, StringComparison.Ordinal);
                });
            }

            return disciplineGroups;
        }

        // =============================================================================
        // APPLY NEW NUMBERING
        // =============================================================================
        private List<SheetChangeData> ApplyNewNumbering(Dictionary<string, List<SheetData>> disciplineGroups)
        {
            List<SheetChangeData> changes = new List<SheetChangeData>();

            foreach (var kvp in disciplineGroups)
            {
                string prefix = kvp.Key;
                List<SheetData> groupSheets = kvp.Value;

                if (string.IsNullOrEmpty(prefix))
                {
                    // Sheets without prefix - use temp numbers
                    for (int i = 0; i < groupSheets.Count; i++)
                    {
                        string tempNum = $"TEMP_{i + 1:D4}";
                        string finalNum = (i + 1).ToString();

                        changes.Add(new SheetChangeData
                        {
                            Sheet = groupSheets[i].Sheet,
                            OldNumber = groupSheets[i].CurrentNumber,
                            TempNumber = tempNum,
                            FinalNumber = finalNum,
                            Name = groupSheets[i].Name
                        });
                    }
                }
                else
                {
                    // Sheets with prefix
                    int startNumber = groupSheets.Min(s => s.OriginalNumber);

                    for (int i = 0; i < groupSheets.Count; i++)
                    {
                        string tempNum = $"TEMP_{prefix}_{i:D4}";
                        string finalNum = $"{prefix}{(startNumber + i):D3}";

                        changes.Add(new SheetChangeData
                        {
                            Sheet = groupSheets[i].Sheet,
                            OldNumber = groupSheets[i].CurrentNumber,
                            TempNumber = tempNum,
                            FinalNumber = finalNum,
                            Name = groupSheets[i].Name
                        });
                    }
                }
            }

            return changes;
        }

        // =============================================================================
        // PREVIEW CHANGES
        // =============================================================================
        private string ShowPreviewChanges(List<SheetChangeData> changes)
        {
            string preview = "STT | Số cũ → Số mới | Tên Sheet\n";
            preview += new string('-', 60) + "\n";

            for (int i = 0; i < Math.Min(20, changes.Count); i++)
            {
                var change = changes[i];
                preview += $"{i + 1:D2}. {change.OldNumber} → {change.FinalNumber} | {change.Name}\n";
            }

            if (changes.Count > 20)
                preview += $"... và {changes.Count - 20} sheets khác\n";

            return preview;
        }

        // =============================================================================
        // APPLY CHANGES SAFELY (2 STEPS)
        // =============================================================================
        private Result ApplyChangesSafely(List<SheetChangeData> changes, Dictionary<string, List<SheetData>> disciplineGroups)
        {
            try
            {
                // STEP 1: Rename to temp numbers
                using (Transaction trans1 = new Transaction(_doc, "Step 1 - Temporary Renumbering"))
                {
                    trans1.Start();

                    foreach (var change in changes)
                    {
                        change.Sheet.SheetNumber = change.TempNumber;
                    }

                    trans1.Commit();
                }

                // STEP 2: Rename to final numbers (with duplicate checking)
                using (Transaction trans2 = new Transaction(_doc, "Step 2 - Final Renumbering"))
                {
                    trans2.Start();

                    foreach (var change in changes)
                    {
                        string finalNumber = change.FinalNumber;

                        // Check for duplicates
                        FilteredElementCollector collector = new FilteredElementCollector(_doc)
                            .OfClass(typeof(ViewSheet))
                            .WhereElementIsNotElementType();

                        HashSet<string> existingNumbers = new HashSet<string>();
                        foreach (ViewSheet s in collector)
                        {
                            existingNumbers.Add(s.SheetNumber);
                        }

                        // If duplicate, find next available
                        if (existingNumbers.Contains(finalNumber))
                        {
                            Match match = Regex.Match(finalNumber, @"([A-Za-z]*)(\d+)");
                            if (match.Success)
                            {
                                string prefix = match.Groups[1].Value;
                                int num = int.Parse(match.Groups[2].Value);

                                while (existingNumbers.Contains(finalNumber))
                                {
                                    num++;
                                    finalNumber = $"{prefix}{num:D3}";
                                }

                                change.FinalNumber = finalNumber;
                            }
                        }

                        change.Sheet.SheetNumber = finalNumber;
                    }

                    trans2.Commit();
                }

                // Report success
                string disciplineSummary = "";
                foreach (var kvp in disciplineGroups)
                {
                    string disciplineName = GetDisciplineName(kvp.Key);
                    disciplineSummary += $"{disciplineName}: {kvp.Value.Count} sheets\n";
                }

                TaskDialog.Show("Hoàn thành", 
                    $"✅ ĐÃ HOÀN THÀNH!\n\n" +
                    $"Đã cập nhật {changes.Count} sheets:\n{disciplineSummary}\n" +
                    $"Số sheet đã được sắp xếp lại liền mạch.");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Lỗi", $"❌ Lỗi khi cập nhật: {ex.Message}");
                return Result.Failed;
            }
        }

        private string GetDisciplineName(string prefix)
        {
            Dictionary<string, string> disciplineNames = new Dictionary<string, string>
            {
                {"", "Chung"}, {"M", "Kiến trúc"}, {"A", "Kết cấu"},
                {"S", "Điện"}, {"E", "Cơ điện"}, {"P", "Cấp thoát nước"},
                {"C", "Nội thất"}, {"L", "Cảnh quan"}, {"T", "Giao thông"},
                {"H", "Thiết bị"}, {"Q", "Giám sát"}, {"G", "PCCC"}
            };

            return disciplineNames.ContainsKey(prefix) ? disciplineNames[prefix] : $"Bộ môn {prefix}";
        }

        // =============================================================================
        // DATA STRUCTURES
        // =============================================================================
        private class SheetData
        {
            public ViewSheet Sheet { get; set; }
            public string CurrentNumber { get; set; }
            public string Name { get; set; }
            public string Prefix { get; set; }
            public int OriginalNumber { get; set; }
            public string Suffix { get; set; }
        }

        private class SheetChangeData
        {
            public ViewSheet Sheet { get; set; }
            public string OldNumber { get; set; }
            public string TempNumber { get; set; }
            public string FinalNumber { get; set; }
            public string Name { get; set; }
        }
    }
}

/* 
**PYREVIT → C# CONVERSIONS (300 LINES PYTHON):**
✅ Sheet parsing với Regex (prefix-number-suffix)
✅ Discipline grouping và sorting
✅ 2-step renumbering để tránh trùng số:
   - Step 1: Rename tất cả sang TEMP numbers
   - Step 2: Rename từ TEMP sang final numbers
✅ Duplicate detection và auto-fix
✅ Preview changes trước khi apply
✅ Transaction management đúng chuẩn
*/
