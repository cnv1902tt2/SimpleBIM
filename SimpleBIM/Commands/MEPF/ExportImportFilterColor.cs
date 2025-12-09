using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Text;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SimpleBIM.Commands.MEPF
{
    /// <summary>
    /// Export/Import Filter Colors from/to View Templates
    /// Converted from Python to C# - FULL CONVERSION (263 lines Python)
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ExportImportFilterColor : IExternalCommand
    {
        private UIDocument _uidoc;
        private Document _doc;

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            _uidoc = commandData.Application.ActiveUIDocument;
            _doc = _uidoc.Document;

            try
            {
                return ShowMainDialog();
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Lỗi", $"Lỗi thực thi: {ex.Message}");
                return Result.Failed;
            }
        }

        // =============================================================================
        // MAIN DIALOG
        // =============================================================================
        private Result ShowMainDialog()
        {
            using (var form = new FilterColorMainForm())
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    if (form.SelectedAction == "Export")
                        return ExportFilterColors();
                    else if (form.SelectedAction == "Import")
                        return ImportFilterColors();
                }
            }
            return Result.Cancelled;
        }

        // =============================================================================
        // EXPORT FILTER COLORS
        // =============================================================================
        private Result ExportFilterColors()
        {
            try
            {
                // Lấy tất cả View Templates
                FilteredElementCollector collector = new FilteredElementCollector(_doc).OfClass(typeof(Autodesk.Revit.DB.View));
                List<Autodesk.Revit.DB.View> templateViews = new List<Autodesk.Revit.DB.View>();

                foreach (Autodesk.Revit.DB.View v in collector)
                {
                    if (v.IsTemplate)
                        templateViews.Add(v);
                }

                if (templateViews.Count == 0)
                {
                    TaskDialog.Show("Thông báo", "Không tìm thấy View Template nào!");
                    return Result.Cancelled;
                }

                // Chọn đường dẫn lưu file
                string outputDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string filePath = Path.Combine(outputDir, "FilterColors.csv");

                int rowCount = 0;

                // Xuất dữ liệu ra CSV
                using (StreamWriter writer = new StreamWriter(filePath, false, Encoding.UTF8))
                {
                    writer.WriteLine("View Template Name,Filter Name,Line Color (RGB),Pattern Foreground Color (RGB),Pattern Background Color (RGB)");

                    foreach (Autodesk.Revit.DB.View vt in templateViews)
                    {
                        try
                        {
                            ICollection<ElementId> filterIds = vt.GetFilters();

                            if (filterIds == null || filterIds.Count == 0)
                                continue;

                            foreach (ElementId fid in filterIds)
                            {
                                try
                                {
                                    Element filterElem = _doc.GetElement(fid);
                                    if (filterElem == null)
                                        continue;

                                    string filterName = filterElem.Name ?? "Unnamed Filter";
                                    OverrideGraphicSettings overrideSettings = vt.GetFilterOverrides(fid);

                                    // Lấy màu cho Lines
                                    Color lineColor = overrideSettings.ProjectionLineColor;
                                    string lineColorStr = (lineColor != null && lineColor.IsValid)
                                        ? $"{lineColor.Red}-{lineColor.Green}-{lineColor.Blue}"
                                        : "None";

                                    // Lấy màu cho Pattern Foreground
                                    Color patternFgColor = overrideSettings.SurfaceForegroundPatternColor;
                                    string patternFgColorStr = (patternFgColor != null && patternFgColor.IsValid)
                                        ? $"{patternFgColor.Red}-{patternFgColor.Green}-{patternFgColor.Blue}"
                                        : "None";

                                    // Lấy màu cho Pattern Background
                                    Color patternBgColor = overrideSettings.SurfaceBackgroundPatternColor;
                                    string patternBgColorStr = (patternBgColor != null && patternBgColor.IsValid)
                                        ? $"{patternBgColor.Red}-{patternBgColor.Green}-{patternBgColor.Blue}"
                                        : "None";

                                    // Escape CSV fields
                                    string vtNameEscaped = EscapeCsvField(vt.Name);
                                    string filterNameEscaped = EscapeCsvField(filterName);

                                    writer.WriteLine($"{vtNameEscaped},{filterNameEscaped},{lineColorStr},{patternFgColorStr},{patternBgColorStr}");
                                    rowCount++;
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"Lỗi khi xử lý filter {fid}: {ex.Message}");
                                    continue;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Lỗi khi xử lý view template '{vt.Name}': {ex.Message}");
                            continue;
                        }
                    }
                }

                string msg = $"EXPORT HOÀN TẤT!\n\n" +
                             $"Số filter đã xuất: {rowCount}\n" +
                             $"File lưu tại:\n{filePath}";
                TaskDialog.Show("Export Complete", msg);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Lỗi", $"Lỗi export: {ex.Message}");
                return Result.Failed;
            }
        }

        // =============================================================================
        // IMPORT FILTER COLORS
        // =============================================================================
        private Result ImportFilterColors()
        {
            try
            {
                // Chọn file CSV
                string outputDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                OpenFileDialog openDialog = new OpenFileDialog
                {
                    Title = "Chọn file FilterColors.csv để Import",
                    Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                    InitialDirectory = outputDir
                };

                if (openDialog.ShowDialog() != DialogResult.OK)
                {
                    TaskDialog.Show("Thông báo", "Không có file nào được chọn!");
                    return Result.Cancelled;
                }

                string csvFile = openDialog.FileName;

                // Đọc file CSV
                List<FilterColorData> rowsData = new List<FilterColorData>();
                using (StreamReader reader = new StreamReader(csvFile, Encoding.UTF8))
                {
                    // Bỏ qua header
                    string headerLine = reader.ReadLine();

                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        var parts = ParseCsvLine(line);
                        if (parts.Count >= 5)
                        {
                            rowsData.Add(new FilterColorData
                            {
                                ViewTemplateName = parts[0].Trim(),
                                FilterName = parts[1].Trim(),
                                LineColorStr = parts[2].Trim(),
                                PatternFgColorStr = parts[3].Trim(),
                                PatternBgColorStr = parts[4].Trim()
                            });
                        }
                    }
                }

                if (rowsData.Count == 0)
                {
                    TaskDialog.Show("Thông báo", "File CSV không có dữ liệu!");
                    return Result.Cancelled;
                }

                // Lấy tất cả View Templates
                FilteredElementCollector collector = new FilteredElementCollector(_doc).OfClass(typeof(Autodesk.Revit.DB.View));
                Dictionary<string, Autodesk.Revit.DB.View> templateViews = new Dictionary<string, Autodesk.Revit.DB.View>();

                foreach (Autodesk.Revit.DB.View v in collector)
                {
                    if (v.IsTemplate)
                    {
                        if (templateViews.ContainsKey(v.Name))
                        {
                            TaskDialog.Show("Lỗi", $"Lỗi: Có View Template bị trùng tên: {v.Name}");
                            return Result.Failed;
                        }
                        templateViews[v.Name] = v;
                    }
                }

                // Bắt đầu transaction
                int successCount = 0;
                List<string> errorLog = new List<string>();

                using (Transaction t = new Transaction(_doc, "Import Filter Colors"))
                {
                    t.Start();

                    try
                    {
                        foreach (FilterColorData row in rowsData)
                        {
                            // Tìm View Template
                            if (!templateViews.ContainsKey(row.ViewTemplateName))
                            {
                                errorLog.Add($"View Template '{row.ViewTemplateName}' không tồn tại");
                                continue;
                            }

                            Autodesk.Revit.DB.View vt = templateViews[row.ViewTemplateName];

                            // Lấy tất cả filters trong View Template
                            ICollection<ElementId> filterIds = vt.GetFilters();
                            List<ElementId> matchingFilters = new List<ElementId>();

                            foreach (ElementId fid in filterIds)
                            {
                                Element filterElem = _doc.GetElement(fid);
                                if (filterElem != null && filterElem.Name == row.FilterName)
                                {
                                    matchingFilters.Add(fid);
                                }
                            }

                            if (matchingFilters.Count == 0)
                            {
                                errorLog.Add($"Filter '{row.FilterName}' không tồn tại trong View Template '{row.ViewTemplateName}'");
                                continue;
                            }

                            if (matchingFilters.Count > 1)
                            {
                                errorLog.Add($"Filter '{row.FilterName}' bị trùng lặp trong View Template '{row.ViewTemplateName}'");
                                continue;
                            }

                            ElementId targetFilterId = matchingFilters[0];

                            // Parse colors
                            Color lineColor = ParseColor(row.LineColorStr);
                            Color patternFgColor = ParseColor(row.PatternFgColorStr);
                            Color patternBgColor = ParseColor(row.PatternBgColorStr);

                            // Lấy override hiện tại và set màu mới
                            try
                            {
                                OverrideGraphicSettings overrideSettings = vt.GetFilterOverrides(targetFilterId);

                                if (lineColor != null)
                                    overrideSettings.SetProjectionLineColor(lineColor);

                                if (patternFgColor != null)
                                    overrideSettings.SetSurfaceForegroundPatternColor(patternFgColor);

                                if (patternBgColor != null)
                                    overrideSettings.SetSurfaceBackgroundPatternColor(patternBgColor);

                                vt.SetFilterOverrides(targetFilterId, overrideSettings);
                                successCount++;
                            }
                            catch (Exception ex)
                            {
                                errorLog.Add($"Lỗi khi apply màu cho Filter '{row.FilterName}' trong '{row.ViewTemplateName}': {ex.Message}");
                                continue;
                            }
                        }

                        t.Commit();
                    }
                    catch (Exception ex)
                    {
                        t.RollBack();
                        TaskDialog.Show("Lỗi", $"Lỗi nghiêm trọng: {ex.Message}");
                        return Result.Failed;
                    }
                }

                // Hiển thị kết quả
                string resultMsg = $"IMPORT HOÀN TẤT!\n\n" +
                                   $"Thành công: {successCount}\n" +
                                   $"Lỗi: {errorLog.Count}";

                if (errorLog.Count > 0)
                {
                    resultMsg += "\n\nChi tiết lỗi:\n" + string.Join("\n", errorLog.Take(10));
                    if (errorLog.Count > 10)
                        resultMsg += $"\n... và {errorLog.Count - 10} lỗi khác";
                }

                TaskDialog.Show("Import Complete", resultMsg);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Lỗi", $"Lỗi import: {ex.Message}");
                return Result.Failed;
            }
        }

        // =============================================================================
        // COLOR PARSING
        // =============================================================================
        private Color ParseColor(string colorStr)
        {
            if (colorStr == "None" || string.IsNullOrWhiteSpace(colorStr))
                return null;

            try
            {
                string[] parts = colorStr.Split('-');
                if (parts.Length != 3)
                    return null;

                int r = int.Parse(parts[0]);
                int g = int.Parse(parts[1]);
                int b = int.Parse(parts[2]);

                if (r >= 0 && r <= 255 && g >= 0 && g <= 255 && b >= 0 && b <= 255)
                {
                    return new Color((byte)r, (byte)g, (byte)b);
                }
            }
            catch { }

            return null;
        }

        // =============================================================================
        // CSV PARSING
        // =============================================================================
        private List<string> ParseCsvLine(string line)
        {
            List<string> fields = new List<string>();
            bool inQuotes = false;
            StringBuilder currentField = new StringBuilder();

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        currentField.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    fields.Add(currentField.ToString());
                    currentField.Clear();
                }
                else
                {
                    currentField.Append(c);
                }
            }

            fields.Add(currentField.ToString());
            return fields;
        }

        private string EscapeCsvField(string field)
        {
            if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
            {
                return "\"" + field.Replace("\"", "\"\"") + "\"";
            }
            return field;
        }

        // =============================================================================
        // DATA STRUCTURES
        // =============================================================================
        private class FilterColorData
        {
            public string ViewTemplateName { get; set; }
            public string FilterName { get; set; }
            public string LineColorStr { get; set; }
            public string PatternFgColorStr { get; set; }
            public string PatternBgColorStr { get; set; }
        }

        // =============================================================================
        // MAIN FORM
        // =============================================================================
        private class FilterColorMainForm : System.Windows.Forms.Form
        {
            public string SelectedAction { get; private set; }

            public FilterColorMainForm()
            {
                Text = "Filter Colors Export/Import";
                Width = 400;
                Height = 250;
                StartPosition = FormStartPosition.CenterScreen;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;

                Label lblTitle = new Label
                {
                    Text = "Export/Import Filter Colors from View Templates",
                    Font = new System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold),
                    Location = new System.Drawing.Point(20, 20),
                    Size = new System.Drawing.Size(350, 40),
                    TextAlign = System.Drawing.ContentAlignment.MiddleCenter
                };

                Button btnExport = new Button
                {
                    Text = "📤 EXPORT FILTER COLORS",
                    Location = new System.Drawing.Point(50, 80),
                    Size = new System.Drawing.Size(280, 40),
                    BackColor = System.Drawing.Color.LightBlue
                };
                btnExport.Click += (s, e) => { SelectedAction = "Export"; DialogResult = DialogResult.OK; Close(); };

                Button btnImport = new Button
                {
                    Text = "📥 IMPORT FILTER COLORS",
                    Location = new System.Drawing.Point(50, 130),
                    Size = new System.Drawing.Size(280, 40),
                    BackColor = System.Drawing.Color.LightGreen
                };
                btnImport.Click += (s, e) => { SelectedAction = "Import"; DialogResult = DialogResult.OK; Close(); };

                Controls.Add(lblTitle);
                Controls.Add(btnExport);
                Controls.Add(btnImport);
            }
        }
    }
}

/* 
**PYREVIT → C# CONVERSIONS (263 LINES PYTHON):**
1. OverrideGraphicSettings API cho filter colors
2. Color parsing từ "R-G-B" string
3. CSV parsing với proper quote handling
4. forms.CommandSwitchWindow → Custom WinForms dialog
5. forms.alert/pick_file → TaskDialog/OpenFileDialog
6. Error logging với List<string>

**ĐÃ TUÂN THỦ:**
✅ Chuyển đổi đầy đủ 263 dòng Python
✅ Color handling chính xác (RGB values)
✅ Filter duplicate checking logic
✅ CSV escaping đúng chuẩn
✅ Transaction management proper
*/
