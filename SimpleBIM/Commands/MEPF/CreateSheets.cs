using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using WinForms = System.Windows.Forms;     // Alias cho WinForms
using Drawing = System.Drawing;          // Alias cho Drawing
using DB = Autodesk.Revit.DB;       // Alias cho Revit DB
using UI = Autodesk.Revit.UI;       // Alias cho Revit UI


namespace SimpleBIM.Commands.MEPF
{
    /// <summary>
    /// Batch Sheet Creator for Revit
    /// Tạo hàng loạt sheets tự động - FULL CONVERSION (655 lines Python)
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CreateSheets : IExternalCommand
    {
        private UIDocument _uidoc;
        private Document _doc;
        private Dictionary<string, string> _disciplines;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            _uidoc = commandData.Application.ActiveUIDocument;
            _doc = _uidoc.Document;

            _disciplines = new Dictionary<string, string>
            {
                {"Architectural", "A"}, {"Structural", "S"}, {"Mechanical", "M"},
                {"Electrical", "E"}, {"Plumbing", "P"}, {"Fire Protection", "FP"},
                {"Civil", "C"}, {"Landscape", "L"}
            };

            if (_doc.IsFamilyDocument)
            {
                TaskDialog.Show("Lỗi", "Script này chỉ chạy trong Project Document!");
                return Result.Failed;
            }

            try
            {
                return ShowMainMenu();
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Lỗi", $"Lỗi thực thi: {ex.Message}");
                return Result.Failed;
            }
        }

        // =============================================================================
        // MAIN MENU
        // =============================================================================
        private Result ShowMainMenu()
        {
            using (var form = new MainMenuForm())
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    switch (form.SelectedOption)
                    {
                        case 1: return CreateSheetsFromPrefix();
                        case 2: return CreateSheetsFromCSV();
                        case 3: return CreateSheetsFromCSVAutoNumber();
                        case 4: return CreateSheetsManual();
                        case 5: CreateSampleCSV(); return Result.Succeeded;
                        case 6: return ExportSheetsToCSV();
                    }
                }
            }
            return Result.Cancelled;
        }

        // =============================================================================
        // OPTION 1: CREATE FROM PREFIX + NUMBER
        // =============================================================================
        private Result CreateSheetsFromPrefix()
        {
            var disciplineForm = new DisciplineSelectionForm(_disciplines.Keys.ToList());
            if (disciplineForm.ShowDialog() != DialogResult.OK || disciplineForm.SelectedDiscipline == null)
                return Result.Cancelled;

            string prefix = _disciplines[disciplineForm.SelectedDiscipline];

            var inputForm = new PrefixInputForm();
            if (inputForm.ShowDialog() != DialogResult.OK)
                return Result.Cancelled;

            int startNum = inputForm.StartNumber;
            int count = inputForm.SheetCount;
            string template = inputForm.NameTemplate;

            List<SheetData> sheetsData = new List<SheetData>();
            for (int i = startNum; i < startNum + count; i++)
            {
                string sheetNumber = $"{prefix}{i:D3}";
                string sheetName = template.Replace("{n}", (i - startNum + 1).ToString());
                sheetsData.Add(new SheetData { Number = sheetNumber, Name = sheetName });
            }

            return BatchCreateSheets(sheetsData);
        }

        // =============================================================================
        // OPTION 2: IMPORT FROM CSV (WITH SHEET NUMBER)
        // =============================================================================
        private Result CreateSheetsFromCSV()
        {
            OpenFileDialog openDialog = new OpenFileDialog
            {
                Title = "Chọn file CSV",
                Filter = "CSV files (*.csv)|*.csv"
            };

            if (openDialog.ShowDialog() != DialogResult.OK)
                return Result.Cancelled;

            var sheetsData = ParseCSV(openDialog.FileName, hasSheetNumber: true);
            if (sheetsData == null || sheetsData.Count == 0)
                return Result.Cancelled;

            return BatchCreateSheets(sheetsData);
        }

        // =============================================================================
        // OPTION 3: IMPORT FROM CSV + AUTO NUMBER
        // =============================================================================
        private Result CreateSheetsFromCSVAutoNumber()
        {
            var disciplineForm = new DisciplineSelectionForm(_disciplines.Keys.ToList());
            if (disciplineForm.ShowDialog() != DialogResult.OK)
                return Result.Cancelled;

            string prefix = _disciplines[disciplineForm.SelectedDiscipline];

            var inputForm = new StartNumberForm();
            if (inputForm.ShowDialog() != DialogResult.OK)
                return Result.Cancelled;

            int startNum = inputForm.StartNumber;

            OpenFileDialog openDialog = new OpenFileDialog { Title = "Chọn file CSV", Filter = "CSV files (*.csv)|*.csv" };
            if (openDialog.ShowDialog() != DialogResult.OK)
                return Result.Cancelled;

            var sheetsData = ParseCSV(openDialog.FileName, hasSheetNumber: false);
            if (sheetsData == null) return Result.Cancelled;

            for (int i = 0; i < sheetsData.Count; i++)
            {
                sheetsData[i].Number = $"{prefix}{(startNum + i):D3}";
            }

            return BatchCreateSheets(sheetsData);
        }

        // =============================================================================
        // OPTION 4: MANUAL INPUT
        // =============================================================================
        private Result CreateSheetsManual()
        {
            List<SheetData> sheetsData = new List<SheetData>();
            
            while (true)
            {
                var inputForm = new ManualSheetInputForm(sheetsData.Count);
                if (inputForm.ShowDialog() != DialogResult.OK || string.IsNullOrWhiteSpace(inputForm.SheetNumber))
                    break;

                sheetsData.Add(new SheetData 
                { 
                    Number = inputForm.SheetNumber.Trim(), 
                    Name = inputForm.SheetName?.Trim() ?? "" 
                });
            }

            return sheetsData.Count > 0 ? BatchCreateSheets(sheetsData) : Result.Cancelled;
        }

        // =============================================================================
        // OPTION 5: CREATE SAMPLE CSV
        // =============================================================================
        private void CreateSampleCSV()
        {
            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Title = "Lưu file CSV mẫu",
                Filter = "CSV files (*.csv)|*.csv",
                FileName = "Sheets_Sample.csv"
            };

            if (saveDialog.ShowDialog() != DialogResult.OK)
                return;

            using (StreamWriter writer = new StreamWriter(saveDialog.FileName, false, Encoding.UTF8))
            {
                writer.WriteLine("Sheet Number,Sheet Name");
                writer.WriteLine("A101,ARCHITECTURAL FLOOR PLAN - LEVEL 1");
                writer.WriteLine("A102,ARCHITECTURAL FLOOR PLAN - LEVEL 2");
                writer.WriteLine("A103,ARCHITECTURAL CEILING PLAN - LEVEL 1");
                writer.WriteLine("S201,STRUCTURAL PLAN - LEVEL 1");
                writer.WriteLine("S202,STRUCTURAL PLAN - LEVEL 2");
                writer.WriteLine("M301,MECHANICAL PLAN - LEVEL 1");
                writer.WriteLine("E401,ELECTRICAL PLAN - LEVEL 1");
                writer.WriteLine("P501,PLUMBING PLAN - LEVEL 1");
            }

            TaskDialog.Show("Thành công", $"Đã tạo file CSV mẫu:\n{Path.GetFileName(saveDialog.FileName)}");
        }

        // =============================================================================
        // OPTION 6: EXPORT SHEETS TO CSV
        // =============================================================================
        private Result ExportSheetsToCSV()
        {
            FilteredElementCollector collector = new FilteredElementCollector(_doc).OfClass(typeof(ViewSheet));
            List<ViewSheet> sheets = collector.Cast<ViewSheet>().Where(s => !s.IsPlaceholder).ToList();

            if (sheets.Count == 0)
            {
                TaskDialog.Show("Thông báo", "Không có sheets nào trong project!");
                return Result.Cancelled;
            }

            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Title = "Lưu danh sách sheets ra CSV",
                Filter = "CSV files (*.csv)|*.csv",
                FileName = $"Sheets_Export_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
            };

            if (saveDialog.ShowDialog() != DialogResult.OK)
                return Result.Cancelled;

            using (StreamWriter writer = new StreamWriter(saveDialog.FileName, false, Encoding.UTF8))
            {
                writer.WriteLine("Sheet Number,Sheet Name,Titleblock");
                
                var sortedSheets = sheets.OrderBy(s => s.SheetNumber).ToList();
                foreach (var sheet in sortedSheets)
                {
                    string titleblock = "No Titleblock";
                    try
                    {
                        ElementId tbId = sheet.GetAllViewports().FirstOrDefault();
                        if (tbId != null)
                        {
                            Element tb = _doc.GetElement(tbId);
                            titleblock = tb?.Name ?? "Unknown";
                        }
                    }
                    catch { }

                    writer.WriteLine($"{EscapeCsvField(sheet.SheetNumber)},{EscapeCsvField(sheet.Name)},{EscapeCsvField(titleblock)}");
                }
            }

            TaskDialog.Show("Thành công", $"Đã xuất {sheets.Count} sheets ra file:\n{Path.GetFileName(saveDialog.FileName)}");
            return Result.Succeeded;
        }

        // =============================================================================
        // BATCH CREATE SHEETS
        // =============================================================================
        private Result BatchCreateSheets(List<SheetData> sheetsData)
        {
            if (sheetsData == null || sheetsData.Count == 0)
            {
                TaskDialog.Show("Lỗi", "Không có dữ liệu sheets!");
                return Result.Cancelled;
            }

            // Get titleblocks
            Dictionary<string, ElementId> titleblocks = GetAvailableTitleblocks();
            if (titleblocks.Count == 0)
            {
                TaskDialog.Show("Lỗi", "Không tìm thấy Titleblock nào!\nHãy load Titleblock family trước.");
                return Result.Failed;
            }

            var titleblockForm = new TitleblockSelectionForm(titleblocks.Keys.ToList());
            if (titleblockForm.ShowDialog() != DialogResult.OK)
                return Result.Cancelled;

            ElementId titleblockId = titleblocks[titleblockForm.SelectedTitleblock];

            // Preview
            string preview = $"📋 SẼ TẠO {sheetsData.Count} SHEETS:\n\n";
            preview += $"Titleblock: {titleblockForm.SelectedTitleblock}\n\n";
            for (int i = 0; i < Math.Min(10, sheetsData.Count); i++)
            {
                preview += $"{i + 1}. {sheetsData[i].Number} - {sheetsData[i].Name}\n";
            }
            if (sheetsData.Count > 10)
                preview += $"... và {sheetsData.Count - 10} sheets khác\n";

            var confirmDialog = new TaskDialog("Xác nhận")
            {
                MainContent = preview + "\n\nTiếp tục tạo sheets?",
                CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
            };

            if (confirmDialog.Show() != TaskDialogResult.Yes)
                return Result.Cancelled;

            // Create sheets
            int successCount = 0;
            int errorCount = 0;
            List<string> errors = new List<string>();

            using (Transaction trans = new Transaction(_doc, $"Batch Create {sheetsData.Count} Sheets"))
            {
                trans.Start();

                try
                {
                    foreach (var sheetData in sheetsData)
                    {
                        try
                        {
                            var (sheet, error) = CreateSheet(sheetData.Number, sheetData.Name, titleblockId);
                            if (sheet != null)
                                successCount++;
                            else
                            {
                                errorCount++;
                                errors.Add($"{sheetData.Number}: {error}");
                            }
                        }
                        catch (Exception ex)
                        {
                            errorCount++;
                            errors.Add($"{sheetData.Number}: {ex.Message}");
                        }
                    }

                    trans.Commit();
                }
                catch (Exception ex)
                {
                    trans.RollBack();
                    TaskDialog.Show("Lỗi", $"Lỗi transaction: {ex.Message}");
                    return Result.Failed;
                }
            }

            // Report
            string result = $"🎉 HOÀN THÀNH TẠO SHEETS!\n\n";
            result += $"✅ Thành công: {successCount} sheets\n";
            if (errorCount > 0)
            {
                result += $"❌ Lỗi: {errorCount} sheets\n\n";
                if (errors.Count <= 5)
                    result += "Chi tiết lỗi:\n" + string.Join("\n", errors);
            }
            result += "\n\n💾 Hãy Save file Revit!";

            TaskDialog.Show("Kết quả", result);
            return Result.Succeeded;
        }

        // =============================================================================
        // HELPER FUNCTIONS
        // =============================================================================
        private (ViewSheet, string) CreateSheet(string sheetNumber, string sheetName, ElementId titleblockId)
        {
            try
            {
                // Check duplicate
                FilteredElementCollector collector = new FilteredElementCollector(_doc).OfClass(typeof(ViewSheet));
                foreach (ViewSheet existingSheet in collector)
                {
                    if (existingSheet.SheetNumber == sheetNumber)
                        return (null, $"Sheet number '{sheetNumber}' đã tồn tại");
                }

                // Clean name
                string cleanedName = CleanSheetName(sheetName);
                if (string.IsNullOrWhiteSpace(cleanedName))
                    cleanedName = "Sheet";

                ViewSheet newSheet = ViewSheet.Create(_doc, titleblockId);
                newSheet.SheetNumber = sheetNumber;
                newSheet.Name = cleanedName;

                return (newSheet, null);
            }
            catch (Exception ex)
            {
                return (null, ex.Message);
            }
        }

        private string CleanSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "";

            string cleaned = Regex.Replace(name, @"[\\{}|<>*?/:""]", " ");
            cleaned = Regex.Replace(cleaned, @"\s+", " ").Trim();
            return cleaned.Length > 250 ? cleaned.Substring(0, 250) + "..." : cleaned;
        }

        private Dictionary<string, ElementId> GetAvailableTitleblocks()
        {
            Dictionary<string, ElementId> titleblocks = new Dictionary<string, ElementId>();
            FilteredElementCollector collector = new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .WhereElementIsElementType();

            foreach (Element tb in collector)
            {
                string familyName = (tb as FamilySymbol)?.FamilyName ?? "Unknown Family";
                string typeName = tb.Name ?? "Unknown Type";
                string key = $"{familyName} - {typeName}";
                titleblocks[key] = tb.Id;
            }

            return titleblocks;
        }

        private List<SheetData> ParseCSV(string filePath, bool hasSheetNumber)
        {
            List<SheetData> data = new List<SheetData>();

            try
            {
                using (StreamReader reader = new StreamReader(filePath, Encoding.UTF8))
                {
                    string headerLine = reader.ReadLine(); // Skip header

                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        var parts = ParseCsvLine(line);
                        if (parts.Count == 0 || string.IsNullOrWhiteSpace(parts[0]))
                            continue;

                        if (hasSheetNumber && parts.Count >= 2)
                        {
                            data.Add(new SheetData
                            {
                                Number = FixVietnameseText(parts[0].Trim()),
                                Name = FixVietnameseText(parts[1].Trim())
                            });
                        }
                        else if (!hasSheetNumber)
                        {
                            data.Add(new SheetData
                            {
                                Number = "", // Will be set later
                                Name = FixVietnameseText(parts[0].Trim())
                            });
                        }
                    }
                }

                return data;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Lỗi", $"Lỗi đọc CSV: {ex.Message}");
                return null;
            }
        }

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
                        inQuotes = !inQuotes;
                }
                else if (c == ',' && !inQuotes)
                {
                    fields.Add(currentField.ToString());
                    currentField.Clear();
                }
                else
                    currentField.Append(c);
            }
            fields.Add(currentField.ToString());
            return fields;
        }

        private string EscapeCsvField(string field)
        {
            if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
                return "\"" + field.Replace("\"", "\"\"") + "\"";
            return field;
        }

        private string FixVietnameseText(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            text = text.Replace("T7NG", "TẦNG").Replace("T?NG", "TẦNG");
            text = text.Replace("THÓNG", "THÔNG").Replace("TH?NG", "THÔNG");
            return text;
        }

        // =============================================================================
        // DATA STRUCTURES
        // =============================================================================
        private class SheetData
        {
            public string Number { get; set; }
            public string Name { get; set; }
        }

        // =============================================================================
        // FORMS
        // =============================================================================
        private class MainMenuForm : System.Windows.Forms.Form
        {
            public int SelectedOption { get; private set; }

            public MainMenuForm()
            {
                Text = "Batch Sheet Creator";
                Width = 500;
                Height = 400;
                StartPosition = FormStartPosition.CenterScreen;

                WinForms.ListBox listBox = new ListBox
                {
                    Location = new System.Drawing.Point(20, 20),
                    Size = new System.Drawing.Size(440, 280)
                };

                listBox.Items.Add("🔤 Tạo sheets với PREFIX + SỐ THỨ TỰ");
                listBox.Items.Add("📊 Import từ CSV (có sẵn Sheet Number)");
                listBox.Items.Add("🎯 Import từ CSV + TỰ ĐỘNG ĐÁNH SỐ");
                listBox.Items.Add("✍️ Nhập thủ công từng sheet");
                listBox.Items.Add("📝 Tạo file CSV mẫu");
                listBox.Items.Add("📤 XUẤT sheets hiện có ra CSV");

                WinForms.Button btnOK = new Button
                {
                    Text = "CHỌN",
                    Location = new System.Drawing.Point(280, 320),
                    Size = new System.Drawing.Size(90, 30)
                };
                btnOK.Click += (s, e) =>
                {
                    if (listBox.SelectedIndex >= 0)
                    {
                        SelectedOption = listBox.SelectedIndex + 1;
                        DialogResult = DialogResult.OK;
                        Close();
                    }
                };

                WinForms.Button btnCancel = new Button
                {
                    Text = "HỦY",
                    Location = new System.Drawing.Point(380, 320),
                    Size = new System.Drawing.Size(80, 30)
                };
                btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

                Controls.Add(listBox);
                Controls.Add(btnOK);
                Controls.Add(btnCancel);
            }
        }

        private class DisciplineSelectionForm : System.Windows.Forms.Form
        {
            public string SelectedDiscipline { get; private set; }

            public DisciplineSelectionForm(List<string> disciplines)
            {
                Text = "Chọn Discipline";
                Width = 350;
                Height = 300;
                StartPosition = FormStartPosition.CenterScreen;

                WinForms.ListBox listBox = new ListBox
                {
                    Location = new System.Drawing.Point(20, 20),
                    Size = new System.Drawing.Size(290, 180)
                };

                foreach (var disc in disciplines.OrderBy(x => x))
                    listBox.Items.Add(disc);

                WinForms.Button btnOK = new Button { Text = "OK", Location = new System.Drawing.Point(150, 220), Size = new System.Drawing.Size(80, 30) };
                btnOK.Click += (s, e) =>
                {
                    if (listBox.SelectedItem != null)
                    {
                        SelectedDiscipline = listBox.SelectedItem.ToString();
                        DialogResult = DialogResult.OK;
                        Close();
                    }
                };

                WinForms.Button btnCancel = new Button { Text = "Cancel", Location = new System.Drawing.Point(240, 220), Size = new System.Drawing.Size(80, 30) };
                btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

                Controls.Add(listBox);
                Controls.Add(btnOK);
                Controls.Add(btnCancel);
            }
        }

        private class PrefixInputForm : System.Windows.Forms.Form
        {
            public int StartNumber { get; private set; }
            public int SheetCount { get; private set; }
            public string NameTemplate { get; private set; }

            public PrefixInputForm()
            {
                Text = "Nhập thông tin sheets";
                Width = 400;
                Height = 250;
                StartPosition = FormStartPosition.CenterScreen;

                WinForms.Label lbl1 = new Label { Text = "Số bắt đầu:", Location = new System.Drawing.Point(20, 20), Width = 100 };
                WinForms.TextBox txt1 = new WinForms.TextBox { Text = "100", Location = new System.Drawing.Point(130, 20), Width = 200 };

                WinForms.Label lbl2 = new Label { Text = "Số lượng sheets:", Location = new System.Drawing.Point(20, 60), Width = 100 };
                WinForms.TextBox txt2 = new WinForms.TextBox { Text = "10", Location = new System.Drawing.Point(130, 60), Width = 200 };

                WinForms.Label lbl3 = new Label { Text = "Tên sheet (dùng {n}):", Location = new System.Drawing.Point(20, 100), Width = 350 };
                WinForms.TextBox txt3 = new WinForms.TextBox { Text = "FLOOR PLAN - LEVEL {n}", Location = new System.Drawing.Point(20, 125), Width = 350 };

                WinForms.Button btnOK = new Button { Text = "OK", Location = new System.Drawing.Point(200, 170), Size = new System.Drawing.Size(80, 30) };
                btnOK.Click += (s, e) =>
                {
                    if (int.TryParse(txt1.Text, out int start) && int.TryParse(txt2.Text, out int count))
                    {
                        StartNumber = start;
                        SheetCount = count;
                        NameTemplate = txt3.Text;
                        DialogResult = DialogResult.OK;
                        Close();
                    }
                    else
                        WinForms.MessageBox.Show("Số không hợp lệ!");
                };

                WinForms.Button btnCancel = new Button { Text = "Cancel", Location = new System.Drawing.Point(290, 170), Size = new System.Drawing.Size(80, 30) };
                btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

                Controls.AddRange(new WinForms.Control[] { lbl1, txt1, lbl2, txt2, lbl3, txt3, btnOK, btnCancel });
            }
        }

        private class StartNumberForm : System.Windows.Forms.Form
        {
            public int StartNumber { get; private set; }

            public StartNumberForm()
            {
                Text = "Nhập số bắt đầu";
                Width = 300;
                Height = 150;
                StartPosition = FormStartPosition.CenterScreen;

                WinForms.Label lbl = new Label { Text = "Số bắt đầu:", Location = new System.Drawing.Point(20, 20), Width = 100 };
                WinForms.TextBox txt = new WinForms.TextBox { Text = "100", Location = new System.Drawing.Point(130, 20), Width = 120 };

                WinForms.Button btnOK = new Button { Text = "OK", Location = new System.Drawing.Point(100, 70), Size = new System.Drawing.Size(80, 30) };
                btnOK.Click += (s, e) =>
                {
                    if (int.TryParse(txt.Text, out int num))
                    {
                        StartNumber = num;
                        DialogResult = DialogResult.OK;
                        Close();
                    }
                };

                WinForms.Button btnCancel = new Button { Text = "Cancel", Location = new System.Drawing.Point(190, 70), Size = new System.Drawing.Size(80, 30) };
                btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

                Controls.AddRange(new WinForms.Control[] { lbl, txt, btnOK, btnCancel });
            }
        }

        private class ManualSheetInputForm : System.Windows.Forms.Form
        {
            public string SheetNumber { get; private set; }
            public string SheetName { get; private set; }

            public ManualSheetInputForm(int currentCount)
            {
                Text = $"Nhập Sheet #{currentCount + 1}";
                Width = 400;
                Height = 200;
                StartPosition = FormStartPosition.CenterScreen;

                WinForms.Label lbl1 = new Label { Text = $"Sheet Number (đã nhập {currentCount}):", Location = new System.Drawing.Point(20, 20), Width = 350 };
                WinForms.TextBox txt1 = new WinForms.TextBox { Location = new System.Drawing.Point(20, 45), Width = 350 };

                WinForms.Label lbl2 = new Label { Text = "Sheet Name:", Location = new System.Drawing.Point(20, 80), Width = 350 };
                WinForms.TextBox txt2 = new WinForms.TextBox { Location = new System.Drawing.Point(20, 105), Width = 350 };

                WinForms.Button btnOK = new Button { Text = "THÊM", Location = new System.Drawing.Point(200, 140), Size = new System.Drawing.Size(80, 30) };
                btnOK.Click += (s, e) =>
                {
                    SheetNumber = txt1.Text;
                    SheetName = txt2.Text;
                    DialogResult = DialogResult.OK;
                    Close();
                };

                WinForms.Button btnCancel = new Button { Text = "HOÀN TẤT", Location = new System.Drawing.Point(290, 140), Size = new System.Drawing.Size(80, 30) };
                btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

                Controls.AddRange(new WinForms.Control[] { lbl1, txt1, lbl2, txt2, btnOK, btnCancel });
            }
        }

        private class TitleblockSelectionForm : System.Windows.Forms.Form
        {
            public string SelectedTitleblock { get; private set; }

            public TitleblockSelectionForm(List<string> titleblocks)
            {
                Text = "Chọn Titleblock";
                Width = 450;
                Height = 350;
                StartPosition = FormStartPosition.CenterScreen;

                WinForms.ListBox listBox = new ListBox
                {
                    Location = new System.Drawing.Point(20, 20),
                    Size = new System.Drawing.Size(390, 240)
                };

                foreach (var tb in titleblocks.OrderBy(x => x))
                    listBox.Items.Add(tb);

                if (listBox.Items.Count > 0)
                    listBox.SelectedIndex = 0;

                WinForms.Button btnOK = new Button { Text = "CHỌN", Location = new System.Drawing.Point(240, 280), Size = new System.Drawing.Size(80, 30) };
                btnOK.Click += (s, e) =>
                {
                    if (listBox.SelectedItem != null)
                    {
                        SelectedTitleblock = listBox.SelectedItem.ToString();
                        DialogResult = DialogResult.OK;
                        Close();
                    }
                };

                WinForms.Button btnCancel = new Button { Text = "HỦY", Location = new System.Drawing.Point(330, 280), Size = new System.Drawing.Size(80, 30) };
                btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

                Controls.Add(listBox);
                Controls.Add(btnOK);
                Controls.Add(btnCancel);
            }
        }
    }
}

/* 
**PYREVIT → C# CONVERSIONS (655 LINES PYTHON):**
✅ 6 menu options với full logic
✅ CSV parsing với Vietnamese text fixing
✅ Sheet duplicate detection
✅ Titleblock selection
✅ Batch creation với transaction
✅ Error reporting đầy đủ
✅ Export sheets to CSV
✅ Sample CSV generation
*/
