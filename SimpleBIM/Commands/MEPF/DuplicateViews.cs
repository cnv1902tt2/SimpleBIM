using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace SimpleBIM.Commands.MEPF
{
    /// <summary>
    /// Batch View Duplicator - FULL CONVERSION (304 lines Python)
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class DuplicateViews : IExternalCommand
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
                {"Electrical", "E"}, {"Plumbing", "P"}, {"Civil", "C"}
            };

            if (_doc.IsFamilyDocument)
            {
                TaskDialog.Show("Lỗi", "Script chỉ chạy trong Project!");
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

        private Result ShowMainMenu()
        {
            using (var form = new MainMenuForm())
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    switch (form.SelectedOption)
                    {
                        case 1: return ImportFromCSV();
                        case 2: return ExportToCSV();
                        case 3: CreateSampleCSV(); return Result.Succeeded;
                    }
                }
            }
            return Result.Cancelled;
        }

        // =============================================================================
        // EXPORT VIEWS TO CSV
        // =============================================================================
        private Result ExportToCSV()
        {
            try
            {
                FilteredElementCollector collector = new FilteredElementCollector(_doc).OfClass(typeof(Autodesk.Revit.DB.View));
                List<ViewData> floorPlans = new List<ViewData>();

                foreach (Autodesk.Revit.DB.View v in collector)
                {
                    if (v.ViewType == ViewType.FloorPlan && !v.IsTemplate && v.CanBePrinted)
                    {
                        string levelName = v.GenLevel?.Name ?? "";
                        floorPlans.Add(new ViewData { Name = v.Name, Level = levelName });
                    }
                }

                if (floorPlans.Count == 0)
                {
                    TaskDialog.Show("Thông báo", "Không tìm thấy Floor Plan views nào!");
                    return Result.Cancelled;
                }

                SaveFileDialog saveDialog = new SaveFileDialog
                {
                    Title = "Xuất views ra CSV",
                    Filter = "CSV files (*.csv)|*.csv",
                    FileName = "Current_Views_Export.csv"
                };

                if (saveDialog.ShowDialog() != DialogResult.OK)
                    return Result.Cancelled;

                using (StreamWriter writer = new StreamWriter(saveDialog.FileName, false, Encoding.UTF8))
                {
                    writer.WriteLine("View Name,Level,New View Name,Discipline");
                    foreach (var view in floorPlans)
                    {
                        writer.WriteLine($"{view.Name},{view.Level},,");
                    }
                }

                TaskDialog.Show("Thành công", $"Đã xuất {floorPlans.Count} views ra file:\n{Path.GetFileName(saveDialog.FileName)}");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Lỗi", $"Lỗi xuất CSV: {ex.Message}");
                return Result.Failed;
            }
        }

        // =============================================================================
        // IMPORT FROM CSV & DUPLICATE
        // =============================================================================
        private Result ImportFromCSV()
        {
            // Step 1: Select duplicate mode
            using (var modeForm = new DuplicateModeForm())
            {
                if (modeForm.ShowDialog() != DialogResult.OK)
                    return Result.Cancelled;

                ViewDuplicateOption duplicateOption = modeForm.SelectedOption;

                // Step 2: Select disciplines
                var disciplineForm = new DisciplineMultiSelectForm(_disciplines.Keys.ToList());
                if (disciplineForm.ShowDialog() != DialogResult.OK || disciplineForm.SelectedDisciplines.Count == 0)
                    return Result.Cancelled;

                List<string> selectedDisciplines = disciplineForm.SelectedDisciplines;

                // Step 3: Import CSV
                OpenFileDialog openDialog = new OpenFileDialog
                {
                    Title = "Chọn file CSV",
                    Filter = "CSV files (*.csv)|*.csv"
                };

                if (openDialog.ShowDialog() != DialogResult.OK)
                    return Result.Cancelled;

                List<ViewDuplicateData> viewsData = ParseCSV(openDialog.FileName);
                if (viewsData == null || viewsData.Count == 0)
                    return Result.Cancelled;

                // Step 4: Batch duplicate
                return BatchDuplicate(viewsData, selectedDisciplines, duplicateOption);
            }
        }

        private Result BatchDuplicate(List<ViewDuplicateData> viewsData, List<string> disciplines, ViewDuplicateOption option)
        {
            int totalCreated = 0;
            int errorCount = 0;

            using (Transaction trans = new Transaction(_doc, "Batch Duplicate Views"))
            {
                trans.Start();

                try
                {
                    foreach (var viewData in viewsData)
                    {
                        Autodesk.Revit.DB.View sourceView = FindViewByName(viewData.ViewName);
                        if (sourceView == null)
                        {
                            errorCount++;
                            continue;
                        }

                        // Determine disciplines
                        List<string> disciplinesToUse = disciplines;
                        if (!string.IsNullOrEmpty(viewData.SpecificDiscipline) && _disciplines.ContainsKey(viewData.SpecificDiscipline))
                        {
                            disciplinesToUse = new List<string> { viewData.SpecificDiscipline };
                        }

                        foreach (string discipline in disciplinesToUse)
                        {
                            bool success = DuplicateView(sourceView, discipline, option, viewData.CustomName, viewData.Level);
                            if (success)
                                totalCreated++;
                            else
                                errorCount++;
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

            TaskDialog.Show("Hoàn thành", $"🎉 Đã tạo thành công {totalCreated} views!\n❌ Lỗi: {errorCount}");
            return Result.Succeeded;
        }

        private bool DuplicateView(Autodesk.Revit.DB.View sourceView, string discipline, ViewDuplicateOption option, string customName = "", string levelName = "")
        {
            try
            {
                string newViewName;
                if (!string.IsNullOrWhiteSpace(customName))
                {
                    newViewName = customName;
                }
                else
                {
                    string prefix = _disciplines.ContainsKey(discipline) ? _disciplines[discipline] : "X";
                    newViewName = $"{prefix} - {sourceView.Name} - {levelName}";
                }

                // Clean name
                newViewName = System.Text.RegularExpressions.Regex.Replace(newViewName, @"[\\{}|<>*?/:""]", " ");
                newViewName = System.Text.RegularExpressions.Regex.Replace(newViewName, @"\s+", " ").Trim();
                if (newViewName.Length > 250)
                    newViewName = newViewName.Substring(0, 250);

                ElementId newViewId = sourceView.Duplicate(option);
                Autodesk.Revit.DB.View newView = _doc.GetElement(newViewId) as Autodesk.Revit.DB.View;
                if (newView != null)
                {
                    newView.Name = newViewName;
                    return true;
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        private Autodesk.Revit.DB.View FindViewByName(string viewName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(_doc).OfClass(typeof(Autodesk.Revit.DB.View));
            foreach (Autodesk.Revit.DB.View v in collector)
            {
                if (v.ViewType == ViewType.FloorPlan && !v.IsTemplate && v.Name.Trim() == viewName.Trim())
                    return v;
            }
            return null;
        }

        private List<ViewDuplicateData> ParseCSV(string filePath)
        {
            List<ViewDuplicateData> data = new List<ViewDuplicateData>();

            try
            {
                using (StreamReader reader = new StreamReader(filePath, Encoding.UTF8))
                {
                    string headerLine = reader.ReadLine(); // Skip header

                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        var parts = ParseCsvLine(line);
                        if (parts.Count < 2 || string.IsNullOrWhiteSpace(parts[0]))
                            continue;

                        data.Add(new ViewDuplicateData
                        {
                            ViewName = parts[0].Trim(),
                            Level = parts.Count > 1 ? parts[1].Trim() : "",
                            CustomName = parts.Count > 2 ? parts[2].Trim() : "",
                            SpecificDiscipline = parts.Count > 3 ? parts[3].Trim() : ""
                        });
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

        private void CreateSampleCSV()
        {
            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Title = "Tạo file CSV mẫu",
                Filter = "CSV files (*.csv)|*.csv",
                FileName = "Revit_Views_Template.csv"
            };

            if (saveDialog.ShowDialog() != DialogResult.OK)
                return;

            using (StreamWriter writer = new StreamWriter(saveDialog.FileName, false, Encoding.UTF8))
            {
                writer.WriteLine("View Name,Level,New View Name,Discipline");
                writer.WriteLine("L1 - Architectural,L1,MẶT BẰNG KIẾN TRÚC TẦNG 1,Architectural");
                writer.WriteLine("L2 - Architectural,L2,MẶT BẰNG KIẾN TRÚC TẦNG 2,Architectural");
            }

            TaskDialog.Show("Thành công", $"Đã tạo file mẫu: {Path.GetFileName(saveDialog.FileName)}");
        }

        // =============================================================================
        // DATA STRUCTURES
        // =============================================================================
        private class ViewData
        {
            public string Name { get; set; }
            public string Level { get; set; }
        }

        private class ViewDuplicateData
        {
            public string ViewName { get; set; }
            public string Level { get; set; }
            public string CustomName { get; set; }
            public string SpecificDiscipline { get; set; }
        }

        // =============================================================================
        // FORMS
        // =============================================================================
        private class MainMenuForm : System.Windows.Forms.Form
        {
            public int SelectedOption { get; private set; }

            public MainMenuForm()
            {
                Text = "Batch View Duplicator";
                Width = 450;
                Height = 250;
                StartPosition = FormStartPosition.CenterScreen;

                ListBox listBox = new ListBox
                {
                    Location = new System.Drawing.Point(20, 20),
                    Size = new System.Drawing.Size(390, 120)
                };

                listBox.Items.Add("🎯 IMPORT VIEWS TỪ CSV VỚI TÊN TÙY CHỈNH");
                listBox.Items.Add("📤 EXPORT VIEWS HIỆN CÓ RA CSV");
                listBox.Items.Add("📝 TẠO FILE CSV MẪU");

                Button btnOK = new Button { Text = "CHỌN", Location = new System.Drawing.Point(240, 160), Size = new System.Drawing.Size(80, 30) };
                btnOK.Click += (s, e) =>
                {
                    if (listBox.SelectedIndex >= 0)
                    {
                        SelectedOption = listBox.SelectedIndex + 1;
                        DialogResult = DialogResult.OK;
                        Close();
                    }
                };

                Button btnCancel = new Button { Text = "HỦY", Location = new System.Drawing.Point(330, 160), Size = new System.Drawing.Size(80, 30) };
                btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

                Controls.Add(listBox);
                Controls.Add(btnOK);
                Controls.Add(btnCancel);
            }
        }

        private class DuplicateModeForm : System.Windows.Forms.Form
        {
            public ViewDuplicateOption SelectedOption { get; private set; }

            public DuplicateModeForm()
            {
                Text = "Chọn chế độ Duplicate View";
                Width = 350;
                Height = 250;
                StartPosition = FormStartPosition.CenterScreen;

                GroupBox groupBox = new GroupBox
                {
                    Text = "Duplicate Options",
                    Location = new System.Drawing.Point(20, 20),
                    Size = new System.Drawing.Size(290, 130)
                };

                RadioButton radio1 = new RadioButton { Text = "Duplicate", Location = new System.Drawing.Point(20, 30), Width = 250, Checked = true };
                RadioButton radio2 = new RadioButton { Text = "Duplicate with Detailing", Location = new System.Drawing.Point(20, 60), Width = 250 };
                RadioButton radio3 = new RadioButton { Text = "Duplicate as Dependent", Location = new System.Drawing.Point(20, 90), Width = 250 };

                groupBox.Controls.Add(radio1);
                groupBox.Controls.Add(radio2);
                groupBox.Controls.Add(radio3);

                Button btnOK = new Button { Text = "OK", Location = new System.Drawing.Point(150, 170), Size = new System.Drawing.Size(80, 30) };
                btnOK.Click += (s, e) =>
                {
                    if (radio1.Checked) SelectedOption = ViewDuplicateOption.Duplicate;
                    else if (radio2.Checked) SelectedOption = ViewDuplicateOption.WithDetailing;
                    else if (radio3.Checked) SelectedOption = ViewDuplicateOption.AsDependent;
                    DialogResult = DialogResult.OK;
                    Close();
                };

                Button btnCancel = new Button { Text = "Cancel", Location = new System.Drawing.Point(240, 170), Size = new System.Drawing.Size(80, 30) };
                btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

                Controls.Add(groupBox);
                Controls.Add(btnOK);
                Controls.Add(btnCancel);
            }
        }

        private class DisciplineMultiSelectForm : System.Windows.Forms.Form
        {
            public List<string> SelectedDisciplines { get; private set; }

            public DisciplineMultiSelectForm(List<string> disciplines)
            {
                Text = "Chọn Disciplines";
                Width = 350;
                Height = 350;
                StartPosition = FormStartPosition.CenterScreen;
                SelectedDisciplines = new List<string>();

                CheckedListBox listBox = new CheckedListBox
                {
                    Location = new System.Drawing.Point(20, 20),
                    Size = new System.Drawing.Size(290, 240),
                    CheckOnClick = true
                };

                foreach (var disc in disciplines.OrderBy(x => x))
                    listBox.Items.Add(disc);

                Button btnOK = new Button { Text = "OK", Location = new System.Drawing.Point(150, 280), Size = new System.Drawing.Size(80, 30) };
                btnOK.Click += (s, e) =>
                {
                    foreach (var item in listBox.CheckedItems)
                        SelectedDisciplines.Add(item.ToString());
                    
                    if (SelectedDisciplines.Count > 0)
                    {
                        DialogResult = DialogResult.OK;
                        Close();
                    }
                    else
                        MessageBox.Show("Vui lòng chọn ít nhất 1 discipline!");
                };

                Button btnCancel = new Button { Text = "Cancel", Location = new System.Drawing.Point(240, 280), Size = new System.Drawing.Size(80, 30) };
                btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

                Controls.Add(listBox);
                Controls.Add(btnOK);
                Controls.Add(btnCancel);
            }
        }
    }
}

/* 
**PYREVIT → C# CONVERSIONS (304 LINES PYTHON):**
✅ 3 duplicate modes: Duplicate, WithDetailing, AsDependent
✅ Multi-discipline selection với CheckedListBox
✅ CSV import/export
✅ Custom view names với template
✅ Batch duplication với transaction
✅ Error handling đầy đủ
*/
