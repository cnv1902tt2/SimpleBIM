using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Diagnostics;

// WINDOWS FORMS
using System.Windows.Forms;
using System.Drawing;

// REVIT API
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Exceptions;

using OperationCanceledException = Autodesk.Revit.Exceptions.OperationCanceledException;

namespace SimpleBIM.Commands.Qs
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ExportSchedules : IExternalCommand
    {
        private class ScheduleInfo
        {
            public string Name { get; set; }
            public List<string> Columns { get; set; } = new List<string>();
            public int RowCount { get; set; }
            public bool HasData { get; set; }
            public ViewSchedule Schedule { get; set; }
            public string Error { get; set; }
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Lấy tất cả schedules
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                List<ViewSchedule> allSchedules = collector.OfClass(typeof(ViewSchedule))
                    .Cast<ViewSchedule>()
                    .ToList();

                if (allSchedules.Count == 0)
                {
                    TaskDialog.Show("Warning", "Không có Schedule nào trong dự án!");
                    return Result.Cancelled;
                }

                // Lấy thông tin schedule
                List<ScheduleInfo> schedulesInfo = new List<ScheduleInfo>();
                foreach (ViewSchedule schedule in allSchedules)
                {
                    try
                    {
                        schedulesInfo.Add(GetScheduleInfo(schedule));
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error getting schedule info: {ex.Message}");
                        continue;
                    }
                }

                // Hiển thị form chọn schedule
                using (ScheduleSelectionForm selectionForm = new ScheduleSelectionForm(schedulesInfo))
                {
                    if (selectionForm.ShowDialog() != DialogResult.OK)
                        return Result.Cancelled;

                    List<ScheduleInfo> selectedSchedules = selectionForm.GetSelectedSchedules();
                    if (selectedSchedules.Count == 0)
                    {
                        TaskDialog.Show("Warning", "Chưa chọn schedule nào!");
                        return Result.Cancelled;
                    }

                    // Chọn thư mục lưu
                    using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
                    {
                        folderDialog.Description = "Chọn thư mục để lưu file CSV";
                        if (folderDialog.ShowDialog() != DialogResult.OK)
                            return Result.Cancelled;

                        string outputFolder = folderDialog.SelectedPath;

                        // Export schedules
                        int successCount = 0;
                        List<string> errors = new List<string>();

                        foreach (ScheduleInfo info in selectedSchedules)
                        {
                            try
                            {
                                string fileName = SanitizeFileName(info.Name);
                                string filePath = Path.Combine(outputFolder, $"{fileName}.csv");
                                ExportScheduleToCsv(info.Schedule, filePath);
                                successCount++;
                            }
                            catch (Exception ex)
                            {
                                errors.Add($"{info.Name}: {ex.Message}");
                            }
                        }

                        // Kết quả
                        string resultMsg = $"Đã xuất {successCount} schedule.";
                        if (errors.Count > 0)
                        {
                            resultMsg += $"\n\nLỗi: {errors.Count}";
                            foreach (string err in errors) Debug.WriteLine($"- {err}");
                        }
                        TaskDialog.Show("Export Complete", resultMsg);
                    }
                }
                return Result.Succeeded;
            }
            catch (OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                Debug.WriteLine($"Error in ExportSchedules: {ex.Message}\n{ex.StackTrace}");
                return Result.Failed;
            }
        }

        private ScheduleInfo GetScheduleInfo(ViewSchedule schedule)
        {
            try
            {
                string name = schedule.Name;
                ScheduleDefinition definition = schedule.Definition;
                int fieldCount = definition.GetFieldCount();

                List<string> columns = new List<string>();
                for (int i = 0; i < fieldCount; i++)
                {
                    try
                    {
                        columns.Add(definition.GetField(i).GetName());
                    }
                    catch
                    {
                        columns.Add($"Column_{i + 1}");
                    }
                }

                TableData tableData = schedule.GetTableData();
                TableSectionData bodyData = tableData.GetSectionData(SectionType.Body);
                int rowCount = bodyData != null ? bodyData.NumberOfRows : 0;

                return new ScheduleInfo
                {
                    Name = name,
                    Columns = columns,
                    RowCount = rowCount,
                    HasData = rowCount > 0,
                    Schedule = schedule
                };
            }
            catch (Exception ex)
            {
                return new ScheduleInfo
                {
                    Name = schedule.Name,
                    Columns = new List<string>(),
                    RowCount = 0,
                    HasData = false,
                    Schedule = schedule,
                    Error = ex.Message
                };
            }
        }

        private void ExportScheduleToCsv(ViewSchedule schedule, string filePath)
        {
            try
            {
                ScheduleDefinition definition = schedule.Definition;
                int fieldCount = definition.GetFieldCount();

                using (StreamWriter sw = new StreamWriter(filePath, false, Encoding.UTF8))
                {
                    // Header
                    List<string> headers = new List<string>();
                    for (int i = 0; i < fieldCount; i++)
                        headers.Add(definition.GetField(i).GetName());
                    sw.WriteLine(string.Join(",", headers.Select(h => EscapeCsvValue(h))));

                    // Data
                    TableData tableData = schedule.GetTableData();
                    TableSectionData bodyData = tableData.GetSectionData(SectionType.Body);
                    if (bodyData != null)
                    {
                        for (int row = 0; row < bodyData.NumberOfRows; row++)
                        {
                            List<string> values = new List<string>();
                            for (int col = 0; col < fieldCount; col++)
                            {
                                string cellText = schedule.GetCellText(SectionType.Body, row, col) ?? "";
                                values.Add(EscapeCsvValue(cellText));
                            }
                            sw.WriteLine(string.Join(",", values));
                        }
                    }
                }
                Debug.WriteLine($"Exported: {schedule.Name} → {filePath}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error exporting {schedule.Name}: {ex.Message}");
                throw;
            }
        }

        private string EscapeCsvValue(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";
            if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
                return $"\"{value.Replace("\"", "\"\"")}\"";
            return value;
        }

        private string SanitizeFileName(string fileName)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                fileName = fileName.Replace(c, '_');
            return fileName;
        }

        // ====================== FORM CHỌN SCHEDULE ======================
        private class ScheduleSelectionForm : System.Windows.Forms.Form
        {
            private readonly List<ScheduleInfo> _allSchedules;
            private CheckedListBox _checkedListBox;

            public ScheduleSelectionForm(List<ScheduleInfo> schedulesInfo)
            {
                _allSchedules = schedulesInfo;
                InitializeComponent();
            }

            private void InitializeComponent()
            {
                this.Text = "Chọn Schedules để Xuất";
                this.Width = 800;
                this.Height = 600;
                this.StartPosition = FormStartPosition.CenterScreen;
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.Font = new Font("Arial", 9);

                var lblPrompt = new Label
                {
                    Text = "Chọn các schedules bạn muốn xuất:",
                    Location = new System.Drawing.Point(10, 10),
                    Size = new Size(760, 20),
                    AutoSize = true
                };

                _checkedListBox = new CheckedListBox
                {
                    Location = new System.Drawing.Point(10, 40),
                    Size = new Size(760, 400),
                    CheckOnClick = true
                };

                foreach (var info in _allSchedules)
                {
                    string text = $"{info.Name} ({info.Columns.Count} cột, {info.RowCount} dòng)";
                    _checkedListBox.Items.Add(text, false);
                }

                var btnSelectAll = new Button { Text = "Select All", Location = new System.Drawing.Point(10, 460), Size = new Size(100, 30) };
                btnSelectAll.Click += (s, e) =>
                {
                    for (int i = 0; i < _checkedListBox.Items.Count; i++) _checkedListBox.SetItemChecked(i, true);
                };

                var btnDeselectAll = new Button { Text = "Deselect All", Location = new System.Drawing.Point(120, 460), Size = new Size(100, 30) };
                btnDeselectAll.Click += (s, e) =>
                {
                    for (int i = 0; i < _checkedListBox.Items.Count; i++) _checkedListBox.SetItemChecked(i, false);
                };

                var btnOK = new Button { Text = "OK", Location = new System.Drawing.Point(630, 460), Size = new Size(80, 30), DialogResult = DialogResult.OK };
                var btnCancel = new Button { Text = "Cancel", Location = new System.Drawing.Point(720, 460), Size = new Size(80, 30), DialogResult = DialogResult.Cancel };

                this.Controls.Add(lblPrompt);
                this.Controls.Add(_checkedListBox);
                this.Controls.Add(btnSelectAll);
                this.Controls.Add(btnDeselectAll);
                this.Controls.Add(btnOK);
                this.Controls.Add(btnCancel);

                this.AcceptButton = btnOK;
                this.CancelButton = btnCancel;
            }

            public List<ScheduleInfo> GetSelectedSchedules()
            {
                var result = new List<ScheduleInfo>();
                foreach (int i in _checkedListBox.CheckedIndices)
                {
                    result.Add(_allSchedules[i]);
                }
                return result;
            }
        }
    }
}