using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Diagnostics;

// WINDOWS FORMS NAMESPACES
using System.Windows.Forms;
using System.Drawing;
using System.ComponentModel;
using System.Data;

// AUTODESK REVIT API NAMESPACES
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.UI.Events;
using Autodesk.Revit.Exceptions;

// ============================================================================
// ALIASES ĐỂ TRÁNH AMBIGUOUS REFERENCES
// ============================================================================

using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;

using RevitApplication = Autodesk.Revit.ApplicationServices.Application;
using RevitDocument = Autodesk.Revit.DB.Document;
using RevitUIApplication = Autodesk.Revit.UI.UIApplication;
using RevitUIDocument = Autodesk.Revit.UI.UIDocument;
using RevitView = Autodesk.Revit.DB.View;
using RevitElement = Autodesk.Revit.DB.Element;
using RevitLevel = Autodesk.Revit.DB.Level;
using RevitTransaction = Autodesk.Revit.DB.Transaction;
using RevitCategory = Autodesk.Revit.DB.Category;
using RevitParameter = Autodesk.Revit.DB.Parameter;

using WinFormsControl = System.Windows.Forms.Control;
using WinFormsTextBox = System.Windows.Forms.TextBox;
using WinFormsButton = System.Windows.Forms.Button;
using WinFormsLabel = System.Windows.Forms.Label;
using WinFormsOpenFileDialog = System.Windows.Forms.OpenFileDialog;
using WinFormsSaveFileDialog = System.Windows.Forms.SaveFileDialog;
using WinFormsDialogResult = System.Windows.Forms.DialogResult;
using OperationCanceledException = Autodesk.Revit.Exceptions.OperationCanceledException;

namespace SimpleBIM.Commands.Qs
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class FamilyNamesExportImport : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            RevitUIApplication uiapp = commandData.Application;
            RevitUIDocument uidoc = uiapp.ActiveUIDocument;
            RevitDocument doc = uidoc.Document;

            try
            {
                // Hiển thị menu lựa chọn
                string[] options = { "Export All Families to CSV", "Import and Rename Families from CSV" };
                string selectedOption = ShowSelectionForm(options, "Family Names Export/Import");

                if (selectedOption == null)
                {
                    return Result.Cancelled;
                }

                if (selectedOption == "Export All Families to CSV")
                {
                    ExportAllFamilies(doc);
                }
                else if (selectedOption == "Import and Rename Families from CSV")
                {
                    ImportRenameFamilies(doc);
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
                Debug.WriteLine($"Error in FamilyNamesExportImport: {ex.Message}\n{ex.StackTrace}");
                return Result.Failed;
            }
        }

        private string ShowSelectionForm(string[] options, string title)
        {
            // Tạo simple dialog để chọn option
            System.Windows.Forms.Form selectionForm = new System.Windows.Forms.Form
            {
                Text = title,
                Width = 400,
                Height = 200,
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                ControlBox = true
            };

            WinFormsLabel lblPrompt = new WinFormsLabel
            {
                Text = "Chọn chức năng:",
                Location = new Drawing.Point(10, 10),
                Size = new Drawing.Size(360, 20),
                AutoSize = false
            };
            selectionForm.Controls.Add(lblPrompt);

            System.Windows.Forms.ComboBox cmbOptions = new System.Windows.Forms.ComboBox
            {
                Location = new Drawing.Point(10, 40),
                Size = new Drawing.Size(360, 25),
                DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            };
            cmbOptions.Items.AddRange(options);
            cmbOptions.SelectedIndex = 0;
            selectionForm.Controls.Add(cmbOptions);

            WinFormsButton btnOK = new WinFormsButton
            {
                Text = "OK",
                Location = new Drawing.Point(150, 100),
                Size = new Drawing.Size(80, 30),
                DialogResult = WinFormsDialogResult.OK
            };
            selectionForm.Controls.Add(btnOK);

            WinFormsButton btnCancel = new WinFormsButton
            {
                Text = "Cancel",
                Location = new Drawing.Point(240, 100),
                Size = new Drawing.Size(80, 30),
                DialogResult = WinFormsDialogResult.Cancel
            };
            selectionForm.Controls.Add(btnCancel);

            if (selectionForm.ShowDialog() == WinFormsDialogResult.OK)
            {
                return (string)cmbOptions.SelectedItem;
            }

            return null;
        }

        private bool IsSystemFamily(Family family)
        {
            try
            {
                // Kiểm tra property IsSystemFamily nếu có
                var propInfo = family.GetType().GetProperty("IsSystemFamily");
                if (propInfo != null)
                {
                    object result = propInfo.GetValue(family, null);
                    if (result is bool)
                    {
                        return (bool)result;
                    }
                }
            }
            catch { }

            // Kiểm tra tên family
            try
            {
                if (family.Name.StartsWith("System"))
                {
                    return true;
                }
            }
            catch { }

            return false;
        }

        private string GetFamilyCategory(Family family)
        {
            try
            {
                if (family.FamilyCategory != null)
                {
                    return family.FamilyCategory.Name;
                }
            }
            catch { }

            return "Unknown";
        }

        private void ExportAllFamilies(RevitDocument doc)
        {
            // Lấy tất cả families
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            List<Family> allFamilies = collector.OfClass(typeof(Family))
                .Cast<Family>()
                .ToList();

            if (allFamilies.Count == 0)
            {
                TaskDialog.Show("Error", "Không tìm thấy Family nào trong dự án!");
                return;
            }

            // Chọn file lưu
            WinFormsSaveFileDialog saveDialog = new WinFormsSaveFileDialog
            {
                Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
                FileName = "Families_Rename.csv",
                Title = "Export Families to CSV"
            };

            if (saveDialog.ShowDialog() != WinFormsDialogResult.OK)
            {
                return;
            }

            string filePath = saveDialog.FileName;

            // Chuẩn bị dữ liệu
            List<Dictionary<string, object>> familiesData = new List<Dictionary<string, object>>();
            foreach (Family family in allFamilies)
            {
                try
                {
                    Dictionary<string, object> familyData = new Dictionary<string, object>
                    {
                        { "id", family.Id.IntegerValue },
                        { "category", GetFamilyCategory(family) },
                        { "name_old", family.Name },
                        { "name_new", family.Name }
                    };
                    familiesData.Add(familyData);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error processing family {family.Id}: {ex.Message}");
                    continue;
                }
            }

            // Sắp xếp theo category
            familiesData = familiesData
                .OrderBy(x => x["category"].ToString())
                .ThenBy(x => x["name_old"].ToString())
                .ToList();

            // Ghi CSV
            int rowCount = 0;
            try
            {
                using (StreamWriter sw = new StreamWriter(filePath, false, Encoding.UTF8))
                {
                    sw.WriteLine("Family ID,Category,Family Name Old,Family Name New");

                    foreach (Dictionary<string, object> familyData in familiesData)
                    {
                        try
                        {
                            int familyId = (int)familyData["id"];
                            string category = familyData["category"].ToString();
                            string familyNameOld = familyData["name_old"].ToString();
                            string familyNameNew = familyData["name_new"].ToString();

                            // Escape dấu phẩy
                            if (category.Contains(","))
                                category = $"\"{category.Replace("\"", "\"\"")}\"";
                            if (familyNameOld.Contains(","))
                                familyNameOld = $"\"{familyNameOld.Replace("\"", "\"\"")}\"";
                            if (familyNameNew.Contains(","))
                                familyNameNew = $"\"{familyNameNew.Replace("\"", "\"\"")}\"";

                            sw.WriteLine($"{familyId},{category},{familyNameOld},{familyNameNew}");
                            rowCount++;
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Error writing family row: {ex.Message}");
                            continue;
                        }
                    }
                }

                TaskDialog.Show("Export Complete",
                    $"Số family đã xuất: {rowCount}\n\nFile được lưu tại:\n{filePath}");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Export Error", $"Lỗi khi ghi file CSV:\n{ex.Message}");
            }
        }

        private void ImportRenameFamilies(RevitDocument doc)
        {
            // Chọn file CSV
            WinFormsOpenFileDialog openDialog = new WinFormsOpenFileDialog
            {
                Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
                Title = "Import Families from CSV"
            };

            if (openDialog.ShowDialog() != WinFormsDialogResult.OK)
            {
                return;
            }

            string csvFile = openDialog.FileName;

            // Đọc CSV
            List<string[]> rowsData = new List<string[]>();
            try
            {
                using (StreamReader sr = new StreamReader(csvFile, Encoding.UTF8))
                {
                    string headerLine = sr.ReadLine(); // Skip header
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        string[] row = ParseCsvLine(line);
                        if (row.Length >= 4)
                        {
                            rowsData.Add(row);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Lỗi khi đọc file CSV:\n{ex.Message}");
                return;
            }

            if (rowsData.Count == 0)
            {
                TaskDialog.Show("Warning", "File CSV không có dữ liệu!");
                return;
            }

            int successCount = 0;
            List<string> errorLog = new List<string>();

            // Thực hiện rename
            using (RevitTransaction trans = new RevitTransaction(doc, "Rename Families"))
            {
                trans.Start();

                try
                {
                    foreach (string[] row in rowsData)
                    {
                        try
                        {
                            string familyIdStr = row[0].Trim();
                            string category = row[1].Trim();
                            string familyNameOld = row[2].Trim();
                            string familyNameNew = row[3].Trim();

                            // Bỏ qua nếu tên giống nhau
                            if (familyNameOld == familyNameNew || string.IsNullOrEmpty(familyNameNew))
                            {
                                continue;
                            }

                            // Parse Family ID
                            if (!int.TryParse(familyIdStr, out int familyIdInt))
                            {
                                errorLog.Add($"Family ID {familyIdStr} không hợp lệ");
                                continue;
                            }

                            ElementId familyId = new ElementId(familyIdInt);
                            RevitElement element = doc.GetElement(familyId);

                            if (element == null || !(element is Family))
                            {
                                errorLog.Add($"Family ID {familyIdStr} không tồn tại");
                                continue;
                            }

                            Family family = (Family)element;

                            // Kiểm tra System Family
                            if (IsSystemFamily(family))
                            {
                                errorLog.Add($"Family '{familyNameOld}' ({category}): Là System Family, không thể rename");
                                continue;
                            }

                            // Kiểm tra tên trùng
                            FilteredElementCollector familyCollector = new FilteredElementCollector(doc);
                            List<Family> allFamilies = familyCollector.OfClass(typeof(Family))
                                .Cast<Family>()
                                .ToList();

                            bool nameExists = allFamilies.Any(f => f.Id.IntegerValue != familyIdInt && f.Name == familyNameNew);
                            if (nameExists)
                            {
                                errorLog.Add($"Family '{familyNameOld}' ({category}): Tên '{familyNameNew}' đã tồn tại");
                                continue;
                            }

                            // Rename
                            try
                            {
                                family.Name = familyNameNew;
                                successCount++;
                                Debug.WriteLine($"Renamed: '{familyNameOld}' [{category}] → '{familyNameNew}'");
                            }
                            catch (Exception ex)
                            {
                                errorLog.Add($"Lỗi khi rename Family '{familyNameOld}': {ex.Message}");
                                continue;
                            }
                        }
                        catch (Exception ex)
                        {
                            errorLog.Add($"Lỗi khi xử lý row: {ex.Message}");
                            continue;
                        }
                    }

                    trans.Commit();
                }
                catch (Exception ex)
                {
                    trans.RollBack();
                    TaskDialog.Show("Error", $"Lỗi nghiêm trọng:\n{ex.Message}");
                    return;
                }
            }

            // Hiển thị kết quả
            string resultMsg = $"Import hoàn tất!\n\nRename thành công: {successCount}\nLỗi: {errorLog.Count}";
            if (errorLog.Count > 0)
            {
                resultMsg += "\n\nXem Output để biết chi tiết.";
                foreach (string err in errorLog)
                {
                    Debug.WriteLine($"- {err}");
                }
            }

            TaskDialog.Show("Import Complete", resultMsg);
        }

        private string[] ParseCsvLine(string line)
        {
            List<string> values = new List<string>();
            StringBuilder current = new StringBuilder();
            bool inQuotes = false;

            foreach (char c in line)
            {
                if (c == '"')
                {
                    inQuotes = !inQuotes;
                }
                else if (c == ',' && !inQuotes)
                {
                    values.Add(current.ToString());
                    current.Clear();
                }
                else
                {
                    current.Append(c);
                }
            }

            values.Add(current.ToString());
            return values.ToArray();
        }
    }
}
