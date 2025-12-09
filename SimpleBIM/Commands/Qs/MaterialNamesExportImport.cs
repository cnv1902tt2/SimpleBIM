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
    public class MaterialNamesExportImport : IExternalCommand
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
                string[] options = { "Export All Materials to CSV", "Import and Rename Materials from CSV" };
                string selectedOption = ShowSelectionForm(options, "Material Names Export/Import");

                if (selectedOption == null)
                {
                    return Result.Cancelled;
                }

                if (selectedOption == "Export All Materials to CSV")
                {
                    ExportAllMaterials(doc);
                }
                else if (selectedOption == "Import and Rename Materials from CSV")
                {
                    ImportRenameMaterials(doc);
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
                Debug.WriteLine($"Error in MaterialNamesExportImport: {ex.Message}\n{ex.StackTrace}");
                return Result.Failed;
            }
        }

        private string ShowSelectionForm(string[] options, string title)
        {
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

        private bool IsSystemMaterial(Material material)
        {
            try
            {
                var propInfo = material.GetType().GetProperty("IsSystemMaterial");
                if (propInfo != null)
                {
                    object result = propInfo.GetValue(material, null);
                    if (result is bool)
                    {
                        return (bool)result;
                    }
                }
            }
            catch { }

            try
            {
                if (material.Name.StartsWith("System"))
                {
                    return true;
                }
            }
            catch { }

            return false;
        }

        private string GetMaterialClass(Material material)
        {
            try
            {
                var propInfo = material.GetType().GetProperty("MaterialClass");
                if (propInfo != null)
                {
                    object result = propInfo.GetValue(material, null);
                    if (result != null)
                    {
                        return result.ToString();
                    }
                }
            }
            catch { }

            return "Unknown";
        }

        private void ExportAllMaterials(RevitDocument doc)
        {
            // Lấy tất cả materials
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            List<Material> allMaterials = collector.OfClass(typeof(Material))
                .Cast<Material>()
                .ToList();

            if (allMaterials.Count == 0)
            {
                TaskDialog.Show("Error", "Không tìm thấy Material nào trong dự án!");
                return;
            }

            // Chọn file lưu
            WinFormsSaveFileDialog saveDialog = new WinFormsSaveFileDialog
            {
                Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
                FileName = "Materials_Rename.csv",
                Title = "Export Materials to CSV"
            };

            if (saveDialog.ShowDialog() != WinFormsDialogResult.OK)
            {
                return;
            }

            string filePath = saveDialog.FileName;

            // Chuẩn bị dữ liệu
            List<Dictionary<string, object>> materialsData = new List<Dictionary<string, object>>();
            foreach (Material material in allMaterials)
            {
                try
                {
                    Dictionary<string, object> matData = new Dictionary<string, object>
                    {
                        { "id", material.Id.IntegerValue },
                        { "mat_class", GetMaterialClass(material) },
                        { "name_old", material.Name },
                        { "name_new", material.Name }
                    };
                    materialsData.Add(matData);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error processing material {material.Id}: {ex.Message}");
                    continue;
                }
            }

            // Sắp xếp theo Material Class
            materialsData = materialsData
                .OrderBy(x => x["mat_class"].ToString().ToLower())
                .ThenBy(x => x["name_old"].ToString().ToLower())
                .ToList();

            // Ghi CSV
            int rowCount = 0;
            try
            {
                using (StreamWriter sw = new StreamWriter(filePath, false, Encoding.UTF8))
                {
                    sw.WriteLine("Material ID,Material Class,Material Name Old,Material Name New");

                    foreach (Dictionary<string, object> matData in materialsData)
                    {
                        try
                        {
                            int matId = (int)matData["id"];
                            string matClass = matData["mat_class"].ToString();
                            string matNameOld = matData["name_old"].ToString();
                            string matNameNew = matData["name_new"].ToString();

                            // Escape dấu phẩy
                            if (matClass.Contains(","))
                                matClass = $"\"{matClass.Replace("\"", "\"\"")}\"";
                            if (matNameOld.Contains(","))
                                matNameOld = $"\"{matNameOld.Replace("\"", "\"\"")}\"";
                            if (matNameNew.Contains(","))
                                matNameNew = $"\"{matNameNew.Replace("\"", "\"\"")}\"";

                            sw.WriteLine($"{matId},{matClass},{matNameOld},{matNameNew}");
                            rowCount++;
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Error writing material row: {ex.Message}");
                            continue;
                        }
                    }
                }

                TaskDialog.Show("Export Complete",
                    $"Số material đã xuất: {rowCount}\n\nFile được lưu tại:\n{filePath}");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Export Error", $"Lỗi khi ghi file CSV:\n{ex.Message}");
            }
        }

        private void ImportRenameMaterials(RevitDocument doc)
        {
            // Chọn file CSV
            WinFormsOpenFileDialog openDialog = new WinFormsOpenFileDialog
            {
                Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
                Title = "Import Materials from CSV"
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
            using (RevitTransaction trans = new RevitTransaction(doc, "Rename Materials"))
            {
                trans.Start();

                try
                {
                    foreach (string[] row in rowsData)
                    {
                        try
                        {
                            string matIdStr = row[0].Trim();
                            string matClass = row[1].Trim();
                            string matNameOld = row[2].Trim();
                            string matNameNew = row[3].Trim();

                            // Bỏ qua nếu tên giống nhau
                            if (matNameOld == matNameNew || string.IsNullOrEmpty(matNameNew))
                            {
                                continue;
                            }

                            // Parse Material ID
                            if (!int.TryParse(matIdStr, out int matIdInt))
                            {
                                errorLog.Add($"Material ID {matIdStr} không hợp lệ");
                                continue;
                            }

                            ElementId matId = new ElementId(matIdInt);
                            RevitElement element = doc.GetElement(matId);

                            if (element == null || !(element is Material))
                            {
                                errorLog.Add($"Material ID {matIdStr} không tồn tại");
                                continue;
                            }

                            Material material = (Material)element;

                            // Kiểm tra System Material
                            if (IsSystemMaterial(material))
                            {
                                errorLog.Add($"Material '{matNameOld}' ({matClass}): Là System Material, không thể rename");
                                continue;
                            }

                            // Kiểm tra tên trùng
                            FilteredElementCollector matCollector = new FilteredElementCollector(doc);
                            List<Material> allMaterials = matCollector.OfClass(typeof(Material))
                                .Cast<Material>()
                                .ToList();

                            bool nameExists = allMaterials.Any(m => m.Id.IntegerValue != matIdInt && m.Name == matNameNew);
                            if (nameExists)
                            {
                                errorLog.Add($"Material '{matNameOld}' ({matClass}): Tên '{matNameNew}' đã tồn tại");
                                continue;
                            }

                            // Rename
                            try
                            {
                                material.Name = matNameNew;
                                successCount++;
                                Debug.WriteLine($"Renamed: '{matNameOld}' [{matClass}] → '{matNameNew}'");
                            }
                            catch (Exception ex)
                            {
                                errorLog.Add($"Lỗi khi rename Material '{matNameOld}': {ex.Message}");
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
