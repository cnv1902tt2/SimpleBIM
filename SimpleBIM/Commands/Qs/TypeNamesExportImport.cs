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
    public class TypeNamesExportImport : IExternalCommand
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
                string[] options = { "Export All Types to CSV", "Import and Rename Types from CSV" };
                string selectedOption = ShowSelectionForm(options, "Type Names Export/Import");

                if (selectedOption == null)
                {
                    return Result.Cancelled;
                }

                if (selectedOption == "Export All Types to CSV")
                {
                    ExportAllTypes(doc);
                }
                else if (selectedOption == "Import and Rename Types from CSV")
                {
                    ImportRenameTypes(doc);
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
                Debug.WriteLine($"Error in TypeNamesExportImport: {ex.Message}\n{ex.StackTrace}");
                return Result.Failed;
            }
        }

        private string ShowSelectionForm(string[] options, string title)
        {
            WinForms.Form selectionForm = new WinForms.Form
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

        private bool IsSystemType(ElementType elementType)
        {
            try
            {
                var propInfo = elementType.GetType().GetProperty("IsSystemType");
                if (propInfo != null)
                {
                    object result = propInfo.GetValue(elementType, null);
                    if (result is bool)
                    {
                        return (bool)result;
                    }
                }
            }
            catch { }

            try
            {
                if (elementType.Name.StartsWith("System"))
                {
                    return true;
                }
            }
            catch { }

            return false;
        }

        private string GetTypeCategory(ElementType elementType)
        {
            try
            {
                if (elementType.Category != null)
                {
                    return elementType.Category.Name;
                }
            }
            catch { }

            return "Unknown";
        }

        private void ExportAllTypes(RevitDocument doc)
        {
            // Lấy tất cả families
            FilteredElementCollector familyCollector = new FilteredElementCollector(doc);
            List<Family> allFamilies = familyCollector.OfClass(typeof(Family))
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
                FileName = "Types_Rename.csv",
                Title = "Export Types to CSV"
            };

            if (saveDialog.ShowDialog() != WinFormsDialogResult.OK)
            {
                return;
            }

            string filePath = saveDialog.FileName;

            // Chuẩn bị dữ liệu
            List<Dictionary<string, object>> typesData = new List<Dictionary<string, object>>();

            // Lặp qua từng family để lấy types (FamilySymbol)
            foreach (Family family in allFamilies)
            {
                try
                {
                    string familyName = family.Name;

                    // Lấy tất cả FamilySymbol IDs từ family
                    ISet<ElementId> symbolIds = family.GetFamilySymbolIds();

                    foreach (ElementId symbolId in symbolIds)
                    {
                        try
                        {
                            RevitElement symbolElement = doc.GetElement(symbolId);
                            if (symbolElement == null || !(symbolElement is FamilySymbol))
                            {
                                continue;
                            }

                            FamilySymbol symbol = (FamilySymbol)symbolElement;
                            int typeId = symbolId.IntegerValue;
                            string typeName = symbol.Name;
                            string category = GetTypeCategory(symbol);

                            Dictionary<string, object> typeData = new Dictionary<string, object>
                            {
                                { "id", typeId },
                                { "family_name", familyName },
                                { "category", category },
                                { "name_old", typeName },
                                { "name_new", typeName }
                            };
                            typesData.Add(typeData);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Error processing symbol {symbolId}: {ex.Message}");
                            continue;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error processing family {family.Id}: {ex.Message}");
                    continue;
                }
            }

            if (typesData.Count == 0)
            {
                TaskDialog.Show("Error", "Không tìm thấy Type nào trong dự án!");
                return;
            }

            // Sắp xếp theo Category → Family → Type
            typesData = typesData
                .OrderBy(x => x["category"].ToString())
                .ThenBy(x => x["family_name"].ToString())
                .ThenBy(x => x["name_old"].ToString())
                .ToList();

            // Ghi CSV
            int rowCount = 0;
            try
            {
                using (StreamWriter sw = new StreamWriter(filePath, false, Encoding.UTF8))
                {
                    sw.WriteLine("Type ID,Family Name,Category,Type Name Old,Type Name New");

                    foreach (Dictionary<string, object> typeData in typesData)
                    {
                        try
                        {
                            int typeId = (int)typeData["id"];
                            string familyName = typeData["family_name"].ToString();
                            string category = typeData["category"].ToString();
                            string typeNameOld = typeData["name_old"].ToString();
                            string typeNameNew = typeData["name_new"].ToString();

                            // Escape dấu phẩy
                            if (familyName.Contains(","))
                                familyName = $"\"{familyName.Replace("\"", "\"\"")}\"";
                            if (category.Contains(","))
                                category = $"\"{category.Replace("\"", "\"\"")}\"";
                            if (typeNameOld.Contains(","))
                                typeNameOld = $"\"{typeNameOld.Replace("\"", "\"\"")}\"";
                            if (typeNameNew.Contains(","))
                                typeNameNew = $"\"{typeNameNew.Replace("\"", "\"\"")}\"";

                            sw.WriteLine($"{typeId},{familyName},{category},{typeNameOld},{typeNameNew}");
                            rowCount++;
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Error writing type row: {ex.Message}");
                            continue;
                        }
                    }
                }

                TaskDialog.Show("Export Complete",
                    $"Số type đã xuất: {rowCount}\n\nFile được lưu tại:\n{filePath}");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Export Error", $"Lỗi khi ghi file CSV:\n{ex.Message}");
            }
        }

        private void ImportRenameTypes(RevitDocument doc)
        {
            // Chọn file CSV
            WinFormsOpenFileDialog openDialog = new WinFormsOpenFileDialog
            {
                Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
                Title = "Import Types from CSV"
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
                        if (row.Length >= 5)
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
            using (RevitTransaction trans = new RevitTransaction(doc, "Rename Types"))
            {
                trans.Start();

                try
                {
                    foreach (string[] row in rowsData)
                    {
                        try
                        {
                            string typeIdStr = row[0].Trim();
                            string familyName = row[1].Trim();
                            string category = row[2].Trim();
                            string typeNameOld = row[3].Trim();
                            string typeNameNew = row[4].Trim();

                            // Bỏ qua nếu tên giống nhau
                            if (typeNameOld == typeNameNew || string.IsNullOrEmpty(typeNameNew))
                            {
                                continue;
                            }

                            // Parse Type ID
                            if (!int.TryParse(typeIdStr, out int typeIdInt))
                            {
                                errorLog.Add($"Type ID {typeIdStr} không hợp lệ");
                                continue;
                            }

                            ElementId typeId = new ElementId(typeIdInt);
                            RevitElement element = doc.GetElement(typeId);

                            if (element == null || !(element is FamilySymbol))
                            {
                                errorLog.Add($"Type ID {typeIdStr} không tồn tại");
                                continue;
                            }

                            FamilySymbol symbol = (FamilySymbol)element;

                            // Kiểm tra System Type
                            if (IsSystemType(symbol))
                            {
                                errorLog.Add($"Type '{typeNameOld}' ({familyName}/{category}): Là System Type, không thể rename");
                                continue;
                            }

                            // Kiểm tra tên trùng trong cùng family
                            Family symbolFamily = symbol.Family;
                            ISet<ElementId> siblingIds = symbolFamily.GetFamilySymbolIds();
                            bool nameExists = false;

                            foreach (ElementId siblingId in siblingIds)
                            {
                                if (siblingId.IntegerValue != typeIdInt)
                                {
                                    RevitElement siblingElem = doc.GetElement(siblingId);
                                    if (siblingElem != null && siblingElem is FamilySymbol)
                                    {
                                        FamilySymbol sibling = (FamilySymbol)siblingElem;
                                        if (sibling.Name == typeNameNew)
                                        {
                                            nameExists = true;
                                            break;
                                        }
                                    }
                                }
                            }

                            if (nameExists)
                            {
                                errorLog.Add($"Type '{typeNameOld}' ({familyName}): Tên '{typeNameNew}' đã tồn tại");
                                continue;
                            }

                            // Rename
                            try
                            {
                                symbol.Name = typeNameNew;
                                successCount++;
                                Debug.WriteLine($"Renamed: '{typeNameOld}' [{familyName}/{category}] → '{typeNameNew}'");
                            }
                            catch (Exception ex)
                            {
                                errorLog.Add($"Lỗi khi rename Type '{typeNameOld}': {ex.Message}");
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
