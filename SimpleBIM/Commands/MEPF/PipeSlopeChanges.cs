using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace SimpleBIM.Commands.MEPF
{
    /// <summary>
    /// Change Slope of Existing Pipe Series
    /// Modify slope of pipes that already have slope
    /// VERSION: 1.0 - Fixed - Converted from Python to C#
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class PipeSlopeChanges : IExternalCommand
    {
        private UIDocument _uidoc;
        private Document _doc;
        private UIApplication _uiapp;

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            _uiapp = commandData.Application;
            _uidoc = _uiapp.ActiveUIDocument;
            _doc = _uidoc.Document;

            try
            {
                return RunMain();
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Error", $"Error executing command: {ex.Message}\n\n{ex.StackTrace}");
                return Result.Failed;
            }
        }

        private Result RunMain()
        {
            // === STEP 1: Get selected elements ===
            ICollection<ElementId> selectedIds;
            try
            {
                selectedIds = _uidoc.Selection.GetElementIds();
            }
            catch
            {
                TaskDialog.Show("pyRevit", "Please select pipes in order.");
                return Result.Cancelled;
            }

            if (selectedIds.Count == 0)
            {
                TaskDialog.Show("pyRevit", "Please select pipes in order.");
                return Result.Cancelled;
            }

            // === STEP 2: Extract Pipes from selection (in order) ===
            List<Element> pipes = new List<Element>();

            foreach (ElementId elId in selectedIds)
            {
                Element element = _doc.GetElement(elId);
                string elementTypeName = element.GetType().Name;
                if (elementTypeName == "Pipe")
                {
                    pipes.Add(element);
                }
            }

            if (pipes.Count < 2)
            {
                TaskDialog.Show("pyRevit", "Please select at least 2 Pipes in series.");
                return Result.Cancelled;
            }

            System.Diagnostics.Debug.WriteLine($"Found {pipes.Count} Pipe(s) in selection order");

            // === STEP 3: Validate pipes (must have slope) ===
            List<(int, Element, double, double)> pipesValid = new List<(int, Element, double, double)>();

            for (int idx = 0; idx < pipes.Count; idx++)
            {
                Element pipe = pipes[idx];
                try
                {
                    LocationCurve location = pipe.Location as LocationCurve;
                    if (location == null)
                        continue;

                    Curve curve = location.Curve;
                    if (curve == null)
                        continue;

                    XYZ startPoint = curve.GetEndPoint(0);
                    XYZ endPoint = curve.GetEndPoint(1);

                    double startElevation = startPoint.Z;
                    double endElevation = endPoint.Z;

                    double elevDiff = Math.Abs(endElevation - startElevation);

                    // Must have slope (elevation difference > 0.01)
                    if (elevDiff > 0.01)
                    {
                        pipesValid.Add((idx, pipe, startElevation, endElevation));
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"Skip: Pipe {pipe.Id.Value} has no slope");
                    }
                }
                catch
                {
                    // Skip pipes with errors
                }
            }

            if (pipesValid.Count < 2)
            {
                TaskDialog.Show("pyRevit", "Not enough Pipes with slope.");
                return Result.Cancelled;
            }

            System.Diagnostics.Debug.WriteLine($"Valid Pipes with slope: {pipesValid.Count}");

            // === STEP 4: Ask for new slope value ===
            string slopeInput = ShowInputDialog("Enter new Slope value (%):", "Pipe Slope Changer", "1.0");

            if (slopeInput == null)
            {
                TaskDialog.Show("pyRevit", "Cancelled.");
                return Result.Cancelled;
            }

            double newSlopePercent;
            try
            {
                newSlopePercent = double.Parse(slopeInput);
                if (newSlopePercent <= 0)
                {
                    TaskDialog.Show("pyRevit", "Slope must be greater than 0.");
                    return Result.Cancelled;
                }
            }
            catch
            {
                TaskDialog.Show("pyRevit", "Invalid slope value.");
                return Result.Cancelled;
            }

            System.Diagnostics.Debug.WriteLine($"New Slope: {newSlopePercent}%");

            // === STEP 5: Ask which end is higher ===
            TaskDialogResult endChoice = TaskDialog.Show(
                "Pipe Slope Changer",
                "Select which end will be HIGHER:\n\n[Yes] Start Point\n[No] End Point",
                TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

            if (endChoice != TaskDialogResult.Yes && endChoice != TaskDialogResult.No)
            {
                TaskDialog.Show("pyRevit", "Cancelled.");
                return Result.Cancelled;
            }

            string highEnd = endChoice == TaskDialogResult.Yes ? "Start Point" : "End Point";
            System.Diagnostics.Debug.WriteLine($"High end: {highEnd}");

            // === STEP 6: Change slope of pipes ===
            using (Transaction t = new Transaction(_doc, "Change Pipe Slope"))
            {
                t.Start();

                try
                {
                    double currentEndElevation = 0;

                    // ===== CASE 1: [Yes] Start Point is HIGHER (Downward slope) =====
                    if (endChoice == TaskDialogResult.Yes)
                    {
                        System.Diagnostics.Debug.WriteLine("Changing slope with Start Point HIGHER (Downward)...");

                        for (int idxInValid = 0; idxInValid < pipesValid.Count; idxInValid++)
                        {
                            var (originalIdx, pipe, oldStartElev, oldEndElev) = pipesValid[idxInValid];
                            try
                            {
                                LocationCurve location = pipe.Location as LocationCurve;
                                Curve curve = location.Curve;

                                XYZ startPoint = curve.GetEndPoint(0);
                                XYZ endPoint = curve.GetEndPoint(1);

                                // Calculate pipe length (horizontal distance on plan)
                                double horizontalLength = Math.Sqrt(
                                    Math.Pow(endPoint.X - startPoint.X, 2) +
                                    Math.Pow(endPoint.Y - startPoint.Y, 2));

                                double slopeDecimal = newSlopePercent / 100.0;
                                double elevationDifference = slopeDecimal * horizontalLength;

                                double newStartElev, newEndElev;

                                // First pipe
                                if (idxInValid == 0)
                                {
                                    newStartElev = oldStartElev;
                                    newEndElev = newStartElev - elevationDifference;
                                    currentEndElevation = newEndElev;
                                }
                                // Subsequent pipes: maintain same slope direction (downward)
                                else
                                {
                                    newStartElev = currentEndElevation;
                                    newEndElev = currentEndElevation - elevationDifference;
                                    currentEndElevation = newEndElev;
                                }

                                // Keep original X, Y coordinates, only change Z
                                XYZ newStartXyz = new XYZ(startPoint.X, startPoint.Y, newStartElev);
                                XYZ newEndXyz = new XYZ(endPoint.X, endPoint.Y, newEndElev);

                                // Update pipe curve
                                location.Curve = Line.CreateBound(newStartXyz, newEndXyz);

                                System.Diagnostics.Debug.WriteLine($"Changed Pipe {pipe.Id.Value} slope (sequence {idxInValid + 1})");
                            }
                            catch (Exception e)
                            {
                                System.Diagnostics.Debug.WriteLine($"Failed Pipe {pipe.Id.Value}: {e.Message}");
                            }
                        }
                    }
                    // ===== CASE 2: [No] End Point is HIGHER (Upward slope) =====
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("Changing slope with End Point HIGHER (Upward)...");

                        for (int idxInValid = 0; idxInValid < pipesValid.Count; idxInValid++)
                        {
                            var (originalIdx, pipe, oldStartElev, oldEndElev) = pipesValid[idxInValid];
                            try
                            {
                                LocationCurve location = pipe.Location as LocationCurve;
                                Curve curve = location.Curve;

                                XYZ startPoint = curve.GetEndPoint(0);
                                XYZ endPoint = curve.GetEndPoint(1);

                                // Calculate pipe length (horizontal distance on plan)
                                double horizontalLength = Math.Sqrt(
                                    Math.Pow(endPoint.X - startPoint.X, 2) +
                                    Math.Pow(endPoint.Y - startPoint.Y, 2));

                                double slopeDecimal = newSlopePercent / 100.0;
                                double elevationDifference = slopeDecimal * horizontalLength;

                                double newStartElev, newEndElev;

                                // First pipe
                                if (idxInValid == 0)
                                {
                                    newStartElev = oldStartElev;
                                    newEndElev = newStartElev + elevationDifference;
                                    currentEndElevation = newEndElev;
                                }
                                // Subsequent pipes: maintain same slope direction (upward)
                                else
                                {
                                    newStartElev = currentEndElevation;
                                    newEndElev = currentEndElevation + elevationDifference;
                                    currentEndElevation = newEndElev;
                                }

                                // Keep original X, Y coordinates, only change Z
                                XYZ newStartXyz = new XYZ(startPoint.X, startPoint.Y, newStartElev);
                                XYZ newEndXyz = new XYZ(endPoint.X, endPoint.Y, newEndElev);

                                // Update pipe curve
                                location.Curve = Line.CreateBound(newStartXyz, newEndXyz);

                                System.Diagnostics.Debug.WriteLine($"Changed Pipe {pipe.Id.Value} slope (sequence {idxInValid + 1})");
                            }
                            catch (Exception e)
                            {
                                System.Diagnostics.Debug.WriteLine($"Failed Pipe {pipe.Id.Value}: {e.Message}");
                            }
                        }
                    }

                    t.Commit();
                    System.Diagnostics.Debug.WriteLine("Pipe Slope Change Complete!");
                    TaskDialog.Show("pyRevit", "Pipe Slope Change Complete!");
                    return Result.Succeeded;
                }
                catch (Exception e)
                {
                    t.RollBack();
                    TaskDialog.Show("pyRevit", $"Error: {e.Message}");
                    return Result.Failed;
                }
            }
        }

        /// <summary>
        /// Show simple input dialog (using TaskDialog as fallback)
        /// </summary>
        private string ShowInputDialog(string prompt, string title, string defaultValue)
        {
            // Simple implementation using Windows Forms TextBox
            using (var form = new System.Windows.Forms.Form())
            {
                var label = new System.Windows.Forms.Label();
                var textBox = new System.Windows.Forms.TextBox();
                var buttonOk = new System.Windows.Forms.Button();
                var buttonCancel = new System.Windows.Forms.Button();

                form.Text = title;
                label.Text = prompt;
                textBox.Text = defaultValue;

                buttonOk.Text = "OK";
                buttonCancel.Text = "Cancel";
                buttonOk.DialogResult = System.Windows.Forms.DialogResult.OK;
                buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;

                label.SetBounds(9, 20, 372, 13);
                textBox.SetBounds(12, 36, 372, 20);
                buttonOk.SetBounds(228, 72, 75, 23);
                buttonCancel.SetBounds(309, 72, 75, 23);

                label.AutoSize = true;
                textBox.Anchor = textBox.Anchor | System.Windows.Forms.AnchorStyles.Right;
                buttonOk.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
                buttonCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;

                form.ClientSize = new System.Drawing.Size(396, 107);
                form.Controls.AddRange(new System.Windows.Forms.Control[] { label, textBox, buttonOk, buttonCancel });
                form.ClientSize = new System.Drawing.Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
                form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                form.MinimizeBox = false;
                form.MaximizeBox = false;
                form.AcceptButton = buttonOk;
                form.CancelButton = buttonCancel;

                System.Windows.Forms.DialogResult dialogResult = form.ShowDialog();
                return dialogResult == System.Windows.Forms.DialogResult.OK ? textBox.Text : null;
            }
        }
    }
}

/* 
**PYREVIT → C# CONVERSIONS APPLIED:**
1. `doc` / `uidoc` → Retrieved from `ExternalCommandData`
2. `forms.alert()` → `TaskDialog.Show()`
3. `forms.ask_for_string()` → Custom `ShowInputDialog()` using WinForms
4. `print()` → `System.Diagnostics.Debug.WriteLine()`
5. `element.Id.IntegerValue` → `element.Id.Value` (Revit 2022+ API)
6. Python tuple `(idx, pipe, old_start_elev, old_end_elev)` → C# ValueTuple `(int, Element, double, double)`
7. Python list comprehension → C# List<T> with for loops
8. `forms.alert(..., yes=True, no=True)` → `TaskDialog.Show()` with `TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No`
9. Transaction pattern: `with Transaction()` → `using (Transaction t = new Transaction())`

**THAM KHẢO TỪ Commands/As/:**
- IExternalCommand structure
- Transaction handling with using statement
- Selection validation patterns
- Error handling with try-catch

**IMPORTANT NOTES:**
- Giữ nguyên logic tính toán slope từ Python
- Pipe selection từ pre-selected elements (không pick one by one)
- Xử lý 2 trường hợp: Start Point higher (downward) và End Point higher (upward)
- Sequential slope modification: first pipe giữ nguyên start elevation, các pipe sau nối tiếp
- Input dialog sử dụng Windows Forms để nhập slope value
*/
