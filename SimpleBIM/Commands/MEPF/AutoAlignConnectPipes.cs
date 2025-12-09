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
    /// Auto Align & Auto Connect Pipes System
    /// Convert slope, align pipes (X,Y,Z), auto-create connections
    /// VERSION: 6.1 Fixed Align Logic - Converted from Python to C#
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class AutoAlignConnectPipes : IExternalCommand
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
                TaskDialog.Show("pyRevit", "Please select pipes.");
                return Result.Cancelled;
            }

            if (selectedIds.Count == 0)
            {
                TaskDialog.Show("pyRevit", "Please select pipes.");
                return Result.Cancelled;
            }

            // === STEP 2: Extract Pipes in order ===
            List<Element> pipes = new List<Element>();
            Dictionary<long, Element> pipeMap = new Dictionary<long, Element>();

            foreach (ElementId elId in selectedIds)
            {
                Element element = _doc.GetElement(elId);
                if (element.GetType().Name == "Pipe")
                {
                    pipes.Add(element);
                    pipeMap[elId.Value] = element;
                }
            }

            if (pipes.Count < 2)
            {
                TaskDialog.Show("pyRevit", "Please select at least 2 Pipes.");
                return Result.Cancelled;
            }

            System.Diagnostics.Debug.WriteLine($"Found {pipes.Count} Pipe(s)");

            // === STEP 3: Ask for slope ===
            string slopeInput = ShowInputDialog("Enter Slope value (%):", "Auto Align Connect V6.1", "1.0");

            if (slopeInput == null)
            {
                TaskDialog.Show("pyRevit", "Cancelled.");
                return Result.Cancelled;
            }

            double slopePercent;
            try
            {
                slopePercent = double.Parse(slopeInput);
                if (slopePercent <= 0)
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

            System.Diagnostics.Debug.WriteLine($"Slope: {slopePercent}%");

            // === STEP 4: Ask direction ===
            TaskDialogResult endChoice = TaskDialog.Show(
                "Auto Align Connect",
                "Select which end will be HIGHER:\n\n[Yes] Start Point\n[No] End Point",
                TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

            if (endChoice != TaskDialogResult.Yes && endChoice != TaskDialogResult.No)
            {
                TaskDialog.Show("pyRevit", "Cancelled.");
                return Result.Cancelled;
            }

            // === STEP 5: Main Transaction ===
            using (Transaction t = new Transaction(_doc, "Auto Align Connect Pipes"))
            {
                t.Start();

                try
                {
                    // ===== PHASE 1: Convert Slope =====
                    System.Diagnostics.Debug.WriteLine("\n=== PHASE 1: Convert Slope ===");
                    double slopeDecimal = slopePercent / 100.0;

                    for (int idx = 0; idx < pipes.Count; idx++)
                    {
                        Element pipe = pipes[idx];
                        try
                        {
                            LocationCurve location = pipe.Location as LocationCurve;
                            Curve curve = location.Curve;

                            XYZ startPoint = curve.GetEndPoint(0);
                            XYZ endPoint = curve.GetEndPoint(1);

                            double horizontalLength = Math.Sqrt(
                                Math.Pow(endPoint.X - startPoint.X, 2) +
                                Math.Pow(endPoint.Y - startPoint.Y, 2));

                            double elevationDifference = slopeDecimal * horizontalLength;

                            double newStartElev, newEndElev;

                            if (endChoice == TaskDialogResult.Yes)
                            {
                                newStartElev = startPoint.Z;
                                newEndElev = startPoint.Z - elevationDifference;
                            }
                            else
                            {
                                newEndElev = startPoint.Z;
                                newStartElev = startPoint.Z - elevationDifference;
                            }

                            XYZ newStartXyz = new XYZ(startPoint.X, startPoint.Y, newStartElev);
                            XYZ newEndXyz = new XYZ(endPoint.X, endPoint.Y, newEndElev);

                            location.Curve = Line.CreateBound(newStartXyz, newEndXyz);
                            System.Diagnostics.Debug.WriteLine($"Converted Pipe {pipe.Id.Value} (idx {idx})");
                        }
                        catch (Exception e)
                        {
                            System.Diagnostics.Debug.WriteLine($"ERROR converting Pipe {pipe.Id.Value}: {e.Message}");
                        }
                    }

                    // ===== PHASE 2: Delete Fittings =====
                    System.Diagnostics.Debug.WriteLine("\n=== PHASE 2: Delete Fittings ===");

                    int fittingsDeleted = 0;
                    List<ElementId> idsToDelete = new List<ElementId>();

                    foreach (ElementId elId in selectedIds)
                    {
                        Element element = _doc.GetElement(elId);
                        if (element.GetType().Name == "FamilyInstance")
                        {
                            idsToDelete.Add(elId);
                        }
                    }

                    foreach (ElementId elId in idsToDelete)
                    {
                        try
                        {
                            _doc.Delete(elId);
                            fittingsDeleted++;
                            System.Diagnostics.Debug.WriteLine($"Deleted Fitting {elId.Value}");
                        }
                        catch (Exception e)
                        {
                            System.Diagnostics.Debug.WriteLine($"Failed to delete fitting: {e.Message}");
                        }
                    }

                    System.Diagnostics.Debug.WriteLine($"Total fittings deleted: {fittingsDeleted}");

                    // ===== PHASE 3: Collect current pipe data =====
                    System.Diagnostics.Debug.WriteLine("\n=== PHASE 3: Collect Pipe Data ===");

                    List<PipeData> pipesData = new List<PipeData>();

                    for (int idx = 0; idx < pipes.Count; idx++)
                    {
                        Element pipe = pipes[idx];
                        try
                        {
                            LocationCurve location = pipe.Location as LocationCurve;
                            Curve curve = location.Curve;

                            XYZ startPoint = curve.GetEndPoint(0);
                            XYZ endPoint = curve.GetEndPoint(1);

                            pipesData.Add(new PipeData
                            {
                                Index = idx,
                                Pipe = pipe,
                                PipeId = pipe.Id.Value,
                                StartPoint = startPoint,
                                EndPoint = endPoint,
                                Location = location
                            });

                            System.Diagnostics.Debug.WriteLine($"Pipe[{idx}] ID={pipe.Id.Value}: Start({startPoint.X:F2},{startPoint.Y:F2},{startPoint.Z:F2}) End({endPoint.X:F2},{endPoint.Y:F2},{endPoint.Z:F2})");
                        }
                        catch (Exception e)
                        {
                            System.Diagnostics.Debug.WriteLine($"ERROR collecting data for Pipe {pipe.Id.Value}: {e.Message}");
                        }
                    }

                    System.Diagnostics.Debug.WriteLine($"Total pipes collected: {pipesData.Count}");

                    // ===== PHASE 4: Align X, Y Sequential =====
                    System.Diagnostics.Debug.WriteLine("\n=== PHASE 4: Align X, Y (Sequential) ===");

                    for (int i = 0; i < pipesData.Count - 1; i++)
                    {
                        try
                        {
                            PipeData data1 = pipesData[i];
                            PipeData data2 = pipesData[i + 1];

                            Element pipe1 = data1.Pipe;
                            Element pipe2 = data2.Pipe;

                            XYZ p1End = data1.EndPoint;
                            XYZ p2Start = data2.StartPoint;
                            XYZ p2End = data2.EndPoint;

                            // Align: pipe2 start X,Y moves to pipe1 end X,Y
                            XYZ p2StartAlignedXY = new XYZ(p1End.X, p1End.Y, p2Start.Z);

                            // Update pipe2
                            (pipe2.Location as LocationCurve).Curve = Line.CreateBound(p2StartAlignedXY, p2End);

                            // Update data
                            pipesData[i + 1].StartPoint = p2StartAlignedXY;

                            System.Diagnostics.Debug.WriteLine($"Aligned X,Y: Pipe[{i}] -> Pipe[{i + 1}]");
                            System.Diagnostics.Debug.WriteLine($"  Pipe[{i}] end: ({p1End.X:F2}, {p1End.Y:F2})");
                            System.Diagnostics.Debug.WriteLine($"  Pipe[{i + 1}] new start: ({p2StartAlignedXY.X:F2}, {p2StartAlignedXY.Y:F2})");
                        }
                        catch (Exception e)
                        {
                            System.Diagnostics.Debug.WriteLine($"ERROR aligning X,Y between Pipe[{i}] and Pipe[{i + 1}]: {e.Message}");
                        }
                    }

                    // ===== PHASE 5: Align Z Sequential =====
                    System.Diagnostics.Debug.WriteLine("\n=== PHASE 5: Align Z (Sequential) ===");

                    for (int i = 0; i < pipesData.Count - 1; i++)
                    {
                        try
                        {
                            PipeData data1 = pipesData[i];
                            PipeData data2 = pipesData[i + 1];

                            Element pipe1 = data1.Pipe;
                            Element pipe2 = data2.Pipe;

                            XYZ p1End = data1.EndPoint;
                            XYZ p2Start = data2.StartPoint;
                            XYZ p2End = data2.EndPoint;

                            double p1EndZ = p1End.Z;

                            // Recalculate slope for pipe2 based on new horizontal length
                            double horizontalLengthP2 = Math.Sqrt(
                                Math.Pow(p2End.X - p2Start.X, 2) +
                                Math.Pow(p2End.Y - p2Start.Y, 2));

                            double elevationDiffP2 = slopeDecimal * horizontalLengthP2;
                            double p2EndZNew = p1EndZ - elevationDiffP2;

                            // Update pipe2 with aligned Z
                            XYZ p2StartAlignedZ = new XYZ(p2Start.X, p2Start.Y, p1EndZ);
                            XYZ p2EndAlignedZ = new XYZ(p2End.X, p2End.Y, p2EndZNew);

                            (pipe2.Location as LocationCurve).Curve = Line.CreateBound(p2StartAlignedZ, p2EndAlignedZ);

                            // Update data
                            pipesData[i + 1].StartPoint = p2StartAlignedZ;
                            pipesData[i + 1].EndPoint = p2EndAlignedZ;

                            System.Diagnostics.Debug.WriteLine($"Aligned Z: Pipe[{i}] -> Pipe[{i + 1}]");
                            System.Diagnostics.Debug.WriteLine($"  Pipe[{i}] end Z: {p1EndZ:F2}");
                            System.Diagnostics.Debug.WriteLine($"  Pipe[{i + 1}] new start Z: {p1EndZ:F2} (aligned)");
                            System.Diagnostics.Debug.WriteLine($"  Pipe[{i + 1}] new end Z: {p2EndZNew:F2} (slope {slopePercent:F1}%)");
                        }
                        catch (Exception e)
                        {
                            System.Diagnostics.Debug.WriteLine($"ERROR aligning Z between Pipe[{i}] and Pipe[{i + 1}]: {e.Message}");
                        }
                    }

                    // ===== PHASE 6: Verify Results =====
                    System.Diagnostics.Debug.WriteLine("\n=== PHASE 6: Verify Results ===");

                    foreach (PipeData data in pipesData)
                    {
                        Element pipe = data.Pipe;
                        LocationCurve location = pipe.Location as LocationCurve;
                        Curve curve = location.Curve;

                        XYZ startPoint = curve.GetEndPoint(0);
                        XYZ endPoint = curve.GetEndPoint(1);

                        System.Diagnostics.Debug.WriteLine($"Pipe[{data.Index}] Final: Start({startPoint.X:F2},{startPoint.Y:F2},{startPoint.Z:F2}) End({endPoint.X:F2},{endPoint.Y:F2},{endPoint.Z:F2})");
                    }

                    t.Commit();
                    System.Diagnostics.Debug.WriteLine("\n=== ALL PHASES COMPLETE ===");
                    System.Diagnostics.Debug.WriteLine("Pipes aligned successfully!");

                    TaskDialog.Show("Success", "Pipes aligned successfully!");
                    return Result.Succeeded;
                }
                catch (Exception e)
                {
                    t.RollBack();
                    System.Diagnostics.Debug.WriteLine($"FATAL ERROR: {e.Message}");
                    TaskDialog.Show("pyRevit", $"Error: {e.Message}");
                    return Result.Failed;
                }
            }
        }

        /// <summary>
        /// Show simple input dialog
        /// </summary>
        private string ShowInputDialog(string prompt, string title, string defaultValue)
        {
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

        /// <summary>
        /// Data structure for pipe information
        /// </summary>
        private class PipeData
        {
            public int Index { get; set; }
            public Element Pipe { get; set; }
            public long PipeId { get; set; }
            public XYZ StartPoint { get; set; }
            public XYZ EndPoint { get; set; }
            public LocationCurve Location { get; set; }
        }
    }
}

/* 
**PYREVIT → C# CONVERSIONS APPLIED:**
1. `forms.alert()` → `TaskDialog.Show()`
2. `forms.ask_for_string()` → Custom `ShowInputDialog()`
3. `print()` → `System.Diagnostics.Debug.WriteLine()`
4. `element.Id.IntegerValue` → `element.Id.Value`
5. Python dict `{}` → C# class `PipeData`
6. Python dict access `data['key']` → C# property access `data.Key`
7. Transaction: `with Transaction()` → `using (Transaction t = new Transaction())`
8. Collection iteration during delete → Collect IDs first, then delete

**THAM KHẢO TỪ Commands/As/:**
- IExternalCommand structure
- Transaction management
- Selection handling
- Data structure patterns

**IMPORTANT NOTES:**
- 6 PHASES đầy đủ: Convert Slope, Delete Fittings, Collect Data, Align XY, Align Z, Verify
- Sequential alignment: mỗi pipe nối tiếp với pipe trước đó
- Delete fittings: collect IDs trước, delete sau (tránh modify collection during iteration)
- PipeData class: thay thế Python dict để type-safe
- Full debugging output giữ nguyên như Python gốc
*/
