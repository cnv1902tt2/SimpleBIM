using System;
using System.Collections.Generic;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace SimpleBIM.Commands.MEPF
{
    /// <summary>
    /// Align MEP 3D - Align series of pipes to intersect in 3D space
    /// User selects pipes one by one in order, then align sequentially
    /// Keep first pipe fixed, move subsequent pipes to align at exact contact point
    /// VERSION: 2.3 - Exact Point Alignment (Fixed) - Converted from Python to C#
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class AlignMEP3D : IExternalCommand
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
            // === STEP 1: Select pipes one by one (press ESC/Cancel to finish) ===
            List<Element> pipes = new List<Element>();
            int pipeCount = 1;

            TaskDialog.Show("pyRevit", "Select pipes one by one. Press ESC or Cancel to finish selection.");

            while (true)
            {
                try
                {
                    // Ask user to select pipe
                    Reference elementRef = _uidoc.Selection.PickObject(
                        ObjectType.Element, 
                        $"Select Pipe {pipeCount}");
                    
                    Element element = _doc.GetElement(elementRef.ElementId);

                    // Check if it's a Pipe
                    if (element.GetType().Name == "Pipe")
                    {
                        pipes.Add(element);
                        System.Diagnostics.Debug.WriteLine($"Selected Pipe {pipeCount}: ID {element.Id.Value}");
                        pipeCount++;
                    }
                    else
                    {
                        TaskDialog.Show("pyRevit", $"Please select a Pipe, not {element.GetType().Name}");
                    }
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    // User pressed ESC / canceled the pick - finish selection
                    System.Diagnostics.Debug.WriteLine($"Selection finished by user. Total selected: {pipes.Count}");
                    break;
                }
                catch (Exception pickError)
                {
                    // Some other error occurred while picking
                    System.Diagnostics.Debug.WriteLine($"Selection error: {pickError.Message}");
                    break;
                }
            }

            // === STEP 2: Validate selection ===
            if (pipes.Count < 2)
            {
                TaskDialog.Show("pyRevit", "Please select at least 2 pipes.");
                return Result.Cancelled;
            }

            System.Diagnostics.Debug.WriteLine($"\nFound {pipes.Count} Pipe(s) in selection order");
            for (int idx = 0; idx < pipes.Count; idx++)
            {
                System.Diagnostics.Debug.WriteLine($"Pipe {idx + 1}: ID {pipes[idx].Id.Value}");
            }

            // === STEP 5: Align pipes sequentially ===
            using (Transaction tx = new Transaction(_doc, "Align MEP 3D Series"))
            {
                tx.Start();

                try
                {
                    // Process each pair of consecutive pipes
                    for (int pairIdx = 0; pairIdx < pipes.Count - 1; pairIdx++)
                    {
                        Element pipe1 = pipes[pairIdx];
                        Element pipe2 = pipes[pairIdx + 1];

                        System.Diagnostics.Debug.WriteLine($"\n=== Aligning Pair {pairIdx + 1}: Pipe {pairIdx + 1} (fixed) + Pipe {pairIdx + 2} (move) ===");

                        // Get CURRENT geometry
                        var (p1Start, p1End) = GetPipeGeometry(pipe1);
                        var (p2Start, p2End) = GetPipeGeometry(pipe2);

                        if (p1Start == null || p2Start == null)
                        {
                            System.Diagnostics.Debug.WriteLine($"Failed to get geometry for pair {pairIdx + 1}");
                            continue;
                        }

                        System.Diagnostics.Debug.WriteLine($"Pipe {pairIdx + 1} (fixed): Start({p1Start.X:F2},{p1Start.Y:F2},{p1Start.Z:F2}) End({p1End.X:F2},{p1End.Y:F2},{p1End.Z:F2})");
                        System.Diagnostics.Debug.WriteLine($"Pipe {pairIdx + 2} (move): Start({p2Start.X:F2},{p2Start.Y:F2},{p2Start.Z:F2}) End({p2End.X:F2},{p2End.Y:F2},{p2End.Z:F2})");

                        // Calculate closest points between centerlines
                        var (p1OnLine, p2OnLine, distance, s, t) = LineIntersection3D(
                            p1Start, p1End, p2Start, p2End);

                        // Check if lines are parallel (p1OnLine is null)
                        if (p1OnLine == null)
                        {
                            System.Diagnostics.Debug.WriteLine("Lines are parallel. Finding closest points to align centerlines...");

                            // For parallel lines, find closest points between them
                            XYZ d1 = new XYZ(p1End.X - p1Start.X, p1End.Y - p1Start.Y, p1End.Z - p1Start.Z);
                            XYZ w = new XYZ(p2Start.X - p1Start.X, p2Start.Y - p1Start.Y, p2Start.Z - p1Start.Z);

                            double d1DotD1 = d1.DotProduct(d1);
                            XYZ closestOnP1;
                            if (d1DotD1 > 1e-9)
                            {
                                double sProj = d1.DotProduct(w) / d1DotD1;
                                // Clamp s to segment [0, 1]
                                sProj = Math.Max(0, Math.Min(1, sProj));
                                closestOnP1 = new XYZ(p1Start.X + sProj * d1.X,
                                                      p1Start.Y + sProj * d1.Y,
                                                      p1Start.Z + sProj * d1.Z);
                            }
                            else
                            {
                                closestOnP1 = p1Start;
                            }

                            // Use Pipe2 midpoint as reference point to move
                            XYZ midP2 = new XYZ((p2Start.X + p2End.X) / 2.0,
                                               (p2Start.Y + p2End.Y) / 2.0,
                                               (p2Start.Z + p2End.Z) / 2.0);

                            // Calculate translation vector to align midpoint of Pipe2 with closest point on Pipe1
                            XYZ moveVecP2 = new XYZ(closestOnP1.X - midP2.X,
                                                   closestOnP1.Y - midP2.Y,
                                                   closestOnP1.Z - midP2.Z);

                            System.Diagnostics.Debug.WriteLine($"Closest point on Pipe {pairIdx + 1}: ({closestOnP1.X:F2},{closestOnP1.Y:F2},{closestOnP1.Z:F2})");
                            System.Diagnostics.Debug.WriteLine($"Pipe {pairIdx + 2} midpoint: ({midP2.X:F2},{midP2.Y:F2},{midP2.Z:F2})");
                            System.Diagnostics.Debug.WriteLine($"Translation vector: ({moveVecP2.X:F3},{moveVecP2.Y:F3},{moveVecP2.Z:F3})");

                            // Update Pipe2
                            XYZ newP2Start = new XYZ(p2Start.X + moveVecP2.X, p2Start.Y + moveVecP2.Y, p2Start.Z + moveVecP2.Z);
                            XYZ newP2End = new XYZ(p2End.X + moveVecP2.X, p2End.Y + moveVecP2.Y, p2End.Z + moveVecP2.Z);
                            Line newCurve2 = Line.CreateBound(newP2Start, newP2End);
                            (pipe2.Location as LocationCurve).Curve = newCurve2;

                            System.Diagnostics.Debug.WriteLine($"Moved Pipe {pairIdx + 2} by: ({moveVecP2.X:F3},{moveVecP2.Y:F3},{moveVecP2.Z:F3})");
                            continue;
                        }

                        System.Diagnostics.Debug.WriteLine($"Closest point on Pipe {pairIdx + 1}: ({p1OnLine.X:F2},{p1OnLine.Y:F2},{p1OnLine.Z:F2})");
                        System.Diagnostics.Debug.WriteLine($"Closest point on Pipe {pairIdx + 2}: ({p2OnLine.X:F2},{p2OnLine.Y:F2},{p2OnLine.Z:F2})");
                        System.Diagnostics.Debug.WriteLine($"Distance between centerlines: {distance:F6}");
                        System.Diagnostics.Debug.WriteLine($"s: {s:F6}, t: {t:F6}");

                        // Move Pipe2 so p2OnLine aligns with p1OnLine
                        XYZ moveVecP2_2 = new XYZ(p1OnLine.X - p2OnLine.X,
                                                 p1OnLine.Y - p2OnLine.Y,
                                                 p1OnLine.Z - p2OnLine.Z);

                        // Update Pipe2
                        XYZ newP2Start_2 = new XYZ(p2Start.X + moveVecP2_2.X, p2Start.Y + moveVecP2_2.Y, p2Start.Z + moveVecP2_2.Z);
                        XYZ newP2End_2 = new XYZ(p2End.X + moveVecP2_2.X, p2End.Y + moveVecP2_2.Y, p2End.Z + moveVecP2_2.Z);
                        Line newCurve2_2 = Line.CreateBound(newP2Start_2, newP2End_2);
                        (pipe2.Location as LocationCurve).Curve = newCurve2_2;

                        System.Diagnostics.Debug.WriteLine($"Moved Pipe {pairIdx + 2} by: ({moveVecP2_2.X:F3},{moveVecP2_2.Y:F3},{moveVecP2_2.Z:F3})");
                    }

                    tx.Commit();
                    TaskDialog.Show("pyRevit", $"Series alignment complete! {pipes.Count - 1} pairs processed.");
                    System.Diagnostics.Debug.WriteLine("\n=== Series Alignment Complete ===");
                    return Result.Succeeded;
                }
                catch (Exception e)
                {
                    tx.RollBack();
                    TaskDialog.Show("pyRevit", $"Error during alignment: {e.Message}");
                    return Result.Failed;
                }
            }
        }

        /// <summary>
        /// Get current centerline endpoints of a pipe
        /// </summary>
        private (XYZ, XYZ) GetPipeGeometry(Element pipe)
        {
            try
            {
                LocationCurve location = pipe.Location as LocationCurve;
                Curve curve = location.Curve;
                XYZ pStart = curve.GetEndPoint(0);
                XYZ pEnd = curve.GetEndPoint(1);
                return (pStart, pEnd);
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting pipe geometry: {e.Message}");
                return (null, null);
            }
        }

        /// <summary>
        /// Calculate closest points between two lines in 3D
        /// Returns: (p1OnLine, p2OnLine, distance, s, t)
        /// </summary>
        private (XYZ, XYZ, double, double, double) LineIntersection3D(
            XYZ p1Start, XYZ p1End, XYZ p2Start, XYZ p2End, double tolerance = 1e-9)
        {
            try
            {
                // Direction vectors
                XYZ d1 = new XYZ(p1End.X - p1Start.X, p1End.Y - p1Start.Y, p1End.Z - p1Start.Z);
                XYZ d2 = new XYZ(p2End.X - p2Start.X, p2End.Y - p2Start.Y, p2End.Z - p2Start.Z);

                // w0 = p1Start -> p2Start
                XYZ w0 = new XYZ(p1Start.X - p2Start.X, p1Start.Y - p2Start.Y, p1Start.Z - p2Start.Z);

                double a = d1.DotProduct(d1);
                double b = d1.DotProduct(d2);
                double c = d2.DotProduct(d2);
                double d = d1.DotProduct(w0);
                double e = d2.DotProduct(w0);

                double denom = a * c - b * b;

                if (Math.Abs(denom) < tolerance)
                {
                    // Lines are parallel
                    return (null, null, 0, 0, 0);
                }

                double s = (b * e - c * d) / denom;
                double t = (a * e - b * d) / denom;

                // Points on lines
                XYZ p1OnLine = new XYZ(p1Start.X + s * d1.X, p1Start.Y + s * d1.Y, p1Start.Z + s * d1.Z);
                XYZ p2OnLine = new XYZ(p2Start.X + t * d2.X, p2Start.Y + t * d2.Y, p2Start.Z + t * d2.Z);

                // Distance between closest points
                double dx = p2OnLine.X - p1OnLine.X;
                double dy = p2OnLine.Y - p1OnLine.Y;
                double dz = p2OnLine.Z - p1OnLine.Z;
                double dist = Math.Sqrt(dx * dx + dy * dy + dz * dz);

                return (p1OnLine, p2OnLine, dist, s, t);
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine($"Error in intersection calculation: {e.Message}");
                return (null, null, 0, 0, 0);
            }
        }
    }
}

/* 
**PYREVIT → C# CONVERSIONS APPLIED:**
1. `__revit__` → `commandData.Application`
2. `doc` / `uidoc` → Retrieved from `ExternalCommandData`
3. `forms.alert()` → `TaskDialog.Show()`
4. `print()` → `System.Diagnostics.Debug.WriteLine()`
5. Python `try/except OperationCanceledException` → C# `catch (Autodesk.Revit.Exceptions.OperationCanceledException)`
6. Python tuple returns `(a, b, c)` → C# ValueTuple `(XYZ, XYZ, double, double, double)`
7. `element.Id.IntegerValue` → `element.Id.Value` (Revit 2022+ deprecated API fix)
8. Python `while True:` loop → C# `while (true)` loop
9. Transaction pattern: Python `Transaction(doc, name)` → C# `using (Transaction tx = new Transaction(doc, name))`

**THAM KHẢO TỪ Commands/As/:**
- IExternalCommand pattern structure
- Transaction handling with using statement
- OperationCanceledException handling
- Debug output patterns
- TaskDialog for user feedback

**IMPORTANT NOTES:**
- Giữ nguyên logic thuật toán tính intersection 3D từ Python
- Sử dụng ValueTuple (C# 7.0+) cho multiple return values
- Pipe selection loop giữ nguyên như Python (select one by one until ESC)
- Alignment logic giữ nguyên: first pipe fixed, move subsequent pipes
- Parallel lines handling: sử dụng midpoint alignment
*/
