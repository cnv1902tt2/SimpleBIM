using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

// Alias
using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;

namespace SimpleBIM.Commands.As
{
    /// <summary>
    /// BATCH INTEGRATED FLOOR TOOL - CREATE MULTIPLE FLOORS WITH SLOPE
    /// Tạo hàng loạt Floor từ nhiều mặt nghiêng với Smart Offset
    /// Slope Arrow xuất phát TỪ lowest edge
    /// 100% Python Equivalent Version
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class FinishFaceFloors : IExternalCommand
    {
        // Instance variables
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

        // ============================================================================
        // UTILITIES
        // ============================================================================

        private void ShowMessage(string message, string title = "Notification")
        {
            WinForms.MessageBox.Show(message, title,
                WinForms.MessageBoxButtons.OK,
                WinForms.MessageBoxIcon.Information);
        }

        private void ShowError(string message, string title = "Error")
        {
            WinForms.MessageBox.Show(message, title,
                WinForms.MessageBoxButtons.OK,
                WinForms.MessageBoxIcon.Error);
        }

        private void ShowWarning(string message, string title = "Warning")
        {
            WinForms.MessageBox.Show(message, title,
                WinForms.MessageBoxButtons.OK,
                WinForms.MessageBoxIcon.Warning);
        }

        private void PrintDebug(string msg)
        {
            System.Diagnostics.Debug.WriteLine(msg);
        }

        // ============================================================================
        // THICKNESS EXTRACTION - PYTHON EQUIVALENT (4 METHODS)
        // ============================================================================

        private double GetFloorTypeThickness(FloorType floorType)
        {
            try
            {
                // Method 1: Compound Structure
                try
                {
                    CompoundStructure compound = floorType.GetCompoundStructure();
                    if (compound != null)
                    {
                        double thickness = compound.GetWidth();
                        if (thickness > 0)
                        {
                            return thickness;
                        }
                    }
                }
                catch { }

                // Method 2: Thickness Parameter
                try
                {
                    Parameter param = floorType.LookupParameter("Thickness");
                    if (param != null && param.HasValue && param.AsDouble() > 0)
                    {
                        return param.AsDouble();
                    }
                }
                catch { }

                // Method 3: Search All Parameters
                try
                {
                    IList<Parameter> parameters = floorType.GetOrderedParameters();
                    foreach (Parameter param in parameters)
                    {
                        try
                        {
                            if (param.Definition != null &&
                                param.Definition.Name != null &&
                                param.Definition.Name.IndexOf("thickness", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                if (param.StorageType == StorageType.Double && param.HasValue)
                                {
                                    double value = param.AsDouble();
                                    if (value > 0)
                                    {
                                        return value;
                                    }
                                }
                            }
                        }
                        catch { continue; }
                    }
                }
                catch { }

                // Method 4: Sum Layer Thicknesses
                try
                {
                    CompoundStructure compound = floorType.GetCompoundStructure();
                    if (compound != null && compound.LayerCount > 0)
                    {
                        double total = 0.0;
                        for (int i = 0; i < compound.LayerCount; i++)
                        {
                            try
                            {
                                total += compound.GetLayerWidth(i);
                            }
                            catch { continue; }
                        }
                        if (total > 0)
                        {
                            return total;
                        }
                    }
                }
                catch { }

                return 0.0;
            }
            catch
            {
                return 0.0;
            }
        }

        // ============================================================================
        // FACE ANALYSIS - PYTHON EQUIVALENT
        // ============================================================================

        private bool IsFaceSloped(Face face)
        {
            try
            {
                BoundingBoxUV bbox = face.GetBoundingBox();
                if (bbox != null)
                {
                    UV centerUV = new UV(
                        (bbox.Min.U + bbox.Max.U) / 2.0,
                        (bbox.Min.V + bbox.Max.V) / 2.0
                    );
                    XYZ normal = face.ComputeNormal(centerUV);
                    double angleTolerance = Math.Cos(Math.PI / 36.0); // 5 degrees
                    return Math.Abs(normal.Z) <= angleTolerance;
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        private int GetFaceNormalDirection(Face face)
        {
            // Return: 1 = TOP, -1 = BOTTOM, 0 = UNKNOWN
            try
            {
                BoundingBoxUV bbox = face.GetBoundingBox();
                if (bbox != null)
                {
                    UV centerUV = new UV(
                        (bbox.Min.U + bbox.Max.U) / 2.0,
                        (bbox.Min.V + bbox.Max.V) / 2.0
                    );
                    XYZ normal = face.ComputeNormal(centerUV);
                    if (normal.Z > 0.5)
                        return 1;  // TOP
                    else if (normal.Z < -0.5)
                        return -1; // BOTTOM
                }
                return 0; // UNKNOWN
            }
            catch
            {
                return 0;
            }
        }

        private double? GetSlopeFromFaceNormal(Face face)
        {
            try
            {
                BoundingBoxUV bbox = face.GetBoundingBox();
                if (bbox != null)
                {
                    UV centerUV = new UV(
                        (bbox.Min.U + bbox.Max.U) / 2.0,
                        (bbox.Min.V + bbox.Max.V) / 2.0
                    );
                    XYZ normal = face.ComputeNormal(centerUV);
                    double nzAbs = Math.Abs(Math.Max(-1.0, Math.Min(1.0, normal.Z)));
                    double angleRad = Math.Acos(nzAbs);
                    double angleDeg = angleRad * (180.0 / Math.PI);
                    if (angleDeg > 90)
                    {
                        angleDeg = 180 - angleDeg;
                    }
                    return angleDeg;
                }
                return null;
            }
            catch
            {
                return null;
            }
        }

        // ============================================================================
        // POINT EXTRACTION & PROCESSING - PYTHON EQUIVALENT WITH TRANSFORM
        // ============================================================================

        private List<XYZ> ExtractAllFacePoints(Face face, Element element)
        {
            try
            {
                List<XYZ> allPoints = new List<XYZ>();

                // Extract all points from edge loops
                foreach (EdgeArray edgeLoop in face.EdgeLoops)
                {
                    foreach (Edge edge in edgeLoop)
                    {
                        Curve curve = edge.AsCurve();
                        if (curve != null)
                        {
                            allPoints.Add(curve.GetEndPoint(0));
                            allPoints.Add(curve.GetEndPoint(1));
                        }
                    }
                }

                // Apply transform like Python version
                List<XYZ> originalPoints = new List<XYZ>();
                try
                {
                    Transform transform = null;

                    if (element is FamilyInstance familyInstance)
                    {
                        transform = familyInstance.GetTransform();
                    }
                    else
                    {
                        transform = Transform.Identity;
                    }

                    if (transform != null && !transform.IsIdentity)
                    {
                        foreach (XYZ pt in allPoints)
                        {
                            originalPoints.Add(transform.OfPoint(pt));
                        }
                    }
                    else
                    {
                        originalPoints = allPoints;
                    }
                }
                catch
                {
                    originalPoints = allPoints;
                }

                return originalPoints;
            }
            catch
            {
                return null;
            }
        }

        private List<XYZ> CleanupDuplicatePoints(List<XYZ> points)
        {
            if (points == null || points.Count == 0)
                return new List<XYZ>();

            List<XYZ> uniquePoints = new List<XYZ>();
            double tolerance = 0.001;

            foreach (XYZ pt in points)
            {
                bool isDuplicate = false;
                foreach (XYZ existing in uniquePoints)
                {
                    double dx = pt.X - existing.X;
                    double dy = pt.Y - existing.Y;
                    double dz = pt.Z - existing.Z;
                    double dist3d = Math.Sqrt(dx * dx + dy * dy + dz * dz);
                    if (dist3d < tolerance)
                    {
                        isDuplicate = true;
                        break;
                    }
                }
                if (!isDuplicate)
                {
                    uniquePoints.Add(pt);
                }
            }

            // Sort points by angle (Python equivalent)
            if (uniquePoints.Count >= 3)
            {
                double centerX = uniquePoints.Sum(pt => pt.X) / uniquePoints.Count;
                double centerY = uniquePoints.Sum(pt => pt.Y) / uniquePoints.Count;

                double GetAngle(XYZ point)
                {
                    return Math.Atan2(point.Y - centerY, point.X - centerX);
                }

                uniquePoints = uniquePoints.OrderBy(GetAngle).ToList();
            }

            return uniquePoints;
        }

        // ============================================================================
        // LEVEL EXTRACTION - PYTHON EQUIVALENT (5 METHODS)
        // ============================================================================

        private (Level level, bool success, string message) GetLevelFromElement(Element element)
        {
            try
            {
                // Method 1: Direct LevelId
                try
                {
                    if (element.LevelId != null &&
                        element.LevelId != ElementId.InvalidElementId)
                    {
                        Level level = _doc.GetElement(element.LevelId) as Level;
                        if (level != null)
                        {
                            return (level, true, $"Using element LevelId: {level.Name}");
                        }
                    }
                }
                catch { }

                // Method 2: Level parameter
                try
                {
                    Parameter levelParam = element.LookupParameter("Level");
                    if (levelParam != null && levelParam.HasValue)
                    {
                        ElementId levelId = levelParam.AsElementId();
                        if (levelId != null && levelId != ElementId.InvalidElementId)
                        {
                            Level level = _doc.GetElement(levelId) as Level;
                            if (level != null)
                            {
                                return (level, true, $"Using Level parameter: {level.Name}");
                            }
                        }
                    }
                }
                catch { }

                // Method 3: Base Level parameter
                try
                {
                    Parameter baseLevelParam = element.LookupParameter("Base Level");
                    if (baseLevelParam != null && baseLevelParam.HasValue)
                    {
                        ElementId levelId = baseLevelParam.AsElementId();
                        if (levelId != null && levelId != ElementId.InvalidElementId)
                        {
                            Level level = _doc.GetElement(levelId) as Level;
                            if (level != null)
                            {
                                return (level, true, $"Using Base Level parameter: {level.Name}");
                            }
                        }
                    }
                }
                catch { }

                // Method 4: Get lowest level
                try
                {
                    FilteredElementCollector collector = new FilteredElementCollector(_doc)
                        .OfClass(typeof(Level));
                    List<Level> allLevels = collector.Cast<Level>().ToList();

                    if (allLevels.Count > 0)
                    {
                        Level lowestLevel = allLevels.OrderBy(l => l.Elevation).First();
                        return (lowestLevel, true, $"Using lowest level: {lowestLevel.Name}");
                    }
                }
                catch { }

                // Method 5: Use default level
                try
                {
                    FilteredElementCollector collector = new FilteredElementCollector(_doc)
                        .OfClass(typeof(Level));
                    List<Level> allLevels = collector.Cast<Level>().ToList();

                    foreach (Level lvl in allLevels)
                    {
                        if (lvl.Name == "0" || lvl.Name == "Ground Floor")
                        {
                            return (lvl, true, $"Using default level: {lvl.Name}");
                        }
                    }

                    if (allLevels.Count > 0)
                    {
                        return (allLevels[0], true, $"Using first available level: {allLevels[0].Name}");
                    }
                }
                catch { }

                return (null, false, "Cannot find any level!");
            }
            catch (Exception ex)
            {
                return (null, false, $"Error getting level: {ex.Message}");
            }
        }

        // ============================================================================
        // FLOOR CREATION WITH SLOPE - PYTHON EQUIVALENT
        // ============================================================================

        private (Floor floor, bool success, string message) CreateSlopedFloor(
            List<XYZ> points, Level level, FloorType floorType,
            double slopeDegrees, bool isStructural, int faceIdx)
        {
            try
            {
                double minZ = points.Min(pt => pt.Z);
                double maxZ = points.Max(pt => pt.Z);
                double levelElevation = level.Elevation;

                double thickness = GetFloorTypeThickness(floorType);
                PrintDebug($"[FACE {faceIdx + 1}] Floor type thickness: {thickness:F3}");

                double elevationDiff = minZ - levelElevation;
                PrintDebug($"[FACE {faceIdx + 1}] Min Z: {minZ:F3}, Level Elev: {levelElevation:F3}, Diff: {elevationDiff:F3}");

                // ========================================================
                // STEP 1: Create CurveLoop at level elevation
                // ========================================================
                List<XYZ> flattenedPoints = points
                    .Select(pt => new XYZ(pt.X, pt.Y, levelElevation))
                    .ToList();

                CurveLoop curveLoop = new CurveLoop();
                int curveCount = 0;

                for (int i = 0; i < flattenedPoints.Count; i++)
                {
                    XYZ startPt = flattenedPoints[i];
                    XYZ endPt = flattenedPoints[(i + 1) % flattenedPoints.Count];
                    double dx = endPt.X - startPt.X;
                    double dy = endPt.Y - startPt.Y;
                    double dist = Math.Sqrt(dx * dx + dy * dy);

                    if (dist > 0.001)
                    {
                        curveLoop.Append(Line.CreateBound(startPt, endPt));
                        curveCount++;
                    }
                }

                if (curveCount < 3)
                {
                    return (null, false, $"CurveLoop has {curveCount} curves (need 3+)");
                }

                // ========================================================
                // STEP 2: Find lowest edge
                // ========================================================
                double tolerance = 0.001;
                List<XYZ> lowestPoints = points
                    .Where(p => Math.Abs(p.Z - minZ) < tolerance)
                    .ToList();

                PrintDebug($"[FACE {faceIdx + 1}] Min Z: {minZ:F3}, Lowest points: {lowestPoints.Count}");

                // Find curves corresponding to lowest edge
                List<(Curve curve, double length, XYZ start, XYZ end)> lowestCurves =
                    new List<(Curve, double, XYZ, XYZ)>();

                foreach (Curve curve in curveLoop)
                {
                    XYZ startPt = curve.GetEndPoint(0);
                    XYZ endPt = curve.GetEndPoint(1);

                    bool startIsLow = lowestPoints.Any(lp =>
                        Math.Abs(startPt.X - lp.X) < 0.01 &&
                        Math.Abs(startPt.Y - lp.Y) < 0.01);

                    bool endIsLow = lowestPoints.Any(lp =>
                        Math.Abs(endPt.X - lp.X) < 0.01 &&
                        Math.Abs(endPt.Y - lp.Y) < 0.01);

                    if (startIsLow && endIsLow)
                    {
                        lowestCurves.Add((curve, curve.Length, startPt, endPt));
                    }
                }

                XYZ tailStart, tailEnd;

                if (lowestCurves.Count == 0)
                {
                    PrintDebug($"[FACE {faceIdx + 1}] No lowest edge found");

                    // Fallback: use longest curve
                    Curve longestCurve = null;
                    double maxLength = 0.0;
                    foreach (Curve curve in curveLoop)
                    {
                        if (curve.Length > maxLength)
                        {
                            maxLength = curve.Length;
                            longestCurve = curve;
                        }
                    }

                    if (longestCurve != null)
                    {
                        tailStart = longestCurve.GetEndPoint(0);
                        tailEnd = longestCurve.GetEndPoint(1);
                    }
                    else
                    {
double fallbackCenterX = flattenedPoints.Sum(pt => pt.X) / flattenedPoints.Count;
                        double fallbackCenterY = flattenedPoints.Sum(pt => pt.Y) / flattenedPoints.Count;
                        tailStart = new XYZ(fallbackCenterX - 2.0, fallbackCenterY, levelElevation);
                        tailEnd = new XYZ(fallbackCenterX + 2.0, fallbackCenterY, levelElevation);
                    }
                }
                else
                {
                    var bestCurve = lowestCurves.OrderByDescending(x => x.length).First();
                    tailStart = bestCurve.start;
                    tailEnd = bestCurve.end;
                    PrintDebug($"[FACE {faceIdx + 1}] Using lowest edge, length: {bestCurve.length:F3}");
                }

                // ========================================================
                // STEP 3: Calculate Slope Arrow - FROM lowest edge
                // ========================================================
                double edgeDirX = tailEnd.X - tailStart.X;
                double edgeDirY = tailEnd.Y - tailStart.Y;
                double edgeLength = Math.Sqrt(edgeDirX * edgeDirX + edgeDirY * edgeDirY);

                if (edgeLength > 0.001)
                {
                    edgeDirX /= edgeLength;
                    edgeDirY /= edgeLength;
                }

                // Perpendicular direction (rotate 90° CCW)
                double perpDirX = -edgeDirY;
                double perpDirY = edgeDirX;

                // Center ON lowest edge
                double midX = (tailStart.X + tailEnd.X) / 2.0;
                double midY = (tailStart.Y + tailEnd.Y) / 2.0;
                double centerZ = levelElevation;
                XYZ slopeArrowStart = new XYZ(midX, midY, centerZ);

                // Dynamic offset based on edge length
                double offset = edgeLength * 0.15;  // 15% of edge length
                offset = Math.Max(0.5, Math.Min(5.0, offset));  // Clamp: 0.5 ≤ offset ≤ 5.0

                XYZ slopeArrowEnd = new XYZ(
                    midX + perpDirX * offset,
                    midY + perpDirY * offset,
                    centerZ
                );

                Line slopeArrow = Line.CreateBound(slopeArrowStart, slopeArrowEnd);

                PrintDebug($"[FACE {faceIdx + 1}] Slope arrow: from ({slopeArrowStart.X:F2}, {slopeArrowStart.Y:F2}) " +
                          $"to ({slopeArrowEnd.X:F2}, {slopeArrowEnd.Y:F2}), length={offset:F2}");

                // ========================================================
                // STEP 4: Create floor with slope
                // ========================================================
                ElementId floorTypeId = floorType?.Id ?? Floor.GetDefaultFloorType(_doc, isStructural);
                double slopeRadians = slopeDegrees * (Math.PI / 180.0);

                try
                {
                    PrintDebug($"[FACE {faceIdx + 1}] Creating floor at level elevation: slope={slopeDegrees:F1}deg");

                    Floor floor = Floor.Create(
                        _doc,
                        new List<CurveLoop> { curveLoop },
                        floorTypeId,
                        level.Id,
                        isStructural,
                        slopeArrow,
                        slopeRadians
                    );

                    if (floor == null)
                    {
                        PrintDebug($"[FACE {faceIdx + 1}] Floor.Create returned null");
                        return (null, false, "Floor.Create() returned null");
                    }

                    PrintDebug($"[FACE {faceIdx + 1}] Floor created: ID={floor.Id.IntegerValue}");

                    // ========================================================
                    // STEP 5: Adjust floor elevation
                    // ========================================================
                    double finalOffset = elevationDiff;
                    PrintDebug($"[FACE {faceIdx + 1}] Final offset needed: {finalOffset:F3}");

                    bool offsetAssigned = false;

                    string[] offsetParamNames = new string[]
                    {
                        "Base Offset",
                        "Level Offset",
                        "Offset",
                        "Height Offset",
                        "Base Level Offset"
                    };

                    foreach (string paramName in offsetParamNames)
                    {
                        try
                        {
                            Parameter param = floor.LookupParameter(paramName);
                            if (param != null && !param.IsReadOnly)
                            {
                                param.Set(finalOffset);
                                PrintDebug($"[FACE {faceIdx + 1}] ✓ Set {paramName} = {finalOffset:F3}");
                                offsetAssigned = true;
                                break;
                            }
                        }
                        catch { continue; }
                    }

                    if (!offsetAssigned)
                    {
                        try
                        {
                            PrintDebug($"[FACE {faceIdx + 1}] Using Move command to adjust elevation...");

                            XYZ moveVector = new XYZ(0, 0, finalOffset);
                            ElementTransformUtils.MoveElement(_doc, floor.Id, moveVector);

                            PrintDebug($"[FACE {faceIdx + 1}] ✓ Moved floor by {finalOffset:F3}");
                            offsetAssigned = true;
                        }
                        catch (Exception moveEx)
                        {
                            PrintDebug($"[FACE {faceIdx + 1}] Error moving floor: {moveEx.Message}");
                        }
                    }

                    // ========================================================
                    // STEP 6: Update slope
                    // ========================================================
                    try
                    {
                        Parameter slopeParam = floor.LookupParameter("Slope");
                        if (slopeParam != null && !slopeParam.IsReadOnly)
                        {
                            double slopeRatio = Math.Tan(slopeRadians);
                            slopeParam.Set(slopeRatio);
                            PrintDebug($"[FACE {faceIdx + 1}] ✓ Updated slope ratio: {slopeRatio:F3}");
                        }
                    }
                    catch { }

                    PrintDebug($"[FACE {faceIdx + 1}] ✓ Floor created successfully, matches original face");

                    return (floor, true, $"Floor created ID: {floor.Id.IntegerValue}");
                }
                catch (Exception ex)
                {
                    PrintDebug($"[FACE {faceIdx + 1}] Floor.Create error: {ex.Message}");
                    return (null, false, $"Floor.Create error: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                PrintDebug($"[FACE {faceIdx + 1}] Error: {ex.Message}");
                return (null, false, $"Error: {ex.Message}");
            }
        }

        // ============================================================================
        // FLOOR TYPE SELECTOR - PYTHON EQUIVALENT
        // ============================================================================

        private List<(string name, FloorType floorType)> GetAllFloorTypes()
        {
            try
            {
                FilteredElementCollector collector = new FilteredElementCollector(_doc)
                    .OfClass(typeof(FloorType));

                List<(string, FloorType)> floorTypeList = new List<(string, FloorType)>();

                foreach (FloorType ft in collector)
                {
                    try
                    {
                        Parameter nameParam = ft.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME);
                        string typeName;
                        if (nameParam != null && nameParam.HasValue)
                        {
                            typeName = nameParam.AsString();
                        }
                        else
                        {
                            typeName = $"Floor Type {ft.Id.IntegerValue}";
                        }

                        floorTypeList.Add((typeName, ft));
                    }
                    catch
                    {
                        floorTypeList.Add(($"Floor Type {ft.Id.IntegerValue}", ft));
                    }
                }

                return floorTypeList;
            }
            catch
            {
                return new List<(string, FloorType)>();
            }
        }

        private (FloorType floorType, string floorTypeName) ShowFloorTypeSelector()
        {
            try
            {
                List<(string name, FloorType floorType)> floorTypes = GetAllFloorTypes();
                if (floorTypes == null || floorTypes.Count == 0)
                {
                    ShowError("No floor types found!");
                    return (null, null);
                }

                Dictionary<string, FloorType> typeDict = floorTypes
                    .ToDictionary(x => x.name, x => x.floorType);
                List<string> typeNames = typeDict.Keys.OrderBy(x => x).ToList();

                using (WinForms.Form form = new WinForms.Form())
                {
                    form.Text = "Select Floor Type";
                    form.Width = 450;
                    form.Height = 380;
                    form.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
                    form.MaximizeBox = false;
                    form.MinimizeBox = false;
                    form.StartPosition = WinForms.FormStartPosition.CenterScreen;

                    WinForms.Label label = new WinForms.Label
                    {
                        Text = "Select floor type:",
                        Location = new Drawing.Point(15, 15),
                        Width = 400,
                        Height = 25
                    };
                    form.Controls.Add(label);

                    WinForms.ListBox listbox = new WinForms.ListBox
                    {
                        Location = new Drawing.Point(15, 45),
                        Width = 400,
                        Height = 230
                    };

                    int defaultIdx = 0;
                    for (int i = 0; i < typeNames.Count; i++)
                    {
                        string name = typeNames[i];
                        listbox.Items.Add(name);
                        if (name.IndexOf("Generic", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            name.IndexOf("Basic", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            name.IndexOf("Standard", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            defaultIdx = i;
                        }
                    }

                    listbox.SelectedIndex = defaultIdx;
                    form.Controls.Add(listbox);

                    FloorType selectedFloorType = null;
                    string selectedFloorTypeName = null;

                    WinForms.Button btnOk = new WinForms.Button
                    {
                        Text = "OK",
                        Location = new Drawing.Point(155, 290),
                        Width = 90,
                        Height = 35
                    };
                    btnOk.Click += (s, e) =>
                    {
                        if (listbox.SelectedIndex >= 0)
                        {
                            string name = typeNames[listbox.SelectedIndex];
                            selectedFloorType = typeDict[name];
                            selectedFloorTypeName = name;
                        }
                        form.DialogResult = WinForms.DialogResult.OK;
                        form.Close();
                    };
                    form.Controls.Add(btnOk);
                    form.AcceptButton = btnOk;

                    WinForms.Button btnCancel = new WinForms.Button
                    {
                        Text = "Cancel",
                        Location = new Drawing.Point(260, 290),
                        Width = 90,
                        Height = 35
                    };
                    btnCancel.Click += (s, e) =>
                    {
                        form.DialogResult = WinForms.DialogResult.Cancel;
                        form.Close();
                    };
                    form.Controls.Add(btnCancel);
                    form.CancelButton = btnCancel;

                    WinForms.DialogResult result = form.ShowDialog();
                    if (result == WinForms.DialogResult.OK)
                    {
                        return (selectedFloorType, selectedFloorTypeName);
                    }
                    else
                    {
                        return (null, null);
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError($"Error: {ex.Message}");
                return (null, null);
            }
        }

        // ============================================================================
        // BATCH FACE SELECTION - PYTHON EQUIVALENT
        // ============================================================================

        private class FaceData
        {
            public Element Element { get; set; }
            public Face Face { get; set; }
            public Reference Reference { get; set; }
            public int Index { get; set; }
        }

        private List<FaceData> PickMultipleFaces()
        {
            List<FaceData> facesData = new List<FaceData>();

            ShowMessage(
                "BATCH FLOOR CREATOR\n\n" +
                "Instructions:\n" +
                "1. Click OK to start\n" +
                "2. Select sloped faces one by one\n" +
                "3. Press ESC when done\n",
                "Batch Mode"
            );

            try
            {
                while (true)
                {
                    try
                    {
                        Reference reference = _uidoc.Selection.PickObject(
                            ObjectType.Face,
                            "Select sloped face (ESC to finish)"
                        );

                        Element element = _doc.GetElement(reference.ElementId);
                        GeometryObject geometryObject = element.GetGeometryObjectFromReference(reference);

                        if (geometryObject is Face face)
                        {
                            if (IsFaceSloped(face))
                            {
                                facesData.Add(new FaceData
                                {
                                    Element = element,
                                    Face = face,
                                    Reference = reference,
                                    Index = facesData.Count + 1
                                });
                            }
                            else
                            {
                                ShowWarning("Selected face is not sloped!");
                            }
                        }
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                    {
                        // User pressed ESC - exit loop
                        break;
                    }
                    catch
                    {
                        break;
                    }
                }
            }
            catch { }

            return facesData;
        }

        // ============================================================================
        // MAIN WORKFLOW - PYTHON EQUIVALENT
        // ============================================================================

        private Result RunMain()
        {
            try
            {
                // Step 1: Pick multiple faces
                List<FaceData> facesData = PickMultipleFaces();

                if (facesData == null || facesData.Count == 0)
                {
                    ShowError("No faces selected!");
                    return Result.Cancelled;
                }

                ShowMessage($"Total faces selected: {facesData.Count}", "Summary");

                // Step 2: Select floor type
                var (floorType, floorTypeName) = ShowFloorTypeSelector();
                if (floorType == null)
                {
                    ShowError("No floor type selected!");
                    return Result.Cancelled;
                }

                // Step 3: Create floors in batch
                using (Transaction trans = new Transaction(_doc, "Create Multiple Floors with Slope"))
                {
                    trans.Start();

                    try
                    {
                        int successCount = 0;
                        int failedCount = 0;
                        List<string> results = new List<string>();

                        for (int idx = 0; idx < facesData.Count; idx++)
                        {
                            FaceData faceData = facesData[idx];

                            try
                            {
                                // Extract face info
                                List<XYZ> originalPoints = ExtractAllFacePoints(
                                    faceData.Face, faceData.Element);

                                if (originalPoints == null || originalPoints.Count == 0)
                                {
                                    results.Add($"Face {idx + 1}: Cannot extract points");
                                    failedCount++;
                                    continue;
                                }

                                List<XYZ> cleanPoints = CleanupDuplicatePoints(originalPoints);
                                if (cleanPoints.Count < 3)
                                {
                                    results.Add($"Face {idx + 1}: Insufficient points ({cleanPoints.Count})");
                                    failedCount++;
                                    continue;
                                }

                                // Get slope
                                double? slopeDegreesNullable = GetSlopeFromFaceNormal(faceData.Face);
                                if (!slopeDegreesNullable.HasValue)
                                {
                                    results.Add($"Face {idx + 1}: Cannot calculate slope");
                                    failedCount++;
                                    continue;
                                }
                                double slopeDegrees = slopeDegreesNullable.Value;

                                // Get level
                                var (level, levelOk, levelMsg) = GetLevelFromElement(faceData.Element);
                                if (!levelOk || level == null)
                                {
                                    results.Add($"Face {idx + 1}: {levelMsg}");
                                    failedCount++;
                                    continue;
                                }

                                // Create sloped floor
                                var (floor, floorOk, floorMsg) = CreateSlopedFloor(
                                    cleanPoints, level, floorType, slopeDegrees, true, idx);

                                if (!floorOk || floor == null)
                                {
                                    results.Add($"Face {idx + 1}: {floorMsg}");
                                    failedCount++;
                                    continue;
                                }

                                results.Add($"Face {idx + 1}: OK (ID: {floor.Id.IntegerValue}, Slope: {slopeDegrees:F1}deg)");
                                successCount++;
                            }
                            catch (Exception ex)
                            {
                                results.Add($"Face {idx + 1}: Exception - {ex.Message}");
                                failedCount++;
                            }
                        }

                        trans.Commit();

                        // Show results
                        string resultMsg = "BATCH FLOOR CREATION COMPLETE\n\n";
                        resultMsg += $"SUCCESS: {successCount}\n";
                        resultMsg += $"FAILED: {failedCount}\n\n";
                        resultMsg += "Details:\n";
                        resultMsg += string.Join("\n", results);

                        ShowMessage(resultMsg, "BATCH RESULTS");
                        return Result.Succeeded;
                    }
                    catch (Exception ex)
                    {
                        try
                        {
                            trans.RollBack();
                        }
                        catch { }

                        ShowError($"Transaction Error: {ex.Message}");
                        return Result.Failed;
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError($"Main Error: {ex.Message}");
                return Result.Failed;
            }
        }
    }
}