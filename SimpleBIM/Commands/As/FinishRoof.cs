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

// Alias để tránh ambiguous references
using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;

namespace SimpleBIM.Commands.As
{
    /// <summary>
    /// BATCH INTEGRATED ROOF TOOL - CREATE MULTIPLE ROOFS
    /// Tạo hàng loạt Roof từ nhiều mặt nghiêng với Smart Offset
    /// 100% Python Equivalent Version
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class FinishRoof : IExternalCommand
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

        // ============================================================================
        // THICKNESS EXTRACTION (4 METHODS) - PYTHON EQUIVALENT
        // ============================================================================

        private double GetRoofTypeThickness(RoofType roofType)
        {
            try
            {
                // Method 1: Compound Structure
                try
                {
                    CompoundStructure compound = roofType.GetCompoundStructure();
                    if (compound != null)
                    {
                        // Revit 2020+: GetWidth() thay cho GetTotalThickness()
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
                    Parameter param = roofType.LookupParameter("Thickness");
                    if (param != null && param.HasValue && param.AsDouble() > 0)
                    {
                        return param.AsDouble();
                    }
                }
                catch { }

                // Method 3: Search All Parameters
                try
                {
                    IList<Parameter> parameters = roofType.GetOrderedParameters();
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
                    CompoundStructure compound = roofType.GetCompoundStructure();
                    if (compound != null && compound.LayerCount > 0)
                    {
                        double total = 0.0;
                        for (int i = 0; i < compound.LayerCount; i++)
                        {
                            try
                            {
                                // Revit 2020+: GetLayerWidth() thay cho GetLayerThickness()
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

                return 0;
            }
            catch
            {
                return 0;
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

                // *** CRITICAL FIX: Apply transform like Python version ***
                List<XYZ> originalPoints = new List<XYZ>();
                try
                {
                    Transform transform = null;

                    // Try to get transform from different element types
                    if (element is FamilyInstance familyInstance)
                    {
                        transform = familyInstance.GetTransform();
                    }
                    else if (element.Location is LocationPoint locationPoint)
                    {
                        transform = locationPoint.Rotation != null
                            ? Transform.CreateRotation(XYZ.BasisZ, locationPoint.Rotation)
                            : Transform.Identity;
                    }
                    else if (element.Location is LocationCurve locationCurve)
                    {
                        // For elements with LocationCurve, typically no rotation needed
                        transform = Transform.Identity;
                    }
                    else
                    {
                        // Default to Identity transform
                        transform = Transform.Identity;
                    }

                    // Apply transform if it's not identity
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
                    // If transform fails, use original points
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

            // Sort points if 4 corners (like Python)
            if (uniquePoints.Count == 4)
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

                // Method 4: Get lowest level (PYTHON EQUIVALENT)
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

                // Method 5: Use default level (PYTHON EQUIVALENT)
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
        // ROOF CREATION - PYTHON EQUIVALENT
        // ============================================================================

        private (RoofBase roof, bool slopeSuccess) CreateValidRoofFootprint(
            List<XYZ> points, Level level, RoofType roofType, double slopeDegrees)
        {
            try
            {
                // Flatten points to level elevation
                List<XYZ> flattenedPoints = points
                    .Select(pt => new XYZ(pt.X, pt.Y, level.Elevation))
                    .ToList();

                // Create curve array
                CurveArray curveArray = new CurveArray();
                for (int i = 0; i < flattenedPoints.Count; i++)
                {
                    XYZ startPt = flattenedPoints[i];
                    XYZ endPt = flattenedPoints[(i + 1) % flattenedPoints.Count];
                    double dx = endPt.X - startPt.X;
                    double dy = endPt.Y - startPt.Y;
                    double distance = Math.Sqrt(dx * dx + dy * dy);
                    if (distance > 0.001)
                    {
                        curveArray.Append(Line.CreateBound(startPt, endPt));
                    }
                }

                if (curveArray.Size < 3)
                {
                    return (null, false);
                }

                // Create roof
                ModelCurveArray modelCurveArray = new ModelCurveArray();
                try
                {
                    RoofBase roof = _doc.Create.NewFootPrintRoof(
                        curveArray, level, roofType, out modelCurveArray);

                    if (roof == null)
                    {
                        return (null, false);
                    }

                    // Set slope on lowest edges
                    bool slopeSuccess = SetSlopeOnLowestEdges(
                        roof, points, slopeDegrees, modelCurveArray);

                    return (roof, slopeSuccess);
                }
                catch
                {
                    return (null, false);
                }
            }
            catch
            {
                return (null, false);
            }
        }

        private bool SetSlopeOnLowestEdges(
            RoofBase roof, List<XYZ> originalPoints,
            double slopeDegrees, ModelCurveArray modelCurveArray)
        {
            try
            {
                double minZ = originalPoints.Min(pt => pt.Z);
                double tolerance = 0.001;
                List<XYZ> lowestPoints = originalPoints
                    .Where(p => Math.Abs(p.Z - minZ) < tolerance)
                    .ToList();

                // Find lowest edges
                List<(ModelCurve element, double length)> lowestEdges =
                    new List<(ModelCurve, double)>();

                foreach (ModelCurve curveElem in modelCurveArray)
                {
                    if (curveElem != null)
                    {
                        Curve curve = curveElem.GeometryCurve;
                        if (curve != null)
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
                                lowestEdges.Add((curveElem, curve.Length));
                            }
                        }
                    }
                }

                if (lowestEdges.Count == 0)
                {
                    return false;
                }

                // Select longest edge
                var bestEdge = lowestEdges.OrderByDescending(x => x.length).First();

                // Set slope defining
                Parameter definesSlopeParam = bestEdge.element
                    .get_Parameter(BuiltInParameter.ROOF_CURVE_IS_SLOPE_DEFINING);

                if (definesSlopeParam != null && !definesSlopeParam.IsReadOnly)
                {
                    definesSlopeParam.Set(1);
                }
                else
                {
                    return false;
                }

                // Set slope angle
                Parameter slopeParam = bestEdge.element
                    .get_Parameter(BuiltInParameter.ROOF_SLOPE);

                if (slopeParam != null && !slopeParam.IsReadOnly)
                {
                    double slopeRadians = slopeDegrees * (Math.PI / 180.0);
                    slopeParam.Set(slopeRadians);
                    return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        // ============================================================================
        // OFFSET CALCULATION - PYTHON EQUIVALENT
        // ============================================================================

        private (bool success, double finalOffset, string offsetType, double thicknessMm)
            OffsetRoofByThickness(
                RoofBase roof, RoofType roofType, int faceNormalDirection,
                List<XYZ> originalPoints, Level level, double slopeDegrees)
        {
            try
            {
                double thickness = GetRoofTypeThickness(roofType);

                double faceMinZ = originalPoints.Min(pt => pt.Z);
                double levelElevation = level.Elevation;
                double elevationDiff = faceMinZ - levelElevation;

                double thicknessOffset;
                string offsetType;

                if (faceNormalDirection == 1) // TOP
                {
                    thicknessOffset = 0.0;
                    offsetType = "TOP";
                }
                else if (faceNormalDirection == -1) // BOTTOM
                {
                    double slopeRadians = slopeDegrees * (Math.PI / 180.0);
                    double cosSlope = Math.Cos(slopeRadians);

                    double adjustedThickness;
                    if (Math.Abs(cosSlope) < 0.001)
                    {
                        adjustedThickness = thickness;
                    }
                    else
                    {
                        adjustedThickness = thickness / cosSlope;
                    }

                    thicknessOffset = -adjustedThickness;
                    offsetType = "BOTTOM";
                }
                else // UNKNOWN
                {
                    thicknessOffset = 0.0;
                    offsetType = "UNKNOWN";
                }

                double finalOffset = elevationDiff + thicknessOffset;

                // Try to set offset parameter
                try
                {
                    Parameter baseOffsetParam = roof.LookupParameter("Base Offset");
                    if (baseOffsetParam == null)
                    {
                        baseOffsetParam = roof.LookupParameter("Base Level Offset");
                    }

                    if (baseOffsetParam != null && !baseOffsetParam.IsReadOnly)
                    {
                        baseOffsetParam.Set(finalOffset);
                        double thicknessMm = UnitUtils.ConvertFromInternalUnits(
                            thickness, UnitTypeId.Millimeters);
                        return (true, finalOffset, offsetType, thicknessMm);
                    }

                    // Search all parameters
                    foreach (Parameter param in roof.Parameters)
                    {
                        try
                        {
                            string paramName = param.Definition?.Name ?? "Unknown";
                            if ((paramName.IndexOf("offset", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                 paramName.IndexOf("base", StringComparison.OrdinalIgnoreCase) >= 0) &&
                                param.StorageType == StorageType.Double &&
                                !param.IsReadOnly)
                            {
                                param.Set(finalOffset);
                                double thicknessMm = UnitUtils.ConvertFromInternalUnits(
                                    thickness, UnitTypeId.Millimeters);
                                return (true, finalOffset, offsetType, thicknessMm);
                            }
                        }
                        catch { continue; }
                    }

                    double thicknessMmFallback = UnitUtils.ConvertFromInternalUnits(
                        thickness, UnitTypeId.Millimeters);
                    return (false, finalOffset, offsetType, thicknessMmFallback);
                }
                catch
                {
                    double thicknessMmFallback = UnitUtils.ConvertFromInternalUnits(
                        thickness, UnitTypeId.Millimeters);
                    return (false, finalOffset, offsetType, thicknessMmFallback);
                }
            }
            catch
            {
                return (false, 0.0, "UNKNOWN", 0.0);
            }
        }

        // ============================================================================
        // UPDATE SLOPE - PYTHON EQUIVALENT
        // ============================================================================

        private bool UpdateRoofSlope(RoofBase roof, double slopeDegrees)
        {
            try
            {
                // Method 1: Slope parameter
                Parameter slopeParam = roof.LookupParameter("Slope");
                if (slopeParam != null && !slopeParam.IsReadOnly)
                {
                    double slopeRatio = Math.Tan(slopeDegrees * (Math.PI / 180.0));
                    slopeParam.Set(slopeRatio);
                    return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        // ============================================================================
        // ROOF TYPE SELECTOR - PYTHON EQUIVALENT
        // ============================================================================

        private List<(string name, RoofType roofType)> GetAllRoofTypes()
        {
            try
            {
                FilteredElementCollector collector = new FilteredElementCollector(_doc)
                    .OfClass(typeof(RoofType));

                List<(string, RoofType)> roofTypeList = new List<(string, RoofType)>();

                foreach (RoofType rt in collector)
                {
                    try
                    {
                        Parameter nameParam = rt.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME);
                        string typeName;
                        if (nameParam != null && nameParam.HasValue)
                        {
                            typeName = nameParam.AsString();
                        }
                        else
                        {
                            typeName = $"Roof Type {rt.Id.IntegerValue}";
                        }

                        roofTypeList.Add((typeName, rt));
                    }
                    catch
                    {
                        roofTypeList.Add(($"Roof Type {rt.Id.IntegerValue}", rt));
                    }
                }

                return roofTypeList;
            }
            catch
            {
                return new List<(string, RoofType)>();
            }
        }

        private (RoofType roofType, string roofTypeName) ShowRoofTypeSelector()
        {
            try
            {
                List<(string name, RoofType roofType)> roofTypes = GetAllRoofTypes();
                if (roofTypes == null || roofTypes.Count == 0)
                {
                    ShowError("No roof types found!");
                    return (null, null);
                }

                Dictionary<string, RoofType> typeDict = roofTypes
                    .ToDictionary(x => x.name, x => x.roofType);
                List<string> typeNames = typeDict.Keys.OrderBy(x => x).ToList();

                using (WinForms.Form form = new WinForms.Form())
                {
                    form.Text = "Select Roof Type";
                    form.Width = 450;
                    form.Height = 380;
                    form.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
                    form.MaximizeBox = false;
                    form.MinimizeBox = false;
                    form.StartPosition = WinForms.FormStartPosition.CenterScreen;

                    WinForms.Label label = new WinForms.Label
                    {
                        Text = "Select roof type:",
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
                            name.IndexOf("Basic", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            defaultIdx = i;
                        }
                    }

                    listbox.SelectedIndex = defaultIdx;
                    form.Controls.Add(listbox);

                    RoofType selectedRoofType = null;
                    string selectedRoofTypeName = null;

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
                            selectedRoofType = typeDict[name];
                            selectedRoofTypeName = name;
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
                        return (selectedRoofType, selectedRoofTypeName);
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
                "BATCH ROOF CREATOR\n\n" +
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

                // Step 2: Select roof type
                var (roofType, roofTypeName) = ShowRoofTypeSelector();
                if (roofType == null)
                {
                    ShowError("No roof type selected!");
                    return Result.Cancelled;
                }

                // Step 3: Create roofs in batch
                using (Transaction trans = new Transaction(_doc, "Create Multiple Roofs"))
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
                                    results.Add($"Face {idx + 1}: Insufficient points");
                                    failedCount++;
                                    continue;
                                }

                                // Get face normal and slope
                                int faceNormalDirection = GetFaceNormalDirection(faceData.Face);
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

                                // Create roof
                                var (roof, slopeSuccess) = CreateValidRoofFootprint(
                                    cleanPoints, level, roofType, slopeDegrees);

                                if (roof == null)
                                {
                                    results.Add($"Face {idx + 1}: Cannot create roof");
                                    failedCount++;
                                    continue;
                                }

                                // Apply offset and slope
                                OffsetRoofByThickness(
                                    roof, roofType, faceNormalDirection,
                                    cleanPoints, level, slopeDegrees);

                                UpdateRoofSlope(roof, slopeDegrees);

                                results.Add($"Face {idx + 1}: OK (ID: {roof.Id.IntegerValue})");
                                successCount++;
                            }
                            catch (Exception ex)
                            {
                                results.Add($"Face {idx + 1}: Error - {ex.Message}");
                                failedCount++;
                            }
                        }

                        trans.Commit();

                        // Show results
                        string resultMsg = "BATCH ROOF CREATION COMPLETE\n\n";
                        resultMsg += $"SUCCESS: {successCount}\n";
                        resultMsg += $"FAILED: {failedCount}\n\n";
                        resultMsg += "Details:\n";

                        // Show first 20 results
                        int displayCount = Math.Min(results.Count, 20);
                        for (int i = 0; i < displayCount; i++)
                        {
                            resultMsg += results[i] + "\n";
                        }

                        if (results.Count > 20)
                        {
                            resultMsg += $"\n... and {results.Count - 20} more";
                        }

                        ShowMessage(resultMsg, "BATCH RESULTS");
                        return Result.Succeeded;
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        ShowError($"Error: {ex.Message}");
                        return Result.Failed;
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError($"Error: {ex.Message}");
                return Result.Failed;
            }
        }
    }
}