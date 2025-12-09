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
            WinForms.MessageBox.Show(message, title, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
        }

        private void ShowError(string message, string title = "Error")
        {
            WinForms.MessageBox.Show(message, title, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
        }

        // ============================================================================
        // THICKNESS EXTRACTION (4 METHODS)
        // ============================================================================

        private double GetRoofTypeThickness(RoofType roofType)
        {
            // Extract roof type thickness using 4 fallback methods
            try
            {
                // Method 1: Compound Structure
                try
                {
                    CompoundStructure compound = roofType.GetCompoundStructure();
                    if (compound != null)
                    {
                        // PYREVIT → C# CONVERSION: GetTotalThickness() → GetWidth()
                        double thickness = compound.GetWidth(); // CONVERSION APPLIED
                        if (thickness > 0)
                        {
                            return thickness;
                        }
                    }
                }
                catch
                {
                    // Continue to next method
                }

                // Method 2: Thickness Parameter
                try
                {
                    Parameter param = roofType.LookupParameter("Thickness");
                    if (param != null && param.HasValue && param.AsDouble() > 0)
                    {
                        return param.AsDouble();
                    }
                }
                catch
                {
                    // Continue to next method
                }

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
                        catch
                        {
                            continue;
                        }
                    }
                }
                catch
                {
                    // Continue to next method
                }

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
                                // PYREVIT → C# CONVERSION: GetLayerThickness(i) → GetLayerWidth(i)
                                total += compound.GetLayerWidth(i); // CONVERSION APPLIED
                            }
                            catch
                            {
                                continue;
                            }
                        }
                        if (total > 0)
                        {
                            return total;
                        }
                    }
                }
                catch
                {
                    // Return 0 if all methods fail
                }

                return 0;
            }
            catch
            {
                return 0;
            }
        }

        // ============================================================================
        // FACE ANALYSIS
        // ============================================================================

        private class FaceData
        {
            public Element Element { get; set; }
            public Face Face { get; set; }
            public Reference Reference { get; set; }
        }

        private FaceData PickSingleFace()
        {
            // Pick a single face from 3D view
            try
            {
                Reference reference = _uidoc.Selection.PickObject(ObjectType.Face, "Select sloped face");
                if (reference != null)
                {
                    Element element = _doc.GetElement(reference.ElementId);
                    GeometryObject geometryObject = element.GetGeometryObjectFromReference(reference);
                    if (geometryObject is Face face)
                    {
                        return new FaceData
                        {
                            Element = element,
                            Face = face,
                            Reference = reference
                        };
                    }
                }
                return null;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return null;
            }
            catch
            {
                return null;
            }
        }

        private bool IsFaceSloped(Face face)
        {
            // Check if face is sloped (not horizontal)
            try
            {
                BoundingBoxUV bbox = face.GetBoundingBox();
                if (bbox != null)
                {
                    UV centerUV = new UV((bbox.Min.U + bbox.Max.U) / 2.0, (bbox.Min.V + bbox.Max.V) / 2.0);
                    XYZ normal = face.ComputeNormal(centerUV);
                    double angleTolerance = Math.Cos(Math.PI / 36.0); // 5 degrees in radians
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
            // Determine if face is TOP (1), BOTTOM (-1), or UNKNOWN (0)
            try
            {
                BoundingBoxUV bbox = face.GetBoundingBox();
                if (bbox != null)
                {
                    UV centerUV = new UV((bbox.Min.U + bbox.Max.U) / 2.0, (bbox.Min.V + bbox.Max.V) / 2.0);
                    XYZ normal = face.ComputeNormal(centerUV);
                    if (normal.Z > 0.5)
                    {
                        return 1;
                    }
                    else if (normal.Z < -0.5)
                    {
                        return -1;
                    }
                }
                return 0;
            }
            catch
            {
                return 0;
            }
        }

        private double? GetSlopeFromFaceNormal(Face face)
        {
            // Calculate slope angle from face normal
            try
            {
                BoundingBoxUV bbox = face.GetBoundingBox();
                if (bbox != null)
                {
                    UV centerUV = new UV((bbox.Min.U + bbox.Max.U) / 2.0, (bbox.Min.V + bbox.Max.V) / 2.0);
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
        // POINT EXTRACTION & PROCESSING
        // ============================================================================

        private List<XYZ> ExtractAllFacePoints(Face face, Element element)
        {
            try
            {
                List<XYZ> allPoints = new List<XYZ>();

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

                return allPoints; // Không cần transform
            }
            catch
            {
                return null;
            }
        }

        private List<XYZ> CleanupDuplicatePoints(List<XYZ> points)
        {
            // Remove duplicate points and sort them
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
        // ROOF CREATION
        // ============================================================================

        private (Level level, bool success, string message) GetLevelFromElement(Element element)
        {
            // Get level from element (wall, floor, roof, etc.)
            try
            {
                // Check LevelId property
                if (element.LevelId != null && element.LevelId != ElementId.InvalidElementId)
                {
                    Level level = _doc.GetElement(element.LevelId) as Level;
                    if (level != null)
                    {
                        return (level, true, $"Using element level: {level.Name}");
                    }
                }

                // Check Level parameter
                Parameter levelParam = element.LookupParameter("Level");
                if (levelParam != null && levelParam.HasValue)
                {
                    ElementId levelId = levelParam.AsElementId();
                    if (levelId != null && levelId != ElementId.InvalidElementId)
                    {
                        Level level = _doc.GetElement(levelId) as Level;
                        if (level != null)
                        {
                            return (level, true, $"Using element level: {level.Name}");
                        }
                    }
                }

                return (null, false, "Cannot find level from element!");
            }
            catch
            {
                return (null, false, "Error getting element level!");
            }
        }

        private (RoofBase roof, bool success) CreateValidRoofFootprint(List<XYZ> points, Level level, RoofType roofType, double slopeDegrees)
        {
            // Create roof footprint and set slope
            try
            {
                List<XYZ> flattenedPoints = points.Select(pt => new XYZ(pt.X, pt.Y, level.Elevation)).ToList();

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

                ModelCurveArray modelCurveArray = new ModelCurveArray();

                try
                {
                    RoofBase roof = _doc.Create.NewFootPrintRoof(curveArray, level, roofType, out modelCurveArray);
                    if (roof == null)
                    {
                        return (null, false);
                    }

                    bool slopeSuccess = SetSlopeOnLowestEdges(roof, points, slopeDegrees, modelCurveArray);
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

        private bool SetSlopeOnLowestEdges(RoofBase roof, List<XYZ> originalPoints, double slopeDegrees, ModelCurveArray modelCurveArray)
        {
            // Set slope on lowest edges of roof
            try
            {
                double minZ = originalPoints.Min(pt => pt.Z);
                double tolerance = 0.001;
                List<XYZ> lowestPoints = originalPoints.Where(p => Math.Abs(p.Z - minZ) < tolerance).ToList();

                List<(ModelCurve curveElem, double length)> lowestEdges = new List<(ModelCurve, double)>();

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

                var bestEdge = lowestEdges.OrderByDescending(x => x.length).First();

                Parameter definesSlopeParam = bestEdge.curveElem.get_Parameter(BuiltInParameter.ROOF_CURVE_IS_SLOPE_DEFINING);
                if (definesSlopeParam != null && !definesSlopeParam.IsReadOnly)
                {
                    definesSlopeParam.Set(1);
                }
                else
                {
                    return false;
                }

                Parameter slopeParam = bestEdge.curveElem.get_Parameter(BuiltInParameter.ROOF_SLOPE);
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
        // OFFSET CALCULATION & APPLICATION
        // ============================================================================

        private (bool success, double finalOffset, string offsetType, double thicknessMm)
            OffsetRoofByThickness(RoofBase roof, RoofType roofType, int faceNormalDirection,
                List<XYZ> originalPoints, Level level, double slopeDegrees)
        {
            // Calculate and apply offset based on thickness, slope, and elevation difference
            try
            {
                double thickness = GetRoofTypeThickness(roofType);

                double faceMinZ = originalPoints.Min(pt => pt.Z);
                double levelElevation = level.Elevation;
                double elevationDiff = faceMinZ - levelElevation;

                double thicknessOffset;
                string offsetType;

                if (faceNormalDirection == 1)
                {
                    thicknessOffset = 0.0;
                    offsetType = "TOP";
                }
                else if (faceNormalDirection == -1)
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
                else
                {
                    thicknessOffset = 0.0;
                    offsetType = "UNKNOWN";
                }

                double finalOffset = elevationDiff + thicknessOffset;

                try
                {
                    // Try to find and set offset parameter
                    Parameter baseOffsetParam = roof.LookupParameter("Base Offset");
                    if (baseOffsetParam == null)
                    {
                        baseOffsetParam = roof.LookupParameter("Base Level Offset");
                    }

                    if (baseOffsetParam != null && !baseOffsetParam.IsReadOnly)
                    {
                        baseOffsetParam.Set(finalOffset);
                        double thicknessMm = UnitUtils.ConvertFromInternalUnits(thickness, UnitTypeId.Millimeters);
                        return (true, finalOffset, offsetType, thicknessMm);
                    }

                    // Search through all parameters
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
                                double thicknessMm = UnitUtils.ConvertFromInternalUnits(thickness, UnitTypeId.Millimeters);
                                return (true, finalOffset, offsetType, thicknessMm);
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }

                    double thicknessMmFallback = UnitUtils.ConvertFromInternalUnits(thickness, UnitTypeId.Millimeters);
                    return (false, finalOffset, offsetType, thicknessMmFallback);
                }
                catch
                {
                    double thicknessMmFallback = UnitUtils.ConvertFromInternalUnits(thickness, UnitTypeId.Millimeters);
                    return (false, finalOffset, offsetType, thicknessMmFallback);
                }
            }
            catch
            {
                return (false, 0.0, "UNKNOWN", 0.0);
            }
        }

        private bool UpdateRoofSlope(RoofBase roof, double slopeDegrees)
        {
            // Update slope angle on roof
            try
            {
                // **SỬA LỖI: RoofBase không có GetSlopeDefiningLines() và SetSlopeAngle()**
                // **Cách sửa: Chỉ dùng slope parameter**

                // Phương pháp 1: Tìm Slope parameter
                Parameter slopeParam = roof.LookupParameter("Slope");
                if (slopeParam != null && !slopeParam.IsReadOnly)
                {
                    double slopeRatio = Math.Tan(slopeDegrees * (Math.PI / 180.0));
                    slopeParam.Set(slopeRatio);
                    return true;
                }

                // Phương pháp 2: Tìm các parameter khác có thể set slope
                foreach (Parameter param in roof.Parameters)
                {
                    try
                    {
                        string paramName = param.Definition?.Name ?? "";

                        // Kiểm tra các parameter name có thể chứa slope
                        if ((paramName.IndexOf("slope", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             paramName.IndexOf("pitch", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             paramName.IndexOf("angle", StringComparison.OrdinalIgnoreCase) >= 0) &&
                            param.StorageType == StorageType.Double &&
                            !param.IsReadOnly)
                        {
                            double slopeRatio = Math.Tan(slopeDegrees * (Math.PI / 180.0));
                            param.Set(slopeRatio);
                            return true;
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

                // Phương pháp 3: Thử set qua ModelCurves nếu có
                try
                {
                    // Lấy ModelCurveArray từ roof
                    ModelCurveArray modelCurves = null;

                    // Tìm các edge curves có thể set slope
                    foreach (ElementId elemId in roof.GetDependentElements(new ElementClassFilter(typeof(ModelCurve))))
                    {
                        ModelCurve modelCurve = _doc.GetElement(elemId) as ModelCurve;
                        if (modelCurve != null)
                        {
                            // Kiểm tra nếu curve này có slope parameter
                            Parameter definesSlopeParam = modelCurve.get_Parameter(BuiltInParameter.ROOF_CURVE_IS_SLOPE_DEFINING);
                            if (definesSlopeParam != null && !definesSlopeParam.IsReadOnly)
                            {
                                definesSlopeParam.Set(1);

                                Parameter curveSlopeParam = modelCurve.get_Parameter(BuiltInParameter.ROOF_SLOPE);
                                if (curveSlopeParam != null && !curveSlopeParam.IsReadOnly)
                                {
                                    double slopeRadians = slopeDegrees * (Math.PI / 180.0);
                                    curveSlopeParam.Set(slopeRadians);
                                    return true;
                                }
                            }
                        }
                    }
                }
                catch
                {
                    // Continue
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        // ============================================================================
        // ROOF TYPE SELECTOR
        // ============================================================================

        private List<(string name, RoofType roofType)> GetAllRoofTypes()
        {
            // Get all roof types in project
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
                            typeName = $"Roof Type {rt.Id.Value}"; // CONVERSION: IntegerValue → Value
                        }

                        roofTypeList.Add((typeName, rt));
                    }
                    catch
                    {
                        roofTypeList.Add(($"Roof Type {rt.Id.Value}", rt)); // CONVERSION: IntegerValue → Value
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
            // Show dialog to select roof type
            try
            {
                List<(string name, RoofType roofType)> roofTypes = GetAllRoofTypes();
                if (roofTypes == null || roofTypes.Count == 0)
                {
                    ShowError("No roof types found!");
                    return (null, null);
                }

                Dictionary<string, RoofType> typeDict = roofTypes.ToDictionary(x => x.name, x => x.roofType);
                List<string> typeNames = typeDict.Keys.OrderBy(x => x).ToList();

                using (WinForms.Form form = new WinForms.Form())
                {
                    form.Text = "Select Roof Type";
                    form.Width = 450;
                    form.Height = 380;
                    form.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
                    form.MaximizeBox = false;
                    form.MinimizeBox = false;

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
            catch
            {
                return (null, null);
            }
        }

        // ============================================================================
        // MAIN WORKFLOW
        // ============================================================================

        private Result RunMain()
        {
            try
            {
                ShowMessage(
                    "INTEGRATED ROOF TOOL\n\n" +
                    "Process:\n" +
                    "1. Select sloped face\n" +
                    "2. Detect face orientation (TOP/BOTTOM)\n" +
                    "3. Create roof footprint\n" +
                    "4. Extract roof type thickness\n" +
                    "5. Calculate slope\n" +
                    "6. Set slope on edges\n" +
                    "7. Apply smart offset\n\n" +
                    "Start now...",
                    "Roof Tool"
                );

                FaceData faceData = PickSingleFace();
                if (faceData == null)
                {
                    return Result.Cancelled;
                }

                if (!IsFaceSloped(faceData.Face))
                {
                    ShowError("Selected face is horizontal!\nPlease select a sloped face.");
                    return Result.Failed;
                }

                int faceNormalDirection = GetFaceNormalDirection(faceData.Face);

                List<XYZ> originalPoints = ExtractAllFacePoints(faceData.Face, faceData.Element);
                if (originalPoints == null || originalPoints.Count == 0)
                {
                    ShowError("Cannot extract points from face!");
                    return Result.Failed;
                }

                List<XYZ> cleanPoints = CleanupDuplicatePoints(originalPoints);
                if (cleanPoints.Count < 3)
                {
                    ShowError("Insufficient points!");
                    return Result.Failed;
                }

                double? slopeDegreesNullable = GetSlopeFromFaceNormal(faceData.Face);
                if (!slopeDegreesNullable.HasValue)
                {
                    ShowError("Cannot calculate slope!");
                    return Result.Failed;
                }
                double slopeDegrees = slopeDegreesNullable.Value;

                var (roofType, roofTypeName) = ShowRoofTypeSelector();
                if (roofType == null)
                {
                    ShowError("No roof type selected!");
                    return Result.Cancelled;
                }

                using (Transaction trans = new Transaction(_doc, "Create Roof"))
                {
                    trans.Start();

                    try
                    {
                        var (level, levelOk, levelMsg) = GetLevelFromElement(faceData.Element);
                        if (!levelOk || level == null)
                        {
                            trans.RollBack();
                            ShowError(levelMsg, "Error!");
                            return Result.Failed;
                        }

                        var (roof, slopeSuccess) = CreateValidRoofFootprint(cleanPoints, level, roofType, slopeDegrees);
                        if (roof == null)
                        {
                            trans.RollBack();
                            ShowError("Cannot create roof!");
                            return Result.Failed;
                        }

                        var (offsetSuccess, finalOffset, offsetType, thicknessMm) =
                            OffsetRoofByThickness(roof, roofType, faceNormalDirection, cleanPoints, level, slopeDegrees);

                        bool slopeUpdateSuccess = UpdateRoofSlope(roof, slopeDegrees);

                        trans.Commit();

                        string resultMessage = "SUCCESS - Roof created!\n\n" +
                                              "ROOF INFO:\n" +
                                              $"• ID: {roof.Id.Value}\n" + // CONVERSION: IntegerValue → Value
                                              $"• Type: {roofTypeName ?? "Unknown"}\n" +
                                              $"• Level: {level.Name}\n\n" +
                                              "FACE INFO:\n" +
                                              $"• Orientation: {(faceNormalDirection == 1 ? "TOP" : faceNormalDirection == -1 ? "BOTTOM" : "UNKNOWN")}\n" +
                                              $"• Thickness: {thicknessMm:F2}mm\n" +
                                              $"• Offset Type: {offsetType}\n\n" +
                                              "SLOPE INFO:\n" +
                                              $"• Angle: {slopeDegrees:F2}°\n" +
                                              $"• Set: {(slopeSuccess ? "OK" : "Partial")}\n" +
                                              $"• Updated: {(slopeUpdateSuccess ? "OK" : "Partial")}\n" +
                                              $"• Offset: {(offsetSuccess ? "OK" : "Partial")}";

                        ShowMessage(resultMessage, "SUCCESS");
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