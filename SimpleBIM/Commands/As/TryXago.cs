using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Windows.Forms;
using System.Drawing;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;

namespace SimpleBIM.Commands.As
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class TryXago : IExternalCommand
    {
        // CLASS NAME: Giữ nguyên ý nghĩa từ Python "INTEGRATED BEAM SYSTEM TOOL", thêm suffix "Command"
        // Python: "INTEGRATED BEAM SYSTEM TOOL" → C#: "IntegratedBeamSystemToolCommand"

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

        [Conditional("DEBUG")]
        private void DebugLog(string message)
        {
            Debug.WriteLine($"[DEBUG] {message}");
        }

        // ============================================================================
        // DATA STRUCTURES
        // ============================================================================
        private class FaceData
        {
            public Element Element { get; set; }
            public Face Face { get; set; }
            public Reference Reference { get; set; }
        }

        private class EdgeData
        {
            public int Index { get; set; }
            public XYZ Start { get; set; }
            public XYZ End { get; set; }
            public double Length { get; set; }
            public XYZ Direction { get; set; }
            public double SlopeAngle { get; set; }
        }

        // ============================================================================
        // FACE OPERATIONS
        // ============================================================================
        private FaceData PickSingleFace(UIDocument uidoc, Document doc)
        {
            try
            {
                Reference reference = uidoc.Selection.PickObject(
                    ObjectType.Face,
                    "Select sloped face");

                if (reference != null)
                {
                    Element element = doc.GetElement(reference.ElementId);
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
            try
            {
                BoundingBoxUV bbox = face.GetBoundingBox();
                if (bbox != null)
                {
                    UV centerUV = new UV(
                        (bbox.Min.U + bbox.Max.U) / 2.0,
                        (bbox.Min.V + bbox.Max.V) / 2.0);

                    XYZ normal = face.ComputeNormal(centerUV);
                    double angleTolerance = Math.Cos(Math.PI * 5.0 / 180.0);
                    return Math.Abs(normal.Z) <= angleTolerance;
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        private List<XYZ> ExtractFacePoints(Face face, Element element)
        {
            List<XYZ> allPoints = new List<XYZ>();

            try
            {
                Options options = new Options();
                options.ComputeReferences = true;
                GeometryElement geometryElement = element.get_Geometry(options);

                if (geometryElement != null)
                {
                    foreach (GeometryObject geomObj in geometryElement)
                    {
                        if (geomObj is Solid solid && solid.Faces.Size > 0)
                        {
                            foreach (Face faceObj in solid.Faces)
                            {
                                // Compare faces by area
                                if (Math.Abs(faceObj.Area - face.Area) < 0.001)
                                {
                                    // Extract edges from this face
                                    foreach (EdgeArray edgeLoop in faceObj.EdgeLoops)
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
                                    break;
                                }
                            }
                        }
                    }
                }

                if (allPoints.Count == 0)
                {
                    // Fallback: try to get points directly from the face
                    try
                    {
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
                    }
                    catch { }
                }

                // Apply transform if available
                List<XYZ> originalPoints = new List<XYZ>();
                try
                {
                    // NOTE: Geometry from GetGeometryObjectFromReference() is ALREADY in world coordinates
                    // NO NEED to get element transform here
                    originalPoints = allPoints;
                    DebugLog("Using points directly - already in world coordinates");
                }
                catch
                {
                    originalPoints = allPoints;
                }

                return originalPoints;
            }
            catch (Exception e)
            {
                DebugLog($"Error extracting face points: {e.Message}");
                return null;
            }
        }

        private List<XYZ> CleanupDuplicatePoints(List<XYZ> points)
        {
            if (points == null || points.Count == 0)
                return new List<XYZ>();

            List<XYZ> uniquePoints = new List<XYZ>();
            double tolerance = 0.01;

            foreach (XYZ pt in points)
            {
                bool isDuplicate = false;
                foreach (XYZ existing in uniquePoints)
                {
                    double dist3D = Math.Sqrt(
                        Math.Pow(pt.X - existing.X, 2) +
                        Math.Pow(pt.Y - existing.Y, 2) +
                        Math.Pow(pt.Z - existing.Z, 2));

                    if (dist3D < tolerance)
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

            return uniquePoints;
        }

        private List<XYZ> SortPointsCCW(List<XYZ> points)
        {
            if (points.Count < 3)
                return points;

            double centerX = points.Sum(pt => pt.X) / points.Count;
            double centerY = points.Sum(pt => pt.Y) / points.Count;

            return points.OrderBy(pt => Math.Atan2(pt.Y - centerY, pt.X - centerX)).ToList();
        }

        private (Level level, bool ok) GetLevelFromElement(Element element, Document doc)
        {
            try
            {
                // Try: LevelId property
                if (element.LevelId != null && element.LevelId != ElementId.InvalidElementId)
                {
                    Level level = doc.GetElement(element.LevelId) as Level;
                    if (level != null)
                        return (level, true);
                }

                // Try: Level parameters
                string[] levelParams = { "Level", "Base Level", "Reference Level", "Constraints Level" };
                foreach (string paramName in levelParams)
                {
                    Parameter levelParam = element.LookupParameter(paramName);
                    if (levelParam != null && levelParam.AsElementId() != null &&
                        levelParam.AsElementId() != ElementId.InvalidElementId)
                    {
                        Level level = doc.GetElement(levelParam.AsElementId()) as Level;
                        if (level != null)
                            return (level, true);
                    }
                }

                // Try: Get closest level from Z coordinate
                BoundingBoxXYZ bbox = element.get_BoundingBox(null);
                if (bbox != null)
                {
                    double avgZ = (bbox.Min.Z + bbox.Max.Z) / 2.0;
                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                    IList<Element> levels = collector.OfClass(typeof(Level)).ToElements();

                    Level closestLevel = null;
                    double minDiff = double.MaxValue;

                    foreach (Element elem in levels)
                    {
                        Level lvl = elem as Level;
                        if (lvl != null)
                        {
                            double diff = Math.Abs(lvl.Elevation - avgZ);
                            if (diff < minDiff)
                            {
                                minDiff = diff;
                                closestLevel = lvl;
                            }
                        }
                    }

                    if (closestLevel != null)
                        return (closestLevel, true);
                }

                return (null, false);
            }
            catch
            {
                return (null, false);
            }
        }

        // ============================================================================
        // EDGE ANALYSIS
        // ============================================================================
        private EdgeData FindBaseEdge(List<XYZ> sortedPoints)
        {
            try
            {
                EdgeData bestEdge = null;
                double bestScore = double.NegativeInfinity;

                for (int i = 0; i < sortedPoints.Count; i++)
                {
                    XYZ p1 = sortedPoints[i];
                    XYZ p2 = sortedPoints[(i + 1) % sortedPoints.Count];

                    double dx = p2.X - p1.X;
                    double dy = p2.Y - p1.Y;
                    double dz = p2.Z - p1.Z;

                    double xyLength = Math.Sqrt(dx * dx + dy * dy);
                    if (xyLength < 0.001)
                        continue;

                    double horizontality = 1.0 / (1.0 + Math.Abs(dz));
                    double score = xyLength * horizontality;

                    if (score > bestScore)
                    {
                        bestScore = score;
                        XYZ dirXyz = new XYZ(dx, dy, dz);
                        if (dirXyz.GetLength() > 0)
                            dirXyz = dirXyz.Normalize();

                        bestEdge = new EdgeData
                        {
                            Index = i,
                            Start = p1,
                            End = p2,
                            Length = xyLength,
                            Direction = dirXyz
                        };
                    }
                }

                return bestEdge;
            }
            catch (Exception e)
            {
                DebugLog($"Error finding base edge: {e.Message}");
                return null;
            }
        }

        private EdgeData FindSlopeEdge(List<XYZ> sortedPoints, XYZ baseDirection)
        {
            try
            {
                EdgeData bestEdge = null;
                double bestSlope = double.NegativeInfinity;

                for (int i = 0; i < sortedPoints.Count; i++)
                {
                    XYZ p1 = sortedPoints[i];
                    XYZ p2 = sortedPoints[(i + 1) % sortedPoints.Count];

                    double dx = p2.X - p1.X;
                    double dy = p2.Y - p1.Y;
                    double dz = p2.Z - p1.Z;

                    double length3D = Math.Sqrt(dx * dx + dy * dy + dz * dz);
                    if (length3D < 0.001)
                        continue;

                    XYZ edgeDir = new XYZ(dx, dy, dz);
                    if (edgeDir.GetLength() > 0)
                        edgeDir = edgeDir.Normalize();

                    double dot = Math.Abs(edgeDir.DotProduct(baseDirection));
                    if (dot > 0.7)
                        continue;

                    double horizontalLength = Math.Sqrt(dx * dx + dy * dy);
                    if (horizontalLength > 0.001)
                    {
                        double slope = Math.Abs(dz) / horizontalLength;

                        if (slope > bestSlope)
                        {
                            bestSlope = slope;
                            double slopeDeg = Math.Atan2(Math.Abs(dz), horizontalLength) * (180.0 / Math.PI);

                            bestEdge = new EdgeData
                            {
                                Index = i,
                                Start = p1,
                                End = p2,
                                Length = length3D,
                                Direction = edgeDir,
                                SlopeAngle = slopeDeg
                            };
                        }
                    }
                }

                return bestEdge;
            }
            catch (Exception e)
            {
                DebugLog($"Error finding slope edge: {e.Message}");
                return null;
            }
        }

        // ============================================================================
        // BEAM TYPE SELECTOR
        // ============================================================================
        private List<(string name, FamilySymbol beamType)> GetAllBeamTypes(Document doc)
        {
            var beamTypeList = new List<(string, FamilySymbol)>();

            try
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                IList<Element> beamTypes = collector
                    .OfClass(typeof(FamilySymbol))
                    .OfCategory(BuiltInCategory.OST_StructuralFraming)
                    .ToElements();

                foreach (Element element in beamTypes)
                {
                    FamilySymbol beamType = element as FamilySymbol;
                    if (beamType != null)
                    {
                        try
                        {
                            Parameter nameParam = beamType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME);
                            string typeName = nameParam != null && nameParam.HasValue ?
                                nameParam.AsString() :
                                $"Beam Type {beamType.Id.Value}";

                            beamTypeList.Add((typeName, beamType));
                        }
                        catch
                        {
                            beamTypeList.Add(($"Beam Type {beamType.Id.Value}", beamType));
                        }
                    }
                }

                return beamTypeList;
            }
            catch
            {
                return beamTypeList;
            }
        }

        private (FamilySymbol beamType, string name) ShowBeamTypeSelector(int layerNumber, Document doc)
        {
            try
            {
                var beamTypes = GetAllBeamTypes(doc);
                if (beamTypes.Count == 0)
                {
                    ShowError("No beam types found!");
                    return (null, null);
                }

                var typeDict = beamTypes.ToDictionary(item => item.name, item => item.beamType);
                List<string> typeNames = typeDict.Keys.OrderBy(k => k).ToList();

                using (var form = new WinForms.Form())
                {
                    form.Text = $"Select Beam Type - Layer {layerNumber}";
                    form.Width = 450;
                    form.Height = 380;
                    form.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
                    form.MaximizeBox = false;
                    form.MinimizeBox = false;

                    var label = new WinForms.Label();
                    label.Text = $"Select beam type for Layer {layerNumber}:";
                    label.Location = new Drawing.Point(15, 15);
                    label.Width = 400;
                    label.Height = 25;
                    form.Controls.Add(label);

                    var listbox = new WinForms.ListBox();
                    listbox.Location = new Drawing.Point(15, 45);
                    listbox.Width = 400;
                    listbox.Height = 230;

                    int defaultIdx = 0;
                    for (int i = 0; i < typeNames.Count; i++)
                    {
                        listbox.Items.Add(typeNames[i]);
                        if (typeNames[i].Contains("W") || typeNames[i].Contains("IPE"))
                        {
                            defaultIdx = i;
                        }
                    }

                    listbox.SelectedIndex = Math.Max(0, defaultIdx);
                    form.Controls.Add(listbox);

                    var resultBeamType = new FamilySymbol[1];
                    var resultName = new string[1];

                    void OnOk(object sender, EventArgs e)
                    {
                        if (listbox.SelectedIndex >= 0)
                        {
                            string name = typeNames[listbox.SelectedIndex];
                            resultBeamType[0] = typeDict[name];
                            resultName[0] = name;
                        }
                        form.DialogResult = WinForms.DialogResult.OK;
                    }

                    void OnCancel(object sender, EventArgs e)
                    {
                        form.DialogResult = WinForms.DialogResult.Cancel;
                    }

                    var btnOk = new WinForms.Button();
                    btnOk.Text = "OK";
                    btnOk.Location = new Drawing.Point(155, 290);
                    btnOk.Width = 90;
                    btnOk.Height = 35;
                    btnOk.Click += OnOk;
                    form.Controls.Add(btnOk);
                    form.AcceptButton = btnOk;

                    var btnCancel = new WinForms.Button();
                    btnCancel.Text = "Cancel";
                    btnCancel.Location = new Drawing.Point(260, 290);
                    btnCancel.Width = 90;
                    btnCancel.Height = 35;
                    btnCancel.Click += OnCancel;
                    form.Controls.Add(btnCancel);
                    form.CancelButton = btnCancel;

                    if (form.ShowDialog() == WinForms.DialogResult.OK)
                    {
                        return (resultBeamType[0], resultName[0]);
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

        private double GetBeamHeight(FamilySymbol beamType)
        {
            try
            {
                if (beamType == null)
                    return 0.5;

                Parameter heightParam = beamType.LookupParameter("Height");
                if (heightParam != null && heightParam.HasValue && heightParam.AsDouble() > 0)
                {
                    return heightParam.AsDouble();
                }

                string[] paramNames = { "h", "H", "d", "D", "Depth" };
                foreach (string paramName in paramNames)
                {
                    Parameter param = beamType.LookupParameter(paramName);
                    if (param != null && param.HasValue && param.AsDouble() > 0)
                    {
                        return param.AsDouble();
                    }
                }

                return 0.5;
            }
            catch
            {
                return 0.5;
            }
        }

        // ============================================================================
        // SPACING DIALOG
        // ============================================================================
        private (double? spacing1, double? spacing2, double? spacing3) ShowSpacingDialog()
        {
            try
            {
                using (var form = new WinForms.Form())
                {
                    form.Text = "Enter Beam Spacing";
                    form.Width = 400;
                    form.Height = 300;
                    form.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
                    form.MaximizeBox = false;
                    form.MinimizeBox = false;

                    var label1 = new WinForms.Label();
                    label1.Text = "Spacing Layer 1 (Horizontal - Bottom) [m]:";
                    label1.Location = new Drawing.Point(15, 15);
                    label1.Width = 350;
                    label1.Height = 20;
                    form.Controls.Add(label1);

                    var text1 = new WinForms.TextBox();
                    text1.Text = "2.0";
                    text1.Location = new Drawing.Point(15, 40);
                    text1.Width = 350;
                    text1.Height = 25;
                    form.Controls.Add(text1);

                    var label2 = new WinForms.Label();
                    label2.Text = "Spacing Layer 2 (Slope - Middle) [m]:";
                    label2.Location = new Drawing.Point(15, 75);
                    label2.Width = 350;
                    label2.Height = 20;
                    form.Controls.Add(label2);

                    var text2 = new WinForms.TextBox();
                    text2.Text = "2.0";
                    text2.Location = new Drawing.Point(15, 100);
                    text2.Width = 350;
                    text2.Height = 25;
                    form.Controls.Add(text2);

                    var label3 = new WinForms.Label();
                    label3.Text = "Spacing Layer 3 (Horizontal - Top) [m]:";
                    label3.Location = new Drawing.Point(15, 135);
                    label3.Width = 350;
                    label3.Height = 20;
                    form.Controls.Add(label3);

                    var text3 = new WinForms.TextBox();
                    text3.Text = "2.0";
                    text3.Location = new Drawing.Point(15, 160);
                    text3.Width = 350;
                    text3.Height = 25;
                    form.Controls.Add(text3);

                    var result = new double?[3];

                    void OnOk(object sender, EventArgs e)
                    {
                        try
                        {
                            double s1 = double.Parse(text1.Text);
                            double s2 = double.Parse(text2.Text);
                            double s3 = double.Parse(text3.Text);

                            if (s1 <= 0 || s2 <= 0 || s3 <= 0)
                            {
                                ShowError("Spacing must be positive!");
                                return;
                            }

                            // Convert meters to feet (Revit internal units)
                            result[0] = s1 * 3.28084;
                            result[1] = s2 * 3.28084;
                            result[2] = s3 * 3.28084;
                            form.DialogResult = WinForms.DialogResult.OK;
                        }
                        catch
                        {
                            ShowError("Invalid spacing value!");
                        }
                    }

                    void OnCancel(object sender, EventArgs e)
                    {
                        form.DialogResult = WinForms.DialogResult.Cancel;
                    }

                    var btnOk = new WinForms.Button();
                    btnOk.Text = "OK";
                    btnOk.Location = new Drawing.Point(120, 210);
                    btnOk.Width = 90;
                    btnOk.Height = 35;
                    btnOk.Click += OnOk;
                    form.Controls.Add(btnOk);
                    form.AcceptButton = btnOk;

                    var btnCancel = new WinForms.Button();
                    btnCancel.Text = "Cancel";
                    btnCancel.Location = new Drawing.Point(225, 210);
                    btnCancel.Width = 90;
                    btnCancel.Height = 35;
                    btnCancel.Click += OnCancel;
                    form.Controls.Add(btnCancel);
                    form.CancelButton = btnCancel;

                    if (form.ShowDialog() == WinForms.DialogResult.OK)
                    {
                        return (result[0], result[1], result[2]);
                    }
                    else
                    {
                        return (null, null, null);
                    }
                }
            }
            catch
            {
                return (null, null, null);
            }
        }

        // ============================================================================
        // SLOPE CALCULATION
        // ============================================================================
        private double CalculateZOnSlope(XYZ basePoint, XYZ currentPoint, EdgeData slopeEdge, double slopeAngleRad, double baseZ)
        {
            try
            {
                // Get slope edge start and end
                XYZ slopeStart = slopeEdge.Start;
                XYZ slopeEnd = slopeEdge.End;

                // 2D projection (ignore Z)
                XYZ slopeStart2D = new XYZ(slopeStart.X, slopeStart.Y, 0);
                XYZ slopeEnd2D = new XYZ(slopeEnd.X, slopeEnd.Y, 0);
                XYZ currentPoint2D = new XYZ(currentPoint.X, currentPoint.Y, 0);

                // Slope direction (horizontal)
                XYZ slopeDir2D = slopeEnd2D - slopeStart2D;
                double slopeLen2D = slopeDir2D.GetLength();

                if (slopeLen2D < 0.001)
                    return baseZ;

                XYZ slopeDirNormalized = slopeDir2D.Normalize();

                // Vector from slope start to current point
                XYZ vecToPoint = currentPoint2D - slopeStart2D;

                // Project onto slope direction
                double distanceAlongSlope = vecToPoint.DotProduct(slopeDirNormalized);

                // Clamp to slope edge length
                if (distanceAlongSlope < 0)
                    distanceAlongSlope = 0;
                else if (distanceAlongSlope > slopeLen2D)
                    distanceAlongSlope = slopeLen2D;

                // Calculate Z change
                double zChange = distanceAlongSlope * Math.Tan(slopeAngleRad);

                // Get actual Z at slope start and end
                double zStart = slopeStart.Z;
                double zEnd = slopeEnd.Z;

                // Interpolate Z based on position along slope
                double t = slopeLen2D > 0 ? distanceAlongSlope / slopeLen2D : 0;
                double zFromSlope = zStart + (zEnd - zStart) * t;

                // Use baseZ + zChange approach
                double newZ = baseZ + zChange;

                DebugLog($"    Z calc: dist={distanceAlongSlope:F2}, z_change={zChange:F2}, new_z={newZ:F2}, from_slope={zFromSlope:F2}");

                return newZ;
            }
            catch (Exception e)
            {
                DebugLog($"Error calculating Z on slope: {e.Message}");
                return baseZ;
            }
        }

        // ============================================================================
        // BEAM CROSS-SECTION ROTATION
        // ============================================================================
        private bool SetBeamCrossSectionRotation(FamilyInstance beam, double slopeAngleDeg)
        {
            try
            {
                if (beam == null)
                    return false;

                // Reverse the angle (negative instead of positive)
                double rotationAngle = -slopeAngleDeg;

                // Try to find and set Rotation parameter
                Parameter rotationParam = beam.LookupParameter("Rotation");
                if (rotationParam != null && !rotationParam.IsReadOnly)
                {
                    // Convert angle to radians for Revit
                    double angleRad = rotationAngle * (Math.PI / 180.0);
                    rotationParam.Set(angleRad);
                    DebugLog($"    Cross-section rotated: {rotationAngle:F1}°");
                    return true;
                }
                else
                {
                    // Try alternative parameter names
                    string[] paramNames = { "Cross-Section Rotation", "Beam Rotation", "Roll" };
                    foreach (string paramName in paramNames)
                    {
                        Parameter param = beam.LookupParameter(paramName);
                        if (param != null && !param.IsReadOnly)
                        {
                            double angleRad = rotationAngle * (Math.PI / 180.0);
                            param.Set(angleRad);
                            DebugLog($"    Cross-section rotated ({paramName}): {rotationAngle:F1}°");
                            return true;
                        }
                    }

                    DebugLog("    Warning: Rotation parameter not found");
                    return false;
                }
            }
            catch (Exception e)
            {
                DebugLog($"    Error setting rotation: {e.Message}");
                return false;
            }
        }

        // ============================================================================
        // BEAM CREATION
        // ============================================================================
        private FamilyInstance CreateBeamDirect(XYZ startPoint, XYZ endPoint, FamilySymbol beamType, Level level, Document doc)
        {
            try
            {
                if (beamType == null || level == null)
                    return null;

                if (!beamType.IsActive)
                {
                    beamType.Activate();
                    try
                    {
                        doc.Regenerate();
                    }
                    catch { }
                }

                if (startPoint.DistanceTo(endPoint) < 1e-3)
                    return null;

                Line line = Line.CreateBound(startPoint, endPoint);
                if (line == null)
                    return null;

                try
                {
                    FamilyInstance beam = doc.Create.NewFamilyInstance(line, beamType, level, StructuralType.Beam);

                    if (beam != null)
                    {
                        return beam;
                    }
                    else
                    {
                        return null;
                    }
                }
                catch
                {
                    try
                    {
                        FamilyInstance beam = doc.Create.NewFamilyInstance(line, beamType, level, StructuralType.Beam);
                        //                                                                       
                        if (beam != null)
                        {
                            return beam;
                        }
                        else
                        {
                            return null;
                        }
                    }
                    catch
                    {
                        return null;
                    }
                }
            }
            catch (Exception e)
            {
                DebugLog($"Error creating beam: {e.Message}");
                return null;
            }
        }

        private XYZ LineIntersection2DDirect(XYZ line1Start, XYZ line1End, XYZ line2Start, XYZ line2End)
        {
            try
            {
                XYZ p1 = new XYZ(line1Start.X, line1Start.Y, 0);
                XYZ p2 = new XYZ(line1End.X, line1End.Y, 0);
                XYZ p3 = new XYZ(line2Start.X, line2Start.Y, 0);
                XYZ p4 = new XYZ(line2End.X, line2End.Y, 0);

                XYZ d1 = p2 - p1;
                XYZ d2 = p4 - p3;

                double det = d1.X * d2.Y - d1.Y * d2.X;

                if (Math.Abs(det) < 1e-10)
                    return null;

                double t = ((p3.X - p1.X) * d2.Y - (p3.Y - p1.Y) * d2.X) / det;
                double u = -((p1.X - p3.X) * d1.Y - (p1.Y - p3.Y) * d1.X) / det;

                if (0 <= t && t <= 1 && 0 <= u && u <= 1)
                {
                    double interX = p1.X + t * d1.X;
                    double interY = p1.Y + t * d1.Y;
                    double z1 = line1Start.Z + t * (line1End.Z - line1Start.Z);
                    double z2 = line2Start.Z + u * (line2End.Z - line2Start.Z);
                    double interZ = (z1 + z2) / 2.0;
                    return new XYZ(interX, interY, interZ);
                }

                return null;
            }
            catch (Exception e)
            {
                DebugLog($"Intersection error: {e.Message}");
                return null;
            }
        }

        private int CreateHorizontalLayerBeamsWithSlope(
            List<XYZ> boundaryPoints, EdgeData baseEdge, EdgeData slopeEdge,
            double spacing, double zHeight, FamilySymbol beamType, Level level,
            double slopeAngleDeg, Document doc)
        {
            int beamCount = 0;

            try
            {
                DebugLog("=== Creating Horizontal Layer Beams (WITH SLOPE) ===");
                DebugLog($"Spacing: {spacing:F2}, Base Z: {zHeight:F2}, Slope: {slopeAngleDeg:F1}°");

                double slopeAngleRad = slopeAngleDeg * (Math.PI / 180.0);

                XYZ startPt = baseEdge.Start;
                XYZ endPt = baseEdge.End;

                XYZ edgeVector = endPt - startPt;
                XYZ edgeDirection = edgeVector.Normalize();
                XYZ perpDirection = new XYZ(-edgeDirection.Y, edgeDirection.X, 0).Normalize();

                double minDist = double.MaxValue;
                double maxDist = double.MinValue;

                foreach (XYZ pt in boundaryPoints)
                {
                    XYZ vecToPt = pt - startPt;
                    double dist = vecToPt.DotProduct(perpDirection);
                    minDist = Math.Min(minDist, dist);
                    maxDist = Math.Max(maxDist, dist);
                }

                int numBeams = (int)((maxDist - minDist) / spacing);
                DebugLog($"Attempting to create {numBeams} beams");

                for (int i = 0; i < numBeams; i++)
                {
                    try
                    {
                        double offset = minDist + (i + 0.5) * spacing;
                        XYZ beamMid = startPt + perpDirection * offset;
                        XYZ beamStart = beamMid - edgeDirection * 100.0;
                        XYZ beamEnd = beamMid + edgeDirection * 100.0;

                        List<XYZ> intersections = new List<XYZ>();
                        for (int j = 0; j < boundaryPoints.Count; j++)
                        {
                            XYZ bp1 = boundaryPoints[j];
                            XYZ bp2 = boundaryPoints[(j + 1) % boundaryPoints.Count];
                            XYZ intersection = LineIntersection2DDirect(beamStart, beamEnd, bp1, bp2);
                            if (intersection != null)
                            {
                                intersections.Add(intersection);
                            }
                        }

                        List<XYZ> uniqueIntersections = new List<XYZ>();
                        foreach (XYZ inter in intersections)
                        {
                            bool isDuplicate = false;
                            foreach (XYZ existing in uniqueIntersections)
                            {
                                double dist2D = Math.Sqrt(
                                    Math.Pow(inter.X - existing.X, 2) +
                                    Math.Pow(inter.Y - existing.Y, 2));

                                if (dist2D < 0.01)
                                {
                                    isDuplicate = true;
                                    break;
                                }
                            }

                            if (!isDuplicate)
                            {
                                uniqueIntersections.Add(inter);
                            }
                        }

                        if (uniqueIntersections.Count >= 2)
                        {
                            XYZ pt1, pt2;

                            if (uniqueIntersections.Count > 2)
                            {
                                double maxDistance = 0;
                                (XYZ, XYZ) bestPair = (uniqueIntersections[0], uniqueIntersections[1]);

                                for (int k = 0; k < uniqueIntersections.Count; k++)
                                {
                                    for (int l = k + 1; l < uniqueIntersections.Count; l++)
                                    {
                                        double dist = uniqueIntersections[k].DistanceTo(uniqueIntersections[l]);
                                        if (dist > maxDistance)
                                        {
                                            maxDistance = dist;
                                            bestPair = (uniqueIntersections[k], uniqueIntersections[l]);
                                        }
                                    }
                                }

                                pt1 = bestPair.Item1;
                                pt2 = bestPair.Item2;
                            }
                            else
                            {
                                pt1 = uniqueIntersections[0];
                                pt2 = uniqueIntersections[1];
                            }

                            // SLOPE CALCULATION - Calculate Z for both points based on slope
                            double z1 = CalculateZOnSlope(slopeEdge.Start, pt1, slopeEdge, slopeAngleRad, zHeight);
                            double z2 = CalculateZOnSlope(slopeEdge.Start, pt2, slopeEdge, slopeAngleRad, zHeight);

                            XYZ pt1WithSlope = new XYZ(pt1.X, pt1.Y, z1);
                            XYZ pt2WithSlope = new XYZ(pt2.X, pt2.Y, z2);

                            FamilyInstance beam = CreateBeamDirect(pt1WithSlope, pt2WithSlope, beamType, level, doc);
                            if (beam != null)
                            {
                                beamCount++;
                                // SET cross-section rotation to align with slope
                                SetBeamCrossSectionRotation(beam, slopeAngleDeg);
                                DebugLog($"  Beam {i} created with rotation: Z1={z1:F3}, Z2={z2:F3}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLog($"  Beam {i} error: {ex.Message}");
                    }
                }

                DebugLog($"Horizontal layer complete: {beamCount} beams");
                return beamCount;
            }
            catch (Exception ex)
            {
                DebugLog($"ERROR in horizontal layer: {ex.Message}");
                return 0;
            }
        }

        private int CreateSlopeLayerBeamsWithSlope(
            List<XYZ> boundaryPoints, EdgeData slopeEdge, double spacing,
            double baseZ, FamilySymbol beamType, Level level, double slopeAngleDeg, Document doc)
        {
            int beamCount = 0;

            try
            {
                DebugLog("=== Creating Slope Layer Beams (WITH SLOPE) ===");
                DebugLog($"Spacing: {spacing:F2}, Base Z: {baseZ:F2}, Slope: {slopeAngleDeg:F1}°");

                double slopeAngleRad = slopeAngleDeg * (Math.PI / 180.0);

                XYZ startPt = slopeEdge.Start;
                XYZ endPt = slopeEdge.End;

                XYZ edgeVector = endPt - startPt;
                XYZ edgeDirection = edgeVector.Normalize();
                XYZ perpToSlope = new XYZ(-edgeDirection.Y, edgeDirection.X, 0).Normalize();

                double minDist = double.MaxValue;
                double maxDist = double.MinValue;

                foreach (XYZ pt in boundaryPoints)
                {
                    XYZ vecToPt = pt - startPt;
                    double dist = vecToPt.DotProduct(perpToSlope);
                    minDist = Math.Min(minDist, dist);
                    maxDist = Math.Max(maxDist, dist);
                }

                int numBeams = (int)((maxDist - minDist) / spacing);
                DebugLog($"Attempting to create {numBeams} slope beams");

                for (int i = 0; i < numBeams; i++)
                {
                    try
                    {
                        double offset = minDist + (i + 0.5) * spacing;
                        XYZ beamStartPoint = startPt + perpToSlope * offset;
                        XYZ beamLineStart = beamStartPoint - edgeDirection * 100.0;
                        XYZ beamLineEnd = beamStartPoint + edgeDirection * 100.0;

                        List<XYZ> intersections = new List<XYZ>();
                        for (int j = 0; j < boundaryPoints.Count; j++)
                        {
                            XYZ bp1 = boundaryPoints[j];
                            XYZ bp2 = boundaryPoints[(j + 1) % boundaryPoints.Count];
                            XYZ intersection = LineIntersection2DDirect(beamLineStart, beamLineEnd, bp1, bp2);
                            if (intersection != null)
                            {
                                intersections.Add(intersection);
                            }
                        }

                        List<XYZ> uniqueIntersections = new List<XYZ>();
                        foreach (XYZ inter in intersections)
                        {
                            bool isDuplicate = false;
                            foreach (XYZ existing in uniqueIntersections)
                            {
                                if (inter.DistanceTo(existing) < 0.01)
                                {
                                    isDuplicate = true;
                                    break;
                                }
                            }

                            if (!isDuplicate)
                            {
                                uniqueIntersections.Add(inter);
                            }
                        }

                        if (uniqueIntersections.Count >= 2)
                        {
                            XYZ pt1, pt2;

                            if (uniqueIntersections.Count > 2)
                            {
                                double maxDistance = 0;
                                (XYZ, XYZ) bestPair = (uniqueIntersections[0], uniqueIntersections[1]);

                                for (int k = 0; k < uniqueIntersections.Count; k++)
                                {
                                    for (int l = k + 1; l < uniqueIntersections.Count; l++)
                                    {
                                        double dist = uniqueIntersections[k].DistanceTo(uniqueIntersections[l]);
                                        if (dist > maxDistance)
                                        {
                                            maxDistance = dist;
                                            bestPair = (uniqueIntersections[k], uniqueIntersections[l]);
                                        }
                                    }
                                }

                                pt1 = bestPair.Item1;
                                pt2 = bestPair.Item2;
                            }
                            else
                            {
                                pt1 = uniqueIntersections[0];
                                pt2 = uniqueIntersections[1];
                            }

                            // SLOPE CALCULATION - Calculate Z for both points
                            double z1 = CalculateZOnSlope(startPt, pt1, slopeEdge, slopeAngleRad, baseZ);
                            double z2 = CalculateZOnSlope(startPt, pt2, slopeEdge, slopeAngleRad, baseZ);

                            XYZ pt1WithSlope = new XYZ(pt1.X, pt1.Y, z1);
                            XYZ pt2WithSlope = new XYZ(pt2.X, pt2.Y, z2);

                            DebugLog($"  Beam {i}: Z1={z1:F3}, Z2={z2:F3}");

                            FamilyInstance beam = CreateBeamDirect(pt1WithSlope, pt2WithSlope, beamType, level, doc);
                            if (beam != null)
                            {
                                beamCount++;
                                DebugLog($"  Slope beam {i} created");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLog($"  Slope beam {i} error: {ex.Message}");
                    }
                }

                DebugLog($"Slope layer complete: {beamCount} beams");
                return beamCount;
            }
            catch (Exception ex)
            {
                DebugLog($"ERROR in slope layer: {ex.Message}");
                return 0;
            }
        }

        // ============================================================================
        // MAIN EXECUTE METHOD
        // ============================================================================
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            Autodesk.Revit.DB.View activeView = uidoc.ActiveView;

            try
            {
                DebugLog("=".PadRight(60, '='));
                DebugLog("BEAM SYSTEM TOOL - WITH SLOPE SUPPORT");
                DebugLog("=".PadRight(60, '='));

                ShowMessage(
                    "INTEGRATED BEAM SYSTEM TOOL\n\n" +
                    "Process:\n" +
                    "1. Select sloped face\n" +
                    "2. Select 3 Beam Types\n" +
                    "3. Enter spacing\n" +
                    "4. Create beams\n\n" +
                    "Start now...",
                    "Beam System Tool"
                );

                FaceData faceData = PickSingleFace(uidoc, doc);
                if (faceData == null)
                {
                    ShowError("No face selected!");
                    return Result.Cancelled;
                }

                DebugLog("Face selected");

                if (!IsFaceSloped(faceData.Face))
                {
                    ShowError("Selected face must be sloped!");
                    return Result.Failed;
                }

                List<XYZ> originalPoints = ExtractFacePoints(faceData.Face, faceData.Element);
                if (originalPoints == null || originalPoints.Count == 0)
                {
                    ShowError("Cannot extract points from face!");
                    return Result.Failed;
                }

                DebugLog($"Extracted {originalPoints.Count} points");

                List<XYZ> cleanPoints = CleanupDuplicatePoints(originalPoints);
                DebugLog($"After cleanup: {cleanPoints.Count} unique points");

                if (cleanPoints.Count < 3)
                {
                    ShowError("Insufficient points!");
                    return Result.Failed;
                }

                List<XYZ> sortedPoints = SortPointsCCW(cleanPoints);
                DebugLog("Points sorted CCW");

                (Level level, bool levelOk) = GetLevelFromElement(faceData.Element, doc);
                if (!levelOk)
                {
                    ShowError("Cannot find level!");
                    return Result.Failed;
                }

                DebugLog($"Level found: {level.Name}");

                EdgeData baseEdge = FindBaseEdge(sortedPoints);
                if (baseEdge == null)
                {
                    ShowError("Cannot find base edge!");
                    return Result.Failed;
                }

                DebugLog($"Base edge found: length={baseEdge.Length:F2}");

                EdgeData slopeEdge = FindSlopeEdge(sortedPoints, baseEdge.Direction);
                if (slopeEdge == null)
                {
                    ShowError("Cannot find slope edge!");
                    return Result.Failed;
                }

                DebugLog($"Slope edge found: angle={slopeEdge.SlopeAngle:F2}°");

                double slopeAngle = slopeEdge.SlopeAngle;
                double zFace = sortedPoints.Min(pt => pt.Z);

                DebugLog($"Slope angle={slopeAngle:F2}°, Min Z={zFace:F2}");

                // Select beam types
                (FamilySymbol beamType1, string name1) = ShowBeamTypeSelector(1, doc);
                if (beamType1 == null)
                    return Result.Cancelled;

                (FamilySymbol beamType2, string name2) = ShowBeamTypeSelector(2, doc);
                if (beamType2 == null)
                    return Result.Cancelled;

                (FamilySymbol beamType3, string name3) = ShowBeamTypeSelector(3, doc);
                if (beamType3 == null)
                    return Result.Cancelled;

                DebugLog($"Beam types: {name1}, {name2}, {name3}");

                double t1 = GetBeamHeight(beamType1);
                double t2 = GetBeamHeight(beamType2);
                double t3 = GetBeamHeight(beamType3);

                DebugLog($"Beam heights (feet): t1={t1:F3}, t2={t2:F3}, t3={t3:F3}");
                DebugLog($"Beam heights (cm): t1={t1 * 30.48:F1}, t2={t2 * 30.48:F1}, t3={t3 * 30.48:F1}");

                if (t1 <= 0 || t2 <= 0 || t3 <= 0)
                {
                    ShowError("Invalid beam heights!");
                    return Result.Failed;
                }

                (double? spacing1, double? spacing2, double? spacing3) = ShowSpacingDialog();
                if (!spacing1.HasValue)
                    return Result.Cancelled;

                DebugLog($"Spacings (feet): s1={spacing1:F3}, s2={spacing2:F3}, s3={spacing3:F3}");
                DebugLog($"Spacings (cm): s1={spacing1 * 30.48:F1}, s2={spacing2 * 30.48:F1}, s3={spacing3 * 30.48:F1}");

                // Calculate beam center heights - accounting for slope distance
                double slopeAngleRad = slopeAngle * (Math.PI / 180.0);
                double cosSlope = Math.Cos(slopeAngleRad);

                // Center of each layer (from bottom)
                double z1Center = zFace + (t1 / 2.0) / cosSlope;
                double z2Center = zFace + (t1 / cosSlope) + (t2 / 2.0) / cosSlope;
                double z3Center = zFace + (t1 / cosSlope) + (t2 / cosSlope) + (t3 / 2.0) / cosSlope;

                DebugLog("\n=== HEIGHT CALCULATION (accounting for slope distance) ===");
                DebugLog($"Slope angle: {slopeAngle:F1}°, cos(slope) = {cosSlope:F4}");
                DebugLog("All heights converted to slope distance by dividing by cos(slope)");
                DebugLog($"Z_face (min Z): {zFace:F3} ft = {zFace * 30.48:F1} cm");
                DebugLog($"Z_1_center = Z_face + (t1/2)/cos = {z1Center:F3} ft = {z1Center * 30.48:F1} cm");
                DebugLog($"Z_2_center = Z_face + t1/cos + (t2/2)/cos = {z2Center:F3} ft = {z2Center * 30.48:F1} cm");
                DebugLog($"Z_3_center = Z_face + t1/cos + t2/cos + (t3/2)/cos = {z3Center:F3} ft = {z3Center * 30.48:F1} cm");
                DebugLog($"Distance L1 to L2 center-to-center: {z2Center - z1Center:F3} ft = {(z2Center - z1Center) * 30.48:F1} cm");
                DebugLog($"Distance L2 to L3 center-to-center: {z3Center - z2Center:F3} ft = {(z3Center - z2Center) * 30.48:F1} cm");

                // Create transaction
                using (Transaction t = new Transaction(doc, "Create 3-Layer Beam System"))
                {
                    t.Start();

                    try
                    {
                        DebugLog("\nLayer 1 (Horizontal - Bottom)...");
                        int beamCount1 = CreateHorizontalLayerBeamsWithSlope(
                            sortedPoints, baseEdge, slopeEdge, spacing1.Value, z1Center,
                            beamType1, level, slopeAngle, doc);

                        DebugLog("\nLayer 2 (Slope - Middle)...");
                        int beamCount2 = CreateSlopeLayerBeamsWithSlope(
                            sortedPoints, slopeEdge, spacing2.Value, z2Center,
                            beamType2, level, slopeAngle, doc);

                        DebugLog("\nLayer 3 (Horizontal - Top)...");
                        int beamCount3 = CreateHorizontalLayerBeamsWithSlope(
                            sortedPoints, baseEdge, slopeEdge, spacing3.Value, z3Center,
                            beamType3, level, slopeAngle, doc);

                        t.Commit();

                        DebugLog("=== TRANSACTION COMMITTED ===");
                        DebugLog($"Total: Layer1={beamCount1}, Layer2={beamCount2}, Layer3={beamCount3}, TOTAL={beamCount1 + beamCount2 + beamCount3}");

                        string resultMessage = "SUCCESS - Beam System Created!\n\n" +
                                              "BEAM COUNTS:\n" +
                                              $"• Layer 1 (Bottom): {beamCount1} beams\n" +
                                              $"• Layer 2 (Middle): {beamCount2} beams\n" +
                                              $"• Layer 3 (Top): {beamCount3} beams\n" +
                                              $"• TOTAL: {beamCount1 + beamCount2 + beamCount3} beams\n\n" +
                                              $"SLOPE ANGLE: {slopeAngle:F1}°";

                        ShowMessage(resultMessage, "SUCCESS");

                        return Result.Succeeded;
                    }
                    catch (Exception ex)
                    {
                        t.RollBack();
                        DebugLog($"TRANSACTION ERROR: {ex.Message}");
                        ShowError($"Transaction failed: {ex.Message}");
                        return Result.Failed;
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLog($"CRITICAL ERROR: {ex.Message}");
                DebugLog(ex.StackTrace);
                ShowError($"Critical error: {ex.Message}");
                return Result.Failed;
            }
        }
    }
}