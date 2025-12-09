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

// QUAN TRỌNG: Sử dụng alias để tránh ambiguous references
using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;

namespace SimpleBIM.Commands.As
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class TryWallFace : IExternalCommand
    {
        // ============================================================================
        // MESSAGE FUNCTIONS
        // ============================================================================
        private void ShowMessage(string message, string title = "Thong bao")
        {
            WinForms.MessageBox.Show(message, title, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
        }

        private void ShowError(string message, string title = "Loi")
        {
            WinForms.MessageBox.Show(message, title, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
        }

        // ============================================================================
        // EDGE DATA STRUCTURE
        // ============================================================================
        private class EdgeData
        {
            public Line Curve { get; set; }
            public XYZ Start { get; set; }
            public XYZ End { get; set; }
            public double Length { get; set; }
            public double MinZ { get; set; }
            public double MaxZ { get; set; }
            public int Index { get; set; }
        }

        // ============================================================================
        // POINT DATA STRUCTURE
        // ============================================================================
        private class PointData
        {
            public XYZ Point { get; set; }
            public double X { get; set; }
            public double Y { get; set; }
            public double Z { get; set; }
            public double Position { get; set; } // For sorting
        }

        // ============================================================================
        // EXTRACT FACE EDGES - FIX FOR FAMILY INSTANCE
        // ============================================================================
        private List<EdgeData> ExtractFaceEdges(Face face, Element element)
        {
            try
            {
                var edgesData = new List<EdgeData>();

                int edgeIndex = 0;

                foreach (EdgeArray edgeLoop in face.EdgeLoops)
                {
                    foreach (Edge edge in edgeLoop)
                    {
                        try
                        {
                            Curve curve = edge.AsCurve();
                            if (curve != null)
                            {
                                XYZ startPt = curve.GetEndPoint(0);
                                XYZ endPt = curve.GetEndPoint(1);

                                Line newLine = Line.CreateBound(startPt, endPt);

                                double edgeLength;
                                try
                                {
                                    edgeLength = curve.Length;
                                }
                                catch
                                {
                                    edgeLength = startPt.DistanceTo(endPt);
                                }

                                edgesData.Add(new EdgeData
                                {
                                    Curve = newLine,
                                    Start = startPt,
                                    End = endPt,
                                    Length = edgeLength,
                                    MinZ = Math.Min(startPt.Z, endPt.Z),
                                    MaxZ = Math.Max(startPt.Z, endPt.Z),
                                    Index = edgeIndex
                                });

                                edgeIndex++;
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }

                return edgesData.Count > 0 ? edgesData : null;
            }
            catch
            {
                return null;
            }
        }

        // ============================================================================
        // FIND LOWEST HORIZONTAL EDGE
        // ============================================================================
        private EdgeData FindLowestHorizontalEdge(List<EdgeData> edgesData)
        {
            try
            {
                if (edgesData == null || edgesData.Count == 0)
                    return null;

                double zTolerance = 0.01;
                var horizontalEdges = new List<EdgeData>();

                foreach (EdgeData edge in edgesData)
                {
                    double startZ = edge.Start.Z;
                    double endZ = edge.End.Z;
                    double zDiff = Math.Abs(startZ - endZ);

                    if (zDiff <= zTolerance)
                    {
                        horizontalEdges.Add(edge);
                    }
                }

                if (horizontalEdges.Count == 0)
                {
                    EdgeData lowestEdge = edgesData.OrderBy(e => Math.Abs(e.Start.Z - e.End.Z)).FirstOrDefault();
                    return lowestEdge;
                }
                else
                {
                    EdgeData lowestEdge = horizontalEdges.OrderBy(e => e.MinZ).FirstOrDefault();
                    return lowestEdge;
                }
            }
            catch
            {
                return null;
            }
        }

        // ============================================================================
        // EXTRACT ALL POINTS FROM EDGES
        // ============================================================================
        private List<PointData> ExtractAllPointsFromEdges(List<EdgeData> edgesData)
        {
            try
            {
                var pointsList = new List<PointData>();

                for (int idx = 0; idx < edgesData.Count; idx++)
                {
                    EdgeData edgeData = edgesData[idx];
                    XYZ startPt = edgeData.Start;
                    XYZ endPt = edgeData.End;

                    pointsList.Add(new PointData
                    {
                        Point = startPt,
                        X = startPt.X,
                        Y = startPt.Y,
                        Z = startPt.Z
                    });

                    pointsList.Add(new PointData
                    {
                        Point = endPt,
                        X = endPt.X,
                        Y = endPt.Y,
                        Z = endPt.Z
                    });
                }

                return pointsList;
            }
            catch
            {
                return new List<PointData>();
            }
        }

        // ============================================================================
        // SORT POINTS BY POSITION ON WALL
        // ============================================================================
        private List<PointData> SortPointsByPosition(List<PointData> pointsList, XYZ lowestEdgeStart, XYZ lowestEdgeEnd)
        {
            try
            {
                if (pointsList == null || pointsList.Count == 0)
                    return new List<PointData>();

                double dirX = lowestEdgeEnd.X - lowestEdgeStart.X;
                double dirY = lowestEdgeEnd.Y - lowestEdgeStart.Y;
                double dirLength = Math.Sqrt(dirX * dirX + dirY * dirY);

                if (dirLength < 1e-10)
                    return pointsList;

                dirX /= dirLength;
                dirY /= dirLength;

                var sortedPoints = new List<PointData>();

                foreach (PointData ptData in pointsList)
                {
                    double px = ptData.X;
                    double py = ptData.Y;

                    double toPtX = px - lowestEdgeStart.X;
                    double toPtY = py - lowestEdgeStart.Y;

                    double projection = toPtX * dirX + toPtY * dirY;

                    sortedPoints.Add(new PointData
                    {
                        Point = ptData.Point,
                        X = ptData.X,
                        Y = ptData.Y,
                        Z = ptData.Z,
                        Position = projection
                    });
                }

                return sortedPoints.OrderBy(p => p.Position).ToList();
            }
            catch
            {
                return pointsList;
            }
        }

        // ============================================================================
        // CALCULATE HEIGHT FROM SORTED POINTS
        // ============================================================================
        private (double heightInternal, double heightMm, double baseOffset, double baseOffsetMm)
            CalculateWallHeight(List<PointData> sortedPoints, Level level)
        {
            try
            {
                if (sortedPoints == null || sortedPoints.Count == 0)
                    return (0, 0, 0, 0);

                double levelElev = level.Elevation;

                double minZ = sortedPoints.Min(p => p.Z);
                double maxZ = sortedPoints.Max(p => p.Z);

                double heightInternal = maxZ - minZ;
                double heightMm = UnitUtils.ConvertFromInternalUnits(heightInternal, UnitTypeId.Millimeters);

                double baseOffset = minZ - levelElev;
                double baseOffsetMm = UnitUtils.ConvertFromInternalUnits(baseOffset, UnitTypeId.Millimeters);

                return (heightInternal, heightMm, baseOffset, baseOffsetMm);
            }
            catch
            {
                return (0, 0, 0, 0);
            }
        }

        // ============================================================================
        // GET FACE NORMAL
        // ============================================================================
        private XYZ GetFaceNormal(Face face)
        {
            try
            {
                XYZ normal = face.ComputeNormal(new UV(0.5, 0.5));
                return normal;
            }
            catch
            {
                return null;
            }
        }

        // ============================================================================
        // OFFSET EDGE BY VECTOR
        // ============================================================================
        private (XYZ offsetStart, XYZ offsetEnd) OffsetEdgeByDistance(
            XYZ edgeStart, XYZ edgeEnd, XYZ normalVector, double offsetDistance)
        {
            try
            {
                if (normalVector == null)
                    return (edgeStart, edgeEnd);

                double normalLength = Math.Sqrt(
                    normalVector.X * normalVector.X +
                    normalVector.Y * normalVector.Y +
                    normalVector.Z * normalVector.Z);

                if (normalLength < 1e-10)
                    return (edgeStart, edgeEnd);

                double normalX = normalVector.X / normalLength;
                double normalY = normalVector.Y / normalLength;
                double normalZ = normalVector.Z / normalLength;

                XYZ offsetStart = new XYZ(
                    edgeStart.X + normalX * offsetDistance,
                    edgeStart.Y + normalY * offsetDistance,
                    edgeStart.Z + normalZ * offsetDistance
                );

                XYZ offsetEnd = new XYZ(
                    edgeEnd.X + normalX * offsetDistance,
                    edgeEnd.Y + normalY * offsetDistance,
                    edgeEnd.Z + normalZ * offsetDistance
                );

                return (offsetStart, offsetEnd);
            }
            catch
            {
                return (edgeStart, edgeEnd);
            }
        }

        // ============================================================================
        // GET LEVEL FROM ELEMENT
        // ============================================================================
        private (Level level, bool ok) GetLevelFromElement(Element element, Document doc)
        {
            try
            {
                if (element.LevelId != null && element.LevelId.Value > 0)
                {
                    Level level = doc.GetElement(element.LevelId) as Level;
                    if (level != null)
                    {
                        return (level, true);
                    }
                }

                Parameter levelParam = element.LookupParameter("Level");
                if (levelParam != null)
                {
                    ElementId levelId = levelParam.AsElementId();
                    if (levelId != null && levelId.Value > 0)
                    {
                        Level level = doc.GetElement(levelId) as Level;
                        if (level != null)
                        {
                            return (level, true);
                        }
                    }
                }

                try
                {
                    BoundingBoxXYZ bbox = element.get_BoundingBox(null);
                    if (bbox != null)
                    {
                        double zCoord = bbox.Min.Z;

                        FilteredElementCollector collector = new FilteredElementCollector(doc);
                        IList<Element> allLevels = collector.OfClass(typeof(Level)).ToElements();

                        Level closestLevel = null;
                        double closestDist = double.MaxValue;

                        foreach (Element elem in allLevels)
                        {
                            Level lv = elem as Level;
                            if (lv != null)
                            {
                                double levelElev = lv.Elevation;
                                double dist = Math.Abs(zCoord - levelElev);
                                if (dist < closestDist)
                                {
                                    closestDist = dist;
                                    closestLevel = lv;
                                }
                            }
                        }

                        if (closestLevel != null)
                        {
                            return (closestLevel, true);
                        }
                    }
                }
                catch
                {
                    // Continue to return false
                }

                return (null, false);
            }
            catch
            {
                return (null, false);
            }
        }

        // ============================================================================
        // WALL TYPE FUNCTIONS
        // ============================================================================
        private string GetWallTypeName(WallType wallTypeElement)
        {
            try
            {
                if (!string.IsNullOrEmpty(wallTypeElement.Name))
                {
                    return wallTypeElement.Name;
                }

                try
                {
                    Parameter param = wallTypeElement.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM);
                    if (param != null && param.HasValue)
                    {
                        string value = param.AsString();
                        if (!string.IsNullOrEmpty(value))
                            return value;
                    }
                }
                catch { }

                try
                {
                    Parameter param = wallTypeElement.LookupParameter("Type Name");
                    if (param != null && param.HasValue)
                    {
                        string value = param.AsString();
                        if (!string.IsNullOrEmpty(value))
                            return value;
                    }
                }
                catch { }

                try
                {
                    string elementName = wallTypeElement.ToString();
                    if (!string.IsNullOrEmpty(elementName))
                        return elementName;
                }
                catch { }

                return "WallType_" + wallTypeElement.Id.Value.ToString();
            }
            catch
            {
                return "Unknown";
            }
        }

        private double GetWallTypeThickness(WallType wallType)
        {
            try
            {
                Parameter param = wallType.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM);
                if (param != null && param.HasValue && param.AsDouble() > 0)
                {
                    double thickness = UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters);
                    return thickness;
                }
                return 0;
            }
            catch
            {
                return 0;
            }
        }

        private Dictionary<string, WallType> GetAllWallTypes(Document doc)
        {
            var wallTypes = new Dictionary<string, WallType>();
            try
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                IList<Element> allWalls = collector.OfClass(typeof(WallType)).ToElements();

                foreach (Element elem in allWalls)
                {
                    WallType wallType = elem as WallType;
                    if (wallType != null)
                    {
                        try
                        {
                            string typeName = GetWallTypeName(wallType);
                            if (!string.IsNullOrEmpty(typeName) && typeName != "Unknown" && !wallTypes.ContainsKey(typeName))
                            {
                                wallTypes.Add(typeName, wallType);
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }

                return wallTypes.Count > 0 ? wallTypes : null;
            }
            catch
            {
                return null;
            }
        }

        // ============================================================================
        // WALL PARAMETERS FORM
        // ============================================================================
        private class WallFromFaceForm : WinForms.Form
        {
            private Dictionary<string, WallType> wallTypesDict;
            public WallType SelectedWallType { get; private set; }
            private WinForms.ComboBox cmbWallType;

            public WallFromFaceForm(Dictionary<string, WallType> wallTypesDict)
            {
                this.wallTypesDict = wallTypesDict;
                this.SelectedWallType = null;
                InitializeComponent();
            }

            private void InitializeComponent()
            {
                this.Text = "Tao Tuong Tu Multi Faces";
                this.Width = 500;
                this.Height = 250;
                this.StartPosition = WinForms.FormStartPosition.CenterScreen;
                this.BackColor = Drawing.Color.White;
                this.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;

                var lblTitle = new WinForms.Label();
                lblTitle.Text = "TAO DONG LOAT WALLS TU MULTI FACES";
                lblTitle.Font = new Drawing.Font("Arial", 12, Drawing.FontStyle.Bold);
                lblTitle.ForeColor = Drawing.Color.DarkBlue;
                lblTitle.Location = new Drawing.Point(20, 20);
                lblTitle.Size = new Drawing.Size(460, 25);
                this.Controls.Add(lblTitle);

                var lblWallType = new WinForms.Label();
                lblWallType.Text = "Loai Tuong:";
                lblWallType.Font = new Drawing.Font("Arial", 10);
                lblWallType.Location = new Drawing.Point(20, 60);
                lblWallType.Size = new Drawing.Size(100, 20);
                this.Controls.Add(lblWallType);

                cmbWallType = new WinForms.ComboBox();
                cmbWallType.Location = new Drawing.Point(130, 60);
                cmbWallType.Size = new Drawing.Size(280, 25);
                cmbWallType.Font = new Drawing.Font("Arial", 9);
                cmbWallType.DropDownStyle = WinForms.ComboBoxStyle.DropDownList;

                var wallTypeNames = wallTypesDict.Keys.OrderBy(k => k).ToList();
                foreach (string name in wallTypeNames)
                {
                    WallType wallType = wallTypesDict[name];
                    double thickness = GetWallTypeThickness(wallType);
                    string displayName = thickness > 0 ?
                        $"{name} ({thickness:F0}mm)" :
                        name;
                    cmbWallType.Items.Add(displayName);
                }

                if (wallTypeNames.Count > 0)
                    cmbWallType.SelectedIndex = 0;

                this.Controls.Add(cmbWallType);

                var lblInfo = new WinForms.Label();
                lblInfo.Text = "✓ Tao dong loat walls tu cac faces da chon\n✓ Offset = thickness/2 | Height = max Z - min Z";
                lblInfo.Font = new Drawing.Font("Arial", 9, Drawing.FontStyle.Bold);
                lblInfo.ForeColor = Drawing.Color.DarkGreen;
                lblInfo.Location = new Drawing.Point(20, 100);
                lblInfo.Size = new Drawing.Size(390, 35);
                this.Controls.Add(lblInfo);

                var btnOk = new WinForms.Button();
                btnOk.Text = "TAO WALLS";
                btnOk.Location = new Drawing.Point(120, 150);
                btnOk.Size = new Drawing.Size(130, 35);
                btnOk.BackColor = Drawing.Color.LightBlue;
                btnOk.Font = new Drawing.Font("Arial", 10, Drawing.FontStyle.Bold);
                btnOk.Click += BtnOK_Click;
                this.Controls.Add(btnOk);

                var btnCancel = new WinForms.Button();
                btnCancel.Text = "HUY";
                btnCancel.Location = new Drawing.Point(260, 150);
                btnCancel.Size = new Drawing.Size(90, 35);
                btnCancel.Font = new Drawing.Font("Arial", 10);
                btnCancel.Click += BtnCancel_Click;
                this.Controls.Add(btnCancel);
            }

            private double GetWallTypeThickness(WallType wallType)
            {
                try
                {
                    Parameter param = wallType.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM);
                    if (param != null && param.HasValue && param.AsDouble() > 0)
                    {
                        double thickness = UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters);
                        return thickness;
                    }
                    return 0;
                }
                catch
                {
                    return 0;
                }
            }

            private void BtnOK_Click(object sender, EventArgs e)
            {
                if (cmbWallType.SelectedIndex < 0)
                {
                    WinForms.MessageBox.Show("Vui long chon loai tuong!", "Canh bao");
                    return;
                }

                try
                {
                    string displayName = cmbWallType.SelectedItem.ToString();

                    string wallTypeName;
                    if (displayName.Contains(" ("))
                    {
                        wallTypeName = displayName.Split(new string[] { " (" }, StringSplitOptions.None)[0];
                    }
                    else
                    {
                        wallTypeName = displayName;
                    }

                    if (wallTypesDict.ContainsKey(wallTypeName))
                    {
                        SelectedWallType = wallTypesDict[wallTypeName];
                        this.DialogResult = WinForms.DialogResult.OK;
                        this.Close();
                    }
                    else
                    {
                        WinForms.MessageBox.Show($"Khong tim thay wall type: {wallTypeName}", "Loi");
                    }
                }
                catch (Exception ex)
                {
                    WinForms.MessageBox.Show($"Loi: {ex.Message}", "Loi");
                }
            }

            private void BtnCancel_Click(object sender, EventArgs e)
            {
                this.DialogResult = WinForms.DialogResult.Cancel;
                this.Close();
            }
        }

        // ============================================================================
        // CREATE WALL
        // ============================================================================
        private Wall CreateWall(
            EdgeData lowestEdge, XYZ offsetStart, XYZ offsetEnd,
            WallType wallType, Level level, double heightInternal, double baseOffset,
            Document doc)
        {
            try
            {
                Line wallCurve = Line.CreateBound(offsetStart, offsetEnd);

                Wall wall = Wall.Create(
                    doc,
                    wallCurve,
                    wallType.Id,
                    level.Id,
                    heightInternal,
                    baseOffset,
                    false,
                    false
                );

                return wall;
            }
            catch
            {
                return null;
            }
        }

        // ============================================================================
        // PROCESS ONE FACE
        // ============================================================================
        private (Wall wall, string status) ProcessFace(
            Reference faceRef, WallType wallType, Document doc)
        {
            try
            {
                Element element = doc.GetElement(faceRef.ElementId);
                GeometryObject geometryObject = element.GetGeometryObjectFromReference(faceRef);

                if (geometryObject == null)
                    return (null, "Cannot get geometry");

                Face face = geometryObject as Face;
                if (face == null)
                    return (null, "Not a valid face");

                List<EdgeData> edgesData = ExtractFaceEdges(face, element);
                if (edgesData == null || edgesData.Count == 0)
                    return (null, "Cannot extract edges");

                EdgeData lowestEdge = FindLowestHorizontalEdge(edgesData);
                if (lowestEdge == null)
                    return (null, "Cannot find lowest edge");

                List<PointData> allPoints = ExtractAllPointsFromEdges(edgesData);
                if (allPoints == null || allPoints.Count == 0)
                    return (null, "Cannot extract points");

                List<PointData> sortedPoints = SortPointsByPosition(allPoints, lowestEdge.Start, lowestEdge.End);

                (Level level, bool levelOk) = GetLevelFromElement(element, doc);
                if (!levelOk || level == null)
                    return (null, "Cannot find level");

                (double heightInternal, double heightMm, double baseOffset, double baseOffsetMm) =
                    CalculateWallHeight(sortedPoints, level);

                if (heightInternal <= 0)
                {
                    return (null, "Invalid height");
                }

                XYZ normalVector = GetFaceNormal(face);
                if (normalVector == null)
                {
                    return (null, "Cannot get normal vector");
                }

                double wallThickness = GetWallTypeThickness(wallType);
                double offsetDistanceMm = wallThickness / 2.0;
                double offsetDistanceInternal = UnitUtils.ConvertToInternalUnits(offsetDistanceMm, UnitTypeId.Millimeters);

                (XYZ offsetStart, XYZ offsetEnd) = OffsetEdgeByDistance(
                    lowestEdge.Start,
                    lowestEdge.End,
                    normalVector,
                    offsetDistanceInternal
                );

                double edgeLength = offsetStart.DistanceTo(offsetEnd);
                if (edgeLength < 0.001)
                    return (null, "Edge too short");

                Wall wall = CreateWall(lowestEdge, offsetStart, offsetEnd, wallType, level, heightInternal, baseOffset, doc);

                if (wall != null)
                {
                    double heightMmDisplay = UnitUtils.ConvertFromInternalUnits(heightInternal, UnitTypeId.Millimeters);
                    int heightDisplay = (int)Math.Round(heightMmDisplay);
                    return (wall, $"OK | H={heightDisplay}mm");
                }
                else
                {
                    return (null, "Wall creation failed");
                }
            }
            catch (Exception ex)
            {
                return (null, ex.Message);
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

            try
            {
                ShowMessage(
                    "TAO DONG LOAT WALLS TU MULTI FACES\n\n" +
                    "1. Chon TAT CA faces (Shift/Ctrl + Click)\n" +
                    "2. Chon loai tuong\n" +
                    "3. Tao dong loat walls\n\n" +
                    "Bat dau...",
                    "Huong dan"
                );

                IList<Reference> refs;
                try
                {
                    refs = uidoc.Selection.PickObjects(
                        ObjectType.Face,
                        "Chon cac faces de tao walls");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    ShowMessage("Ban da huy chon faces.", "Thong bao");
                    return Result.Cancelled;
                }

                if (refs == null || refs.Count == 0)
                {
                    ShowError("Khong chon face nao!", "Loi");
                    return Result.Failed;
                }

                Dictionary<string, WallType> wallTypesDict = GetAllWallTypes(doc);
                if (wallTypesDict == null || wallTypesDict.Count == 0)
                {
                    ShowError("Khong tim thay wall type!", "Loi");
                    return Result.Failed;
                }

                using (WallFromFaceForm form = new WallFromFaceForm(wallTypesDict))
                {
                    if (form.ShowDialog() != WinForms.DialogResult.OK)
                    {
                        ShowMessage("Ban da huy.", "Thong bao");
                        return Result.Cancelled;
                    }

                    WallType wallType = form.SelectedWallType;
                    if (wallType == null)
                    {
                        ShowError("Khong chon duoc wall type!", "Loi");
                        return Result.Failed;
                    }

                    string wallTypeName = GetWallTypeName(wallType);

                    using (Transaction t = new Transaction(doc, "Create Walls from Faces"))
                    {
                        t.Start();

                        try
                        {
                            int successCount = 0;
                            int failCount = 0;
                            List<string> resultsInfo = new List<string>();

                            for (int idx = 0; idx < refs.Count; idx++)
                            {
                                Reference faceRef = refs[idx];
                                (Wall wall, string status) = ProcessFace(faceRef, wallType, doc);

                                if (wall != null)
                                {
                                    successCount++;
                                    resultsInfo.Add($"Face {idx + 1}: [✓] {status}");
                                }
                                else
                                {
                                    failCount++;
                                    resultsInfo.Add($"Face {idx + 1}: [✗] {status}");
                                }
                            }

                            t.Commit();

                            string msg = "========== KET QUA ==========\n\n";
                            msg += $"Tong chon: {refs.Count}\n";
                            msg += $"Tao thanh cong: {successCount}\n";
                            msg += $"That bai: {failCount}\n";
                            msg += $"Wall Type: {wallTypeName}\n";
                            double thicknessVal = Math.Round(GetWallTypeThickness(wallType));
                            msg += $"Thickness: {thicknessVal}mm\n\n";
                            msg += "CHI TIET:\n";
                            msg += new string('=', 40) + "\n";
                            foreach (string info in resultsInfo)
                            {
                                msg += info + "\n";
                            }

                            ShowMessage(msg, "Thanh cong");
                        }
                        catch (Exception ex)
                        {
                            t.RollBack();
                            ShowError($"Loi: {ex.Message}", "Loi");
                            return Result.Failed;
                        }
                    }
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                ShowError($"Loi: {ex.Message}", "Loi");
                return Result.Failed;
            }
        }
    }
}