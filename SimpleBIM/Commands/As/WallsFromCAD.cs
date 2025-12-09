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

//===================== HOÀN CHỈNH =====================

namespace SimpleBIM.Commands.As
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class WallsFromCAD : IExternalCommand
    {
        // =============================================================================
        // PARAMETERS
        // =============================================================================
        private const double SAME_WALL_THRESHOLD_M = 0.5;
        private const double PARALLEL_THRESHOLD = 0.999;
        private const double MIN_OVERLAP_LENGTH_M = 0.01;
        private const double MAX_DISTANCE_M = 0.45;
        private const double MIN_DISTANCE_M = 0.08;
        private const double MIN_LINE_LENGTH_M = 0.20;
        private const double COINCIDENT_THRESHOLD_M = 0.02;
        private const double EPSILON_M = 0.001;
        private const double WEIGHT_DISTANCE = 0.5;
        private const double WEIGHT_OVERLAP = 0.3;
        private const double WEIGHT_PARALLEL = 0.2;
        private const double POINT_TOLERANCE_M = 0.01;

        // Internal units (converted from meters)
        private double MIN_DISTANCE;
        private double MAX_DISTANCE;
        private double SAME_WALL_THRESHOLD;
        private double MIN_OVERLAP_LENGTH;
        private double MIN_LINE_LENGTH;
        private double COINCIDENT_THRESHOLD;
        private double EPSILON;
        private double POINT_TOLERANCE;

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

            // Initialize unit conversions
            InitializeUnits();

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
                TaskDialog.Show("Lỗi", $"Lỗi thực thi: {ex.Message}\n\n{ex.StackTrace}");
                return Result.Failed;
            }
        }

        private void InitializeUnits()
        {
            MIN_DISTANCE = UnitUtils.ConvertToInternalUnits(MIN_DISTANCE_M, UnitTypeId.Meters);
            MAX_DISTANCE = UnitUtils.ConvertToInternalUnits(MAX_DISTANCE_M, UnitTypeId.Meters);
            SAME_WALL_THRESHOLD = UnitUtils.ConvertToInternalUnits(SAME_WALL_THRESHOLD_M, UnitTypeId.Meters);
            MIN_OVERLAP_LENGTH = UnitUtils.ConvertToInternalUnits(MIN_OVERLAP_LENGTH_M, UnitTypeId.Meters);
            MIN_LINE_LENGTH = UnitUtils.ConvertToInternalUnits(MIN_LINE_LENGTH_M, UnitTypeId.Meters);
            COINCIDENT_THRESHOLD = UnitUtils.ConvertToInternalUnits(COINCIDENT_THRESHOLD_M, UnitTypeId.Meters);
            EPSILON = UnitUtils.ConvertToInternalUnits(EPSILON_M, UnitTypeId.Meters);
            POINT_TOLERANCE = POINT_TOLERANCE_M;
        }

        // =============================================================================
        // CAD LINK SELECTION
        // =============================================================================

        private class CadLinkFilter : ISelectionFilter
        {
            public bool AllowElement(Element element)
            {
                return element is ImportInstance;
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }

        private ImportInstance PickCadLink()
        {
            try
            {
                Reference reference = _uidoc.Selection.PickObject(
                    ObjectType.Element,
                    new CadLinkFilter(),
                    "Chọn CAD Link trong view");

                Element element = _doc.GetElement(reference.ElementId);
                return element as ImportInstance;
            }
            catch (Exception)
            {
                ShowMessage("Không chọn được CAD link hoặc đã hủy chọn.", "Thông báo");
                return null;
            }
        }

        // =============================================================================
        // LAYER EXTRACTION
        // =============================================================================

        private void ExtractLayersFromGeometry(GeometryElement geometryElement, HashSet<string> layerSet)
        {
            if (geometryElement == null) return;

            foreach (GeometryObject geomObj in geometryElement)
            {
                try
                {
                    // Xử lý GeometryInstance
                    if (geomObj is GeometryInstance geometryInstance)
                    {
                        try
                        {
                            GeometryElement instGeom = geometryInstance.GetInstanceGeometry();
                            if (instGeom != null)
                            {
                                ExtractLayersFromGeometry(instGeom, layerSet);
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }

                    // Lấy layer name từ GraphicsStyle
                    if (geomObj.GraphicsStyleId != null &&
                        geomObj.GraphicsStyleId != ElementId.InvalidElementId)
                    {
                        GraphicsStyle graphicsStyle = _doc.GetElement(geomObj.GraphicsStyleId) as GraphicsStyle;
                        if (graphicsStyle?.GraphicsStyleCategory != null)
                        {
                            Category category = graphicsStyle.GraphicsStyleCategory;
                            if (category?.Name != null)
                            {
                                layerSet.Add(category.Name);
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

        private List<string> GetLayersFromCad(ImportInstance cadLink)
        {
            try
            {
                Options opts = new Options();
                opts.DetailLevel = ViewDetailLevel.Fine;
                opts.ComputeReferences = true;

                GeometryElement geo = cadLink.get_Geometry(opts);

                if (geo == null)
                {
                    ShowMessage("Không có geometry trong CAD link được chọn.", "Thông báo");
                    return new List<string>();
                }

                HashSet<string> layerNames = new HashSet<string>();
                ExtractLayersFromGeometry(geo, layerNames);

                return layerNames.OrderBy(x => x).ToList();
            }
            catch (Exception ex)
            {
                ShowError($"Lỗi khi đọc layer từ CAD: {ex.Message}", "Lỗi");
                return new List<string>();
            }
        }

        // =============================================================================
        // LAYER SELECTION FORM
        // =============================================================================

        private class LayerSelectionForm : System.Windows.Forms.Form
        {
            public string SelectedLayer { get; private set; }
            private ListBox _lstLayers;

            public LayerSelectionForm(List<string> layers)
            {
                InitializeComponent(layers);
            }

            private void InitializeComponent(List<string> layers)
            {
                this.Text = "Chọn Layer từ CAD";
                this.Width = 400;
                this.Height = 500;
                this.StartPosition = FormStartPosition.CenterScreen;
                this.BackColor = System.Drawing.Color.White;

                // Title
                Label lblTitle = new Label
                {
                    Text = "CHỌN LAYER ĐỂ TẠO TƯỜNG",
                    Font = new Font("Arial", 12, FontStyle.Bold),
                    ForeColor = System.Drawing.Color.DarkBlue,
                    Location = new System.Drawing.Point(20, 20),
                    Size = new Size(350, 25)
                };
                this.Controls.Add(lblTitle);

                // Instruction
                Label lblInstruction = new Label
                {
                    Text = "Chọn 1 layer từ danh sách bên dưới:",
                    Font = new Font("Arial", 9),
                    Location = new System.Drawing.Point(20, 50),
                    Size = new Size(350, 20)
                };
                this.Controls.Add(lblInstruction);

                // Layer ListBox
                _lstLayers = new ListBox
                {
                    Location = new System.Drawing.Point(20, 80),
                    Size = new Size(340, 300),
                    Font = new Font("Arial", 9),
                    SelectionMode = SelectionMode.One
                };
                foreach (string layer in layers)
                {
                    _lstLayers.Items.Add(layer);
                }
                this.Controls.Add(_lstLayers);

                // Buttons
                Button btnOK = new Button
                {
                    Text = "CHỌN LAYER NÀY",
                    Location = new System.Drawing.Point(120, 400),
                    Size = new Size(120, 30),
                    BackColor = System.Drawing.Color.LightBlue
                };
                btnOK.Click += BtnOK_Click;
                this.Controls.Add(btnOK);

                Button btnCancel = new Button
                {
                    Text = "HỦY",
                    Location = new System.Drawing.Point(250, 400),
                    Size = new Size(80, 30)
                };
                btnCancel.Click += BtnCancel_Click;
                this.Controls.Add(btnCancel);
            }

            private void BtnOK_Click(object sender, EventArgs e)
            {
                if (_lstLayers.SelectedItem != null)
                {
                    SelectedLayer = _lstLayers.SelectedItem.ToString();
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Vui lòng chọn một layer!", "Thông báo");
                }
            }

            private void BtnCancel_Click(object sender, EventArgs e)
            {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            }
        }

        private string SelectLayer(List<string> layers)
        {
            using (LayerSelectionForm form = new LayerSelectionForm(layers))
            {
                DialogResult result = form.ShowDialog();
                if (result == DialogResult.OK)
                {
                    return form.SelectedLayer;
                }
            }
            return null;
        }

        // =============================================================================
        // CORE WALL TYPE FUNCTIONS
        // =============================================================================

        private Dictionary<string, (double widthMm, WallType wallType)> GetAllWallTypes()
        {
            Dictionary<string, (double, WallType)> wallTypes = new Dictionary<string, (double, WallType)>();

            try
            {
                FilteredElementCollector collector = new FilteredElementCollector(_doc)
                    .OfClass(typeof(WallType));

                if (!collector.Any())
                {
                    collector = new FilteredElementCollector(_doc)
                        .OfCategory(BuiltInCategory.OST_Walls)
                        .WhereElementIsElementType();
                }

                foreach (WallType wallType in collector.Cast<WallType>())
                {
                    try
                    {
                        string typeName = wallType.Name ?? $"Unknown_{wallType.Id}";
                        double widthMm = 0;

                        // Cách 1: Lấy từ CompoundStructure
                        try
                        {
                            CompoundStructure compoundStructure = wallType.GetCompoundStructure();
                            if (compoundStructure != null)
                            {
                                double widthFeet = compoundStructure.GetWidth();
                                widthMm = widthFeet * 304.8;
                            }
                        }
                        catch
                        {
                            // Ignore
                        }

                        // Cách 2: Lấy từ Width parameter
                        if (widthMm == 0)
                        {
                            try
                            {
                                Parameter widthParam = wallType.LookupParameter("Width");
                                if (widthParam != null && widthParam.HasValue)
                                {
                                    double widthFeet = widthParam.AsDouble();
                                    widthMm = widthFeet * 304.8;
                                }
                            }
                            catch
                            {
                                // Ignore
                            }
                        }

                        // Cách 3: Lấy từ "Thickness" parameter
                        if (widthMm == 0)
                        {
                            try
                            {
                                Parameter thickParam = wallType.LookupParameter("Thickness");
                                if (thickParam != null && thickParam.HasValue)
                                {
                                    double widthFeet = thickParam.AsDouble();
                                    widthMm = widthFeet * 304.8;
                                }
                            }
                            catch
                            {
                                // Ignore
                            }
                        }

                        // Nếu không tìm được độ dày, dùng giá trị mặc định
                        if (widthMm == 0)
                        {
                            widthMm = 200;
                        }

                        wallTypes[typeName] = (Math.Round(widthMm, 1), wallType);
                    }
                    catch
                    {
                        continue;
                    }
                }

                if (!wallTypes.Any())
                {
                    ShowError("Không tìm thấy wall type nào trong project! Vui lòng tạo ít nhất một wall type.", "Lỗi");
                    return null;
                }

                return wallTypes;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"DEBUG: Error collecting wall types: {ex.Message}");
                return null;
            }
        }

        private (double minDiff, List<(string name, double width, WallType wallType)> candidates)
            FindCandidatesWallType(double measuredThicknessMm, Dictionary<string, (double widthMm, WallType wallType)> wallTypesDict)
        {
            double minDiff = double.MaxValue;
            List<(string name, double width, WallType wallType)> candidates = new List<(string, double, WallType)>();

            foreach (var kvp in wallTypesDict)
            {
                double diff = Math.Abs(kvp.Value.widthMm - measuredThicknessMm);
                if (diff < minDiff)
                {
                    minDiff = diff;
                    candidates.Clear();
                    candidates.Add((kvp.Key, kvp.Value.widthMm, kvp.Value.wallType));
                }
                else if (Math.Abs(diff - minDiff) < 0.001) // diff == minDiff
                {
                    candidates.Add((kvp.Key, kvp.Value.widthMm, kvp.Value.wallType));
                }
            }

            return (minDiff, candidates);
        }

        private (string name, double width, WallType wallType) AskUserChooseWallType(
            int pairIndex,
            double thicknessMm,
            List<(string name, double width, WallType wallType)> candidates)
        {
            if (candidates.Count < 2)
            {
                return candidates[0];
            }

            var candidate1 = candidates[0];
            var candidate2 = candidates[1];

            string msg = $"Cặp line #{pairIndex + 1}\n";
            msg += $"Độ dày đo được: {thicknessMm:F1} mm\n\n";
            msg += "Có 2 loại tường gần nhất như nhau:\n";
            msg += $"• {candidate1.name} ({candidate1.width} mm) - Chênh lệch: {Math.Abs(candidate1.width - thicknessMm):F1} mm\n";
            msg += $"• {candidate2.name} ({candidate2.width} mm) - Chênh lệch: {Math.Abs(candidate2.width - thicknessMm):F1} mm\n\n";
            msg += "Vui lòng chọn loại tường:";

            DialogResult result = MessageBox.Show(msg, "Chọn Loại Tường",
                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                return candidate1;
            }
            else if (result == DialogResult.No)
            {
                return candidate2;
            }
            else
            {
                return (null, 0, null);
            }
        }

        // =============================================================================
        // GEOMETRY EXTRACTION WITH LAYER FILTER
        // =============================================================================

        private List<Line> GetAllCurvesFromImportInstance(ImportInstance importInstance, string targetLayer = null)
        {
            List<Line> curves = new List<Line>();
            try
            {
                Options options = new Options
                {
                    DetailLevel = ViewDetailLevel.Fine,
                    ComputeReferences = true
                };

                GeometryElement geometryElement = importInstance.get_Geometry(options);
                if (geometryElement == null)
                {
                    return curves;
                }

                foreach (GeometryObject geoObj in geometryElement)
                {
                    curves.AddRange(ExtractCurvesFromGeometryObject(geoObj, targetLayer));
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Lỗi khi trích xuất geometry: {ex.Message}");
            }
            return curves;
        }

        private List<Line> ExtractCurvesFromGeometryObject(GeometryObject geoObj, string targetLayer = null)
        {
            List<Line> curves = new List<Line>();
            try
            {
                if (geoObj is Line line)
                {
                    if (line.Length > 1e-6)
                    {
                        // Kiểm tra layer nếu có targetLayer
                        if (targetLayer == null || BelongsToLayer(geoObj, targetLayer))
                        {
                            curves.Add(line);
                        }
                    }
                }
                else if (geoObj is GeometryInstance geometryInstance)
                {
                    try
                    {
                        GeometryElement instanceGeometry = geometryInstance.GetInstanceGeometry();
                        Transform transform = geometryInstance.Transform;
                        if (instanceGeometry != null)
                        {
                            foreach (GeometryObject instGeo in instanceGeometry)
                            {
                                // Kiểm tra layer của instGeo
                                if (targetLayer == null || BelongsToLayer(instGeo, targetLayer))
                                {
                                    if (instGeo is Line instLine)
                                    {
                                        XYZ startPoint = transform.OfPoint(instLine.GetEndPoint(0));
                                        XYZ endPoint = transform.OfPoint(instLine.GetEndPoint(1));
                                        Line transformedLine = Line.CreateBound(startPoint, endPoint);
                                        if (transformedLine != null && transformedLine.Length > 1e-6)
                                        {
                                            curves.Add(transformedLine);
                                        }
                                    }
                                    else if (instGeo is PolyLine polyLine)
                                    {
                                        IList<XYZ> points = polyLine.GetCoordinates();
                                        for (int i = 0; i < points.Count - 1; i++)
                                        {
                                            try
                                            {
                                                XYZ p1 = transform.OfPoint(points[i]);
                                                XYZ p2 = transform.OfPoint(points[i + 1]);
                                                Line lineSegment = Line.CreateBound(p1, p2);
                                                if (lineSegment != null && lineSegment.Length > 1e-6)
                                                {
                                                    curves.Add(lineSegment);
                                                }
                                            }
                                            catch
                                            {
                                                continue;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Lỗi xử lý GeometryInstance: {ex.Message}");
                    }
                }
                else if (geoObj is PolyLine polyLine)
                {
                    // Kiểm tra layer nếu có targetLayer
                    if (targetLayer == null || BelongsToLayer(geoObj, targetLayer))
                    {
                        IList<XYZ> points = polyLine.GetCoordinates();
                        for (int i = 0; i < points.Count - 1; i++)
                        {
                            try
                            {
                                Line lineSegment = Line.CreateBound(points[i], points[i + 1]);
                                if (lineSegment != null && lineSegment.Length > 1e-6)
                                {
                                    curves.Add(lineSegment);
                                }
                            }
                            catch
                            {
                                continue;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Lỗi trích xuất curves: {ex.Message}");
            }
            return curves;
        }

        private bool BelongsToLayer(GeometryObject geometryObj, string layerName)
        {
            try
            {
                // Kiểm tra qua GraphicsStyle
                if (geometryObj.GraphicsStyleId != null && geometryObj.GraphicsStyleId != ElementId.InvalidElementId)
                {
                    GraphicsStyle graphicsStyle = _doc.GetElement(geometryObj.GraphicsStyleId) as GraphicsStyle;
                    if (graphicsStyle?.GraphicsStyleCategory != null)
                    {
                        Category category = graphicsStyle.GraphicsStyleCategory;
                        if (category?.Name == layerName)
                        {
                            return true;
                        }
                    }
                }
            }
            catch
            {
                // Ignore
            }
            return false;
        }

        // =============================================================================
        // LINE COLLECTION & ANALYSIS
        // =============================================================================

        private class LineData
        {
            public Line Line { get; set; }
            public XYZ Direction { get; set; }
            public XYZ Start { get; set; }
            public XYZ End { get; set; }
            public double Length { get; set; }
        }

        private List<LineData> CollectLinesFromCadAndLayer(ImportInstance cadLink, string layerName)
        {
            List<LineData> allLinesData = new List<LineData>();
            try
            {
                List<Line> curves = GetAllCurvesFromImportInstance(cadLink, layerName);
                foreach (Line curve in curves)
                {
                    if (curve.Length < MIN_LINE_LENGTH)
                    {
                        continue;
                    }

                    LineData lineData = ExtractLineData(curve);
                    if (lineData != null)
                    {
                        allLinesData.Add(lineData);
                    }
                }

                if (allLinesData.Count > 0)
                {
                    ShowMessage($"Đã tìm thấy {allLinesData.Count} line từ layer: {layerName}", "Thông báo");
                    return allLinesData;
                }
                else
                {
                    ShowError($"Không tìm thấy line nào trong layer: {layerName}", "Lỗi");
                    return null;
                }
            }
            catch (Exception ex)
            {
                ShowError($"Lỗi thu thập line: {ex.Message}", "Lỗi");
                return null;
            }
        }

        private LineData ExtractLineData(Line line)
        {
            try
            {
                if (line == null || line.Length < 1e-6)
                {
                    return null;
                }

                XYZ p0 = line.GetEndPoint(0);
                XYZ p1 = line.GetEndPoint(1);
                XYZ directionVec = (p1 - p0).Normalize();

                return new LineData
                {
                    Line = line,
                    Direction = directionVec,
                    Start = p0,
                    End = p1,
                    Length = line.Length
                };
            }
            catch
            {
                return null;
            }
        }

        // =============================================================================
        // GEOMETRY ANALYSIS
        // =============================================================================

        private (double x, double y, double z) PointToKey(XYZ point, double toleranceM)
        {
            try
            {
                int precision;
                if (toleranceM <= 0)
                {
                    precision = 6;
                }
                else
                {
                    precision = Math.Max(0, (int)-Math.Log10(toleranceM));
                }

                double x = Math.Round(point.X, precision);
                double y = Math.Round(point.Y, precision);
                double z = Math.Round(point.Z, precision);

                return (x, y, z);
            }
            catch
            {
                return (point.X, point.Y, point.Z);
            }
        }

        private double PointsDistance(XYZ p1, XYZ p2)
        {
            double dx = p1.X - p2.X;
            double dy = p1.Y - p2.Y;
            double dz = p1.Z - p2.Z;
            return Math.Sqrt(dx * dx + dy * dy + dz * dz);
        }

        private class CoincidentPointData
        {
            public (double x, double y, double z) PointKey { get; set; }
            public List<(int lineIndex, string pointType, XYZ point)> LineList { get; set; }
        }

        private (List<CoincidentPointData> coincidentPoints, string report) FindCoincidentPoints(List<LineData> allLinesData)
        {
            Dictionary<(double, double, double), List<(int, string, XYZ)>> pointMap =
                new Dictionary<(double, double, double), List<(int, string, XYZ)>>();

            for (int idx = 0; idx < allLinesData.Count; idx++)
            {
                LineData lineData = allLinesData[idx];

                var startKey = PointToKey(lineData.Start, POINT_TOLERANCE);
                var endKey = PointToKey(lineData.End, POINT_TOLERANCE);

                if (!pointMap.ContainsKey(startKey))
                {
                    pointMap[startKey] = new List<(int, string, XYZ)>();
                }
                pointMap[startKey].Add((idx, "START", lineData.Start));

                if (!pointMap.ContainsKey(endKey))
                {
                    pointMap[endKey] = new List<(int, string, XYZ)>();
                }
                pointMap[endKey].Add((idx, "END", lineData.End));
            }

            List<CoincidentPointData> coincidentPoints = new List<CoincidentPointData>();
            foreach (var kvp in pointMap)
            {
                if (kvp.Value.Count > 1)
                {
                    coincidentPoints.Add(new CoincidentPointData
                    {
                        PointKey = kvp.Key,
                        LineList = kvp.Value
                    });
                }
            }

            string report = $"Bước 1: Tìm điểm trùng\nTìm được {coincidentPoints.Count} điểm trùng\n";
            return (coincidentPoints, report);
        }

        private string SplitCoincidentPoints(List<LineData> allLinesData, List<CoincidentPointData> coincidentPoints)
        {
            double epsilonInternal = EPSILON;
            int modificationsCount = 0;

            foreach (CoincidentPointData coincident in coincidentPoints)
            {
                foreach (var (lineIdx, pointType, point) in coincident.LineList)
                {
                    LineData lineData = allLinesData[lineIdx];
                    XYZ direction = lineData.Direction;
                    XYZ offsetVec = direction * epsilonInternal;

                    if (pointType == "END")
                    {
                        lineData.End = point + offsetVec;
                    }
                    else if (pointType == "START")
                    {
                        lineData.Start = point - offsetVec;
                    }

                    Line newLine = Line.CreateBound(lineData.Start, lineData.End);
                    lineData.Line = newLine;
                    modificationsCount++;
                }
            }

            string report = $"Bước 2: Tách điểm trùng\nTách được {modificationsCount} điểm\n";
            return report;
        }

        private class ParallelPairData
        {
            public Line Line1 { get; set; }
            public Line Line2 { get; set; }
            public LineData Data1 { get; set; }
            public LineData Data2 { get; set; }
            public double Distance { get; set; }
            public double OverlapLength { get; set; }
            public double DotProduct { get; set; }
            public bool IsCoincident { get; set; }
        }

        private List<ParallelPairData> FindParallelPairs(List<LineData> allLinesData)
        {
            List<ParallelPairData> parallelPairs = new List<ParallelPairData>();

            int n = allLinesData.Count;
            for (int i = 0; i < n; i++)
            {
                for (int j = i + 1; j < n; j++)
                {
                    LineData lineData1 = allLinesData[i];
                    LineData lineData2 = allLinesData[j];

                    XYZ dir1 = lineData1.Direction;
                    XYZ dir2 = lineData2.Direction;

                    double dotProd = Math.Abs(dir1.DotProduct(dir2));
                    if (dotProd < PARALLEL_THRESHOLD)
                    {
                        continue;
                    }

                    XYZ vec12 = lineData2.Start - lineData1.Start;
                    double perpDist = vec12.CrossProduct(dir1).GetLength();

                    if (perpDist < MIN_DISTANCE || perpDist > MAX_DISTANCE)
                    {
                        continue;
                    }

                    XYZ p1End = lineData1.End;
                    XYZ p2Start = lineData2.Start;
                    double distEndStart = PointsDistance(p1End, p2Start);
                    bool isCoincident = distEndStart < COINCIDENT_THRESHOLD;

                    double overlapLen = CalculateOverlap(lineData1, lineData2, dir1);

                    if (!isCoincident && overlapLen < MIN_OVERLAP_LENGTH)
                    {
                        continue;
                    }

                    parallelPairs.Add(new ParallelPairData
                    {
                        Line1 = lineData1.Line,
                        Line2 = lineData2.Line,
                        Data1 = lineData1,
                        Data2 = lineData2,
                        Distance = perpDist,
                        OverlapLength = overlapLen,
                        DotProduct = dotProd,
                        IsCoincident = isCoincident
                    });
                }
            }

            return parallelPairs;
        }

        private double CalculateOverlap(LineData lineData1, LineData lineData2, XYZ dir1)
        {
            try
            {
                XYZ p1a = lineData1.Start;
                XYZ p1b = lineData1.End;
                XYZ p2a = lineData2.Start;
                XYZ p2b = lineData2.End;

                double t2a, t2b;
                GetProjectionOnLine(p2a, p1a, p1b, out t2a, out _);
                GetProjectionOnLine(p2b, p1a, p1b, out t2b, out _);

                double s = Math.Max(0.0, Math.Min(t2a, t2b));
                double e = Math.Min(1.0, Math.Max(t2a, t2b));

                XYZ line1Vec = p1b - p1a;
                double line1Len = line1Vec.GetLength();
                double overlapLen = (e - s) * line1Len;

                if (overlapLen < UnitUtils.ConvertToInternalUnits(0.01, UnitTypeId.Meters))
                {
                    double distEndStart = PointsDistance(p1b, p2a);
                    double distEndStart2 = PointsDistance(p2b, p1a);
                    if (distEndStart < COINCIDENT_THRESHOLD || distEndStart2 < COINCIDENT_THRESHOLD)
                    {
                        return UnitUtils.ConvertToInternalUnits(0.001, UnitTypeId.Meters);
                    }
                }

                return Math.Max(0.0, overlapLen);
            }
            catch
            {
                return 0.0;
            }
        }

        private void GetProjectionOnLine(XYZ point, XYZ lineStart, XYZ lineEnd, out double t, out XYZ projPoint)
        {
            XYZ lineVec = lineEnd - lineStart;
            XYZ pointVec = point - lineStart;
            double lineLenSq = lineVec.DotProduct(lineVec);

            if (lineLenSq < 1e-9)
            {
                t = 0.0;
                projPoint = lineStart;
                return;
            }

            t = pointVec.DotProduct(lineVec) / lineLenSq;
            t = Math.Max(0.0, Math.Min(1.0, t));
            projPoint = lineStart + lineVec * t;
        }

        // =============================================================================
        // WALL CREATION
        // =============================================================================

        private Line ComputeCenterlineFromPairCurves(Line line1, Line line2)
        {
            XYZ p1a = line1.GetEndPoint(0);
            XYZ p1b = line1.GetEndPoint(1);
            XYZ p2a = line2.GetEndPoint(0);
            XYZ p2b = line2.GetEndPoint(1);

            double t2a, t2b;
            GetProjectionOnLine(p2a, p1a, p1b, out t2a, out XYZ proj2a);
            GetProjectionOnLine(p2b, p1a, p1b, out t2b, out XYZ proj2b);

            double s = Math.Max(0.0, Math.Min(t2a, t2b));
            double e = Math.Min(1.0, Math.Max(t2a, t2b));

            if (e - s <= 1e-6)
            {
                s = 0.0;
                e = 1.0;
            }

            XYZ line1Vec = p1b - p1a;
            XYZ start1 = p1a + line1Vec * s;
            XYZ end1 = p1a + line1Vec * e;

            GetProjectionOnLine(start1, p2a, p2b, out _, out XYZ start2);
            GetProjectionOnLine(end1, p2a, p2b, out _, out XYZ end2);

            XYZ centerStart = (start1 + start2) / 2.0;
            XYZ centerEnd = (end1 + end2) / 2.0;

            return Line.CreateBound(centerStart, centerEnd);
        }

        private Wall CreateWallFromCenterline(Line centerLine, WallType wallType, Level level, double heightInternal)
        {
            try
            {
                Wall wall = Wall.Create(_doc, centerLine, wallType.Id, level.Id, heightInternal, 0.0, false, false);

                try
                {
                    Parameter param = wall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM);
                    if (param != null && !param.IsReadOnly)
                    {
                        param.Set((int)WallLocationLine.WallCenterline);
                    }
                }
                catch
                {
                    // Ignore
                }

                return wall;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Lỗi tạo wall: {ex.Message}");
                return null;
            }
        }

        // =============================================================================
        // LEVEL DETECTION FUNCTIONS
        // =============================================================================

        private Level GetActiveViewLevel()
        {
            try
            {
                Autodesk.Revit.DB.View activeView = _uidoc.ActiveView;

                // Kiểm tra nếu active view là ViewPlan (mặt bằng)
                if (activeView is ViewPlan viewPlan)
                {
                    Level level = viewPlan.GenLevel;
                    if (level != null)
                    {
                        Debug.WriteLine($"DEBUG: Sử dụng Level từ active view: {level.Name}");
                        return level;
                    }
                }

                // Nếu không phải ViewPlan hoặc không có GenLevel, thông báo và tìm Level khác
                ShowMessage("Không thể xác định Level từ view hiện tại. Sẽ sử dụng Level thấp nhất.", "Thông báo");

                // Fallback: lấy Level thấp nhất
                List<Level> levels = new FilteredElementCollector(_doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .ToList();

                if (!levels.Any())
                {
                    ShowError("Không có Level trong project.", "Lỗi");
                    return null;
                }

                Level lowestLevel = levels.OrderBy(x => x.Elevation).First();
                Debug.WriteLine($"DEBUG: Sử dụng Level thấp nhất: {lowestLevel.Name}");
                return lowestLevel;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Lỗi khi lấy active view level: {ex.Message}");
                return null;
            }
        }

        // =============================================================================
        // HELPER METHODS
        // =============================================================================

        private void ShowMessage(string message, string title = "Thông báo")
        {
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ShowError(string message, string title = "Lỗi")
        {
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static class Debug
        {
            public static void WriteLine(string message)
            {
#if DEBUG
                System.Diagnostics.Debug.WriteLine(message);
#endif
            }
        }

        // =============================================================================
        // MAIN EXECUTION
        // =============================================================================

        private Result RunMain()
        {
            try
            {
                // Bước 1: Chọn CAD link
                ShowMessage("Vui lòng chọn CAD link trong view", "Thông báo");
                ImportInstance cadLink = PickCadLink();
                if (cadLink == null)
                {
                    return Result.Cancelled;
                }

                // Bước 2: Lấy danh sách layer từ CAD link
                ShowMessage("Đang trích xuất danh sách layer từ CAD...", "Thông báo");
                List<string> layers = GetLayersFromCad(cadLink);
                if (layers == null || !layers.Any())
                {
                    ShowError("Không tìm thấy layer nào trong CAD link.", "Lỗi");
                    return Result.Failed;
                }

                // Bước 3: Chọn layer
                string selectedLayer = SelectLayer(layers);
                if (string.IsNullOrEmpty(selectedLayer))
                {
                    ShowMessage("Bạn chưa chọn layer.", "Thông báo");
                    return Result.Cancelled;
                }

                // Bước 4: Thu thập lines từ layer được chọn
                ShowMessage($"Đang thu thập line từ layer: {selectedLayer}", "Thông báo");
                List<LineData> allLinesData = CollectLinesFromCadAndLayer(cadLink, selectedLayer);
                if (allLinesData == null || !allLinesData.Any())
                {
                    return Result.Failed;
                }

                // Bước 5: Xử lý geometry và tìm cặp song song
                var (coincidentPoints, report1) = FindCoincidentPoints(allLinesData);

                string report2;
                if (coincidentPoints.Any())
                {
                    report2 = SplitCoincidentPoints(allLinesData, coincidentPoints);
                }
                else
                {
                    report2 = "Bước 2: Tách điểm trùng\nKhông có điểm trùng\n";
                }

                List<ParallelPairData> parallelPairs = FindParallelPairs(allLinesData);

                if (!parallelPairs.Any())
                {
                    ShowError($"Không tìm thấy cặp line song song trong layer: {selectedLayer}", "Lỗi");
                    return Result.Failed;
                }

                string report3 = $"Bước 3: Tìm cặp song song\nTìm được {parallelPairs.Count} cặp\n";

                ShowMessage(report1 + report2 + report3, "Tiến trình");

                // Bước 6: Lấy wall types
                Dictionary<string, (double widthMm, WallType wallType)> wallTypesDict = GetAllWallTypes();
                if (wallTypesDict == null || !wallTypesDict.Any())
                {
                    return Result.Failed;
                }

                // Bước 7: Xác định Level
                Level level = GetActiveViewLevel();
                if (level == null)
                {
                    ShowError("Không thể xác định Level để tạo tường.", "Lỗi");
                    return Result.Failed;
                }

                double heightInternal = UnitUtils.ConvertToInternalUnits(3000, UnitTypeId.Millimeters);

                int successCount = 0;
                int failedCount = 0;

                using (Transaction trans = new Transaction(_doc, $"Create Walls from CAD Layer: {selectedLayer}"))
                {
                    trans.Start();

                    try
                    {
                        for (int pairIdx = 0; pairIdx < parallelPairs.Count; pairIdx++)
                        {
                            try
                            {
                                ParallelPairData pair = parallelPairs[pairIdx];
                                double thicknessMm = UnitUtils.ConvertFromInternalUnits(pair.Distance, UnitTypeId.Millimeters);

                                var (minDiff, candidates) = FindCandidatesWallType(thicknessMm, wallTypesDict);

                                (string selectedWallTypeName, double selectedWallTypeWidth, WallType selectedWallTypeObj) selectedWallType;

                                if (candidates.Count == 1)
                                {
                                    selectedWallType = candidates[0];
                                }
                                else if (candidates.Count == 2)
                                {
                                    double diff1 = Math.Abs(candidates[0].width - thicknessMm);
                                    double diff2 = Math.Abs(candidates[1].width - thicknessMm);

                                    if (Math.Abs(diff1 - diff2) < 0.001) // diff1 == diff2
                                    {
                                        var result = AskUserChooseWallType(pairIdx, thicknessMm, candidates);
                                        if (result.name == null)
                                        {
                                            failedCount++;
                                            continue;
                                        }
                                        selectedWallType = result;
                                    }
                                    else
                                    {
                                        selectedWallType = candidates[0];
                                    }
                                }
                                else
                                {
                                    selectedWallType = candidates[0];
                                }

                                // Sử dụng trực tiếp wall type object
                                WallType wallType = selectedWallType.selectedWallTypeObj;
                                if (wallType == null)
                                {
                                    Debug.WriteLine($"DEBUG: Wall type object is None for: {selectedWallType.selectedWallTypeName}");
                                    failedCount++;
                                    continue;
                                }

                                Line centerLine = ComputeCenterlineFromPairCurves(pair.Line1, pair.Line2);
                                if (centerLine == null)
                                {
                                    failedCount++;
                                    continue;
                                }

                                Wall wall = CreateWallFromCenterline(centerLine, wallType, level, heightInternal);

                                if (wall != null)
                                {
                                    successCount++;
                                    Debug.WriteLine($"DEBUG: Đã tạo wall thành công với type: {selectedWallType.selectedWallTypeName}");
                                }
                                else
                                {
                                    failedCount++;
                                    Debug.WriteLine($"DEBUG: Không thể tạo wall với type: {selectedWallType.selectedWallTypeName}");
                                }
                            }
                            catch (Exception ex)
                            {
                                failedCount++;
                                Debug.WriteLine($"Lỗi xử lý cặp {pairIdx}: {ex.Message}");
                            }
                        }

                        trans.Commit();

                        string msg = "========== KẾT QUẢ ==========\n\n";
                        msg += $"Layer: {selectedLayer}\n";
                        msg += $"Level: {level.Name}\n";
                        msg += $"✓ Tường tạo được: {successCount}\n";
                        msg += $"✗ Tường thất bại: {failedCount}\n";

                        ShowMessage(msg, "Hoàn thành");
                        return Result.Succeeded;
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        ShowError($"Lỗi tạo tường: {ex.Message}", "Lỗi");
                        return Result.Failed;
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError($"Lỗi: {ex.Message}", "Lỗi");
                return Result.Failed;
            }
        }
    }
}