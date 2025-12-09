using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

// QUAN TRỌNG: Sử dụng alias để tránh ambiguous references
using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;
using DB = Autodesk.Revit.DB;

namespace SimpleBIM.Commands.MEPF
{
    /// <summary>
    /// CREATE CABLE TRAYS FROM CAD LINES (v1.0 - CABLE TRAY CREATION)
    /// Converted from Python to C# - FULL CONVERSION (883 lines Python source)
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CableTraysFromCAD : IExternalCommand
    {
        // =============================================================================
        // PARAMETERS
        // =============================================================================
        private const double SAME_WALL_THRESHOLD_M = 0.5;
        private const double PARALLEL_THRESHOLD = 0.999;
        private const double MIN_OVERLAP_LENGTH_M = 0.01;
        private const double MAX_DISTANCE_M = 2.5;
        private const double MIN_DISTANCE_M = 0.08;
        private const double MIN_LINE_LENGTH_M = 0.20;
        private const double COINCIDENT_THRESHOLD_M = 0.02;
        private const double EPSILON_M = 0.001;
        private const double WEIGHT_DISTANCE = 0.5;
        private const double WEIGHT_OVERLAP = 0.3;
        private const double WEIGHT_PARALLEL = 0.2;
        private const double POINT_TOLERANCE_M = 0.01;

        // Cable Tray parameters (mặc định)
        private const double CABLE_TRAY_HEIGHT_MM = 100;
        private const double CABLE_TRAY_MIDDLE_ELEVATION_MM = 3000;

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
                    return Result.Cancelled;

                // Bước 2: Lấy danh sách layer từ CAD link
                ShowMessage("Đang trích xuất danh sách layer từ CAD...", "Thông báo");
                List<string> layers = GetLayersFromCad(cadLink);
                if (layers == null || layers.Count == 0)
                {
                    ShowError("Không tìm thấy layer nào trong CAD link.", "Lỗi");
                    return Result.Cancelled;
                }

                // Bước 3: Chọn layer
                string selectedLayer = SelectLayer(layers);
                if (selectedLayer == null)
                {
                    ShowMessage("Bạn chưa chọn layer.", "Thông báo");
                    return Result.Cancelled;
                }

                // Bước 4: Thu thập lines từ layer được chọn
                ShowMessage($"Đang thu thập line từ layer: {selectedLayer}", "Thông báo");
                List<LineData> allLinesData = CollectLinesFromCadAndLayer(cadLink, selectedLayer);
                if (allLinesData == null || allLinesData.Count == 0)
                    return Result.Cancelled;

                // Bước 5: Xử lý geometry và tìm cặp song song
                var (coincidentPoints, report1) = FindCoincidentPoints(allLinesData);

                string report2;
                if (coincidentPoints != null && coincidentPoints.Count > 0)
                {
                    report2 = SplitCoincidentPoints(allLinesData, coincidentPoints);
                }
                else
                {
                    report2 = "Bước 2: Tách điểm trùng\nKhông có điểm trùng\n";
                }

                List<ParallelPair> parallelPairs = FindParallelPairs(allLinesData);

                if (parallelPairs == null || parallelPairs.Count == 0)
                {
                    ShowError($"Không tìm thấy cặp line song song trong layer: {selectedLayer}", "Lỗi");
                    return Result.Cancelled;
                }

                string report3 = $"Bước 3: Tìm cặp song song\nTìm được {parallelPairs.Count} cặp\n";

                ShowMessage(report1 + report2 + report3, "Tiến trình");

                // Bước 6: Chọn Cable Tray Type
                ShowMessage("Vui lòng chọn loại thang máng cáp", "Thông báo");

                Dictionary<string, CableTrayType> cableTrayTypesDict = GetAllCableTrayTypes();
                if (cableTrayTypesDict == null || cableTrayTypesDict.Count == 0)
                    return Result.Cancelled;

                string selectedCableTrayTypeName = SelectCableTrayType(cableTrayTypesDict);

                if (selectedCableTrayTypeName == null)
                {
                    ShowMessage("Bạn chưa chọn cable tray type.", "Thông báo");
                    return Result.Cancelled;
                }

                CableTrayType selectedCableTrayTypeObj = cableTrayTypesDict[selectedCableTrayTypeName];

                // Xác định Level
                Level level = GetActiveViewLevel();
                if (level == null)
                {
                    ShowError("Không thể xác định Level để tạo thang máng cáp.", "Lỗi");
                    return Result.Cancelled;
                }

                // Thông báo cấu hình
                string configMsg = "========== CẤU HÌNH THANG MÁNG CÁP ==========\n\n";
                configMsg += $"Loại Thang Máng Cáp: {selectedCableTrayTypeName}\n";
                configMsg += $"Chiều Cao (Height): {CABLE_TRAY_HEIGHT_MM}mm\n";
                configMsg += $"Cao độ (Middle Elevation): {CABLE_TRAY_MIDDLE_ELEVATION_MM}mm\n";
                configMsg += $"\nSẵn sàng tạo {parallelPairs.Count} thang máng cáp từ {parallelPairs.Count} cặp line";

                ShowMessage(configMsg, "Cấu hình");

                // Bước 7: Tạo Cable Tray
                int successCount = 0;
                int failedCount = 0;

                using (Transaction t = new Transaction(_doc, $"Create Cable Trays from CAD Layer: {selectedLayer}"))
                {
                    t.Start();

                    try
                    {
                        for (int pairIdx = 0; pairIdx < parallelPairs.Count; pairIdx++)
                        {
                            ParallelPair pair = parallelPairs[pairIdx];
                            try
                            {
                                // Lấy khoảng cách giữa 2 line → chiều rộng thang máng cáp
                                double cableTrayWidthMm = UnitUtils.ConvertFromInternalUnits(pair.Distance, UnitTypeId.Millimeters);
                                double cableTrayWidthInternal = pair.Distance;

                                System.Diagnostics.Debug.WriteLine($"DEBUG: Cặp #{pairIdx + 1} - Width: {cableTrayWidthMm}mm");

                                // Tính centerline
                                Line centerLine = ComputeCenterlineFromPairCurves(pair.Line1, pair.Line2);
                                if (centerLine == null)
                                {
                                    failedCount++;
                                    System.Diagnostics.Debug.WriteLine($"DEBUG: Không thể tính centerline cho cặp {pairIdx + 1}");
                                    continue;
                                }

                                // Tạo cable tray
                                CableTray cableTray = CreateCableTrayFromCenterline(
                                    centerLine, 
                                    selectedCableTrayTypeObj, 
                                    cableTrayWidthInternal, 
                                    level);

                                if (cableTray != null)
                                {
                                    successCount++;
                                    System.Diagnostics.Debug.WriteLine($"DEBUG: Đã tạo cable tray #{pairIdx + 1} thành công - Width: {cableTrayWidthMm}mm");
                                }
                                else
                                {
                                    failedCount++;
                                    System.Diagnostics.Debug.WriteLine($"DEBUG: Không thể tạo cable tray #{pairIdx + 1}");
                                }
                            }
                            catch (Exception ex)
                            {
                                failedCount++;
                                System.Diagnostics.Debug.WriteLine($"Lỗi xử lý cặp {pairIdx + 1}: {ex.Message}");
                            }
                        }

                        t.Commit();

                        // Báo cáo kết quả
                        string msg = "========== KẾT QUẢ TẠO THANG MÁNG CÁP ==========\n\n";
                        msg += $"Layer: {selectedLayer}\n";
                        msg += $"Cable Tray Type: {selectedCableTrayTypeName}\n";
                        msg += $"Height (mặc định): {CABLE_TRAY_HEIGHT_MM}mm\n";
                        msg += $"Middle Elevation: {CABLE_TRAY_MIDDLE_ELEVATION_MM}mm\n\n";
                        msg += $"✓ Thang máng cáp tạo được: {successCount}\n";
                        msg += $"✗ Thang máng cáp thất bại: {failedCount}\n";
                        msg += "━━━━━━━━━━━━━━━━━━\n";
                        msg += $"Tổng cộng: {successCount + failedCount}";

                        ShowMessage(msg, "Hoàn thành");
                        return Result.Succeeded;
                    }
                    catch (Exception ex)
                    {
                        t.RollBack();
                        ShowError($"Lỗi tạo thang máng cáp: {ex.Message}", "Lỗi");
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

                    // PYREVIT → C# CONVERSION: Cannot access Category directly from GeometryObject
                    // Must use GraphicsStyleId to get GraphicsStyle, then get GraphicsStyleCategory
                    if (geomObj.GraphicsStyleId != ElementId.InvalidElementId)
                    {
                        GraphicsStyle graphicsStyle = _doc.GetElement(geomObj.GraphicsStyleId) as GraphicsStyle;
                        if (graphicsStyle != null)
                        {
                            Category category = graphicsStyle.GraphicsStyleCategory;
                            if (category != null && !string.IsNullOrEmpty(category.Name))
                            {
                                layerSet.Add(category.Name);
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    continue;
                }
            }
        }

        private List<string> GetLayersFromCad(ImportInstance cadLink)
        {
            try
            {
                Options opts = new Options
                {
                    DetailLevel = ViewDetailLevel.Fine,
                    ComputeReferences = true
                };

                GeometryElement geo = cadLink.get_Geometry(opts);

                if (geo == null)
                {
                    ShowMessage("Không có geometry trong CAD link được chọn.", "Thông báo");
                    return null;
                }

                HashSet<string> layerNames = new HashSet<string>();
                ExtractLayersFromGeometry(geo, layerNames);

                return layerNames.OrderBy(x => x).ToList();
            }
            catch (Exception ex)
            {
                ShowError($"Lỗi khi đọc layer từ CAD: {ex.Message}", "Lỗi");
                return null;
            }
        }

        // =============================================================================
        // LAYER SELECTION FORM
        // =============================================================================
        private string SelectLayer(List<string> layers)
        {
            using (LayerSelectionForm form = new LayerSelectionForm(layers))
            {
                if (form.ShowDialog() == WinForms.DialogResult.OK)
                {
                    return form.SelectedLayer;
                }
                return null;
            }
        }

        private class LayerSelectionForm : WinForms.Form
        {
            public string SelectedLayer { get; private set; }
            private List<string> _layers;
            private WinForms.Label _lblTitle;
            private WinForms.Label _lblInstruction;
            private WinForms.ListBox _lstLayers;
            private WinForms.Button _btnOK;
            private WinForms.Button _btnCancel;

            public LayerSelectionForm(List<string> layers)
            {
                SelectedLayer = null;
                _layers = layers;
                InitializeComponent();
            }

            private void InitializeComponent()
            {
                Text = "Chọn Layer từ CAD";
                Width = 400;
                Height = 500;
                StartPosition = WinForms.FormStartPosition.CenterScreen;
                BackColor = Drawing.Color.White;

                _lblTitle = new WinForms.Label();
                _lblTitle.Text = "CHỌN LAYER ĐỂ TẠO THANG MÁNG CÁP";
                _lblTitle.Font = new Drawing.Font("Arial", 12, Drawing.FontStyle.Bold);
                _lblTitle.ForeColor = Drawing.Color.DarkBlue;
                _lblTitle.Location = new Drawing.Point(20, 20);
                _lblTitle.Size = new Drawing.Size(350, 25);
                Controls.Add(_lblTitle);

                _lblInstruction = new WinForms.Label();
                _lblInstruction.Text = "Chọn 1 layer từ danh sách bên dưới:";
                _lblInstruction.Font = new Drawing.Font("Arial", 9);
                _lblInstruction.Location = new Drawing.Point(20, 50);
                _lblInstruction.Size = new Drawing.Size(350, 20);
                Controls.Add(_lblInstruction);

                _lstLayers = new WinForms.ListBox();
                _lstLayers.Location = new Drawing.Point(20, 80);
                _lstLayers.Size = new Drawing.Size(340, 300);
                _lstLayers.Font = new Drawing.Font("Arial", 9);
                _lstLayers.SelectionMode = WinForms.SelectionMode.One;
                foreach (string layer in _layers)
                {
                    _lstLayers.Items.Add(layer);
                }
                Controls.Add(_lstLayers);

                _btnOK = new WinForms.Button();
                _btnOK.Text = "CHỌN LAYER NÀY";
                _btnOK.Location = new Drawing.Point(120, 400);
                _btnOK.Size = new Drawing.Size(120, 30);
                _btnOK.BackColor = Drawing.Color.LightBlue;
                _btnOK.Click += BtnOK_Click;
                Controls.Add(_btnOK);

                _btnCancel = new WinForms.Button();
                _btnCancel.Text = "HỦY";
                _btnCancel.Location = new Drawing.Point(250, 400);
                _btnCancel.Size = new Drawing.Size(80, 30);
                _btnCancel.Click += BtnCancel_Click;
                Controls.Add(_btnCancel);
            }

            private void BtnOK_Click(object sender, EventArgs e)
            {
                if (_lstLayers.SelectedItem != null)
                {
                    SelectedLayer = _lstLayers.SelectedItem.ToString();
                    DialogResult = WinForms.DialogResult.OK;
                    Close();
                }
                else
                {
                    WinForms.MessageBox.Show("Vui lòng chọn một layer!", "Thông báo");
                }
            }

            private void BtnCancel_Click(object sender, EventArgs e)
            {
                DialogResult = WinForms.DialogResult.Cancel;
                Close();
            }
        }

        // =============================================================================
        // CABLE TRAY TYPE SELECTION
        // =============================================================================
        private Dictionary<string, CableTrayType> GetAllCableTrayTypes()
        {
            Dictionary<string, CableTrayType> cableTrayTypes = new Dictionary<string, CableTrayType>();
            try
            {
                FilteredElementCollector collector = new FilteredElementCollector(_doc).OfClass(typeof(CableTrayType));
                List<CableTrayType> cableTrayTypesList = collector.Cast<CableTrayType>().ToList();

                if (cableTrayTypesList.Count == 0)
                {
                    ShowError("Không tìm thấy Cable Tray Type nào trong dự án.", "Lỗi");
                    return null;
                }

                foreach (CableTrayType trayType in cableTrayTypesList)
                {
                    try
                    {
                        string name = trayType.Name;
                        cableTrayTypes[name] = trayType;
                    }
                    catch
                    {
                        try
                        {
                            Parameter param = trayType.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM);
                            string name = param != null ? param.AsString() : "Unknown";
                            cableTrayTypes[name] = trayType;
                        }
                        catch
                        {
                            cableTrayTypes["Unknown"] = trayType;
                        }
                    }
                }

                System.Diagnostics.Debug.WriteLine($"✅ Tìm thấy {cableTrayTypes.Count} Cable Tray Types.");
                return cableTrayTypes;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ DEBUG: Error collecting cable tray types: {ex.Message}");
                ShowError($"Lỗi khi lấy Cable Tray Types: {ex.Message}", "Lỗi");
                return null;
            }
        }

        private string SelectCableTrayType(Dictionary<string, CableTrayType> cableTrayTypesDict)
        {
            using (CableTrayTypeSelectionForm form = new CableTrayTypeSelectionForm(cableTrayTypesDict))
            {
                if (form.ShowDialog() == WinForms.DialogResult.OK)
                {
                    return form.SelectedCableTrayType;
                }
                return null;
            }
        }

        private class CableTrayTypeSelectionForm : WinForms.Form
        {
            public string SelectedCableTrayType { get; private set; }
            private Dictionary<string, CableTrayType> _cableTrayTypesDict;
            private WinForms.Label _lblTitle;
            private WinForms.Label _lblCableTrayType;
            private WinForms.ComboBox _cboCableTrayType;
            private WinForms.Label _lblInfo;
            private WinForms.Button _btnOK;
            private WinForms.Button _btnCancel;

            public CableTrayTypeSelectionForm(Dictionary<string, CableTrayType> cableTrayTypesDict)
            {
                SelectedCableTrayType = null;
                _cableTrayTypesDict = cableTrayTypesDict;
                InitializeComponent();
            }

            private void InitializeComponent()
            {
                Text = "Chọn Loại Thang Máng Cáp";
                Width = 450;
                Height = 350;
                StartPosition = WinForms.FormStartPosition.CenterScreen;
                BackColor = Drawing.Color.White;

                // Title
                _lblTitle = new WinForms.Label();
                _lblTitle.Text = "CẤU HÌNH THANG MÁNG CÁP";
                _lblTitle.Font = new Drawing.Font("Arial", 12, Drawing.FontStyle.Bold);
                _lblTitle.ForeColor = Drawing.Color.DarkBlue;
                _lblTitle.Location = new Drawing.Point(20, 20);
                _lblTitle.Size = new Drawing.Size(400, 25);
                Controls.Add(_lblTitle);

                // Cable Tray Type Label
                _lblCableTrayType = new WinForms.Label();
                _lblCableTrayType.Text = "Loại Thang Máng Cáp (Cable Tray Type):";
                _lblCableTrayType.Font = new Drawing.Font("Arial", 10, Drawing.FontStyle.Bold);
                _lblCableTrayType.Location = new Drawing.Point(20, 60);
                _lblCableTrayType.Size = new Drawing.Size(300, 20);
                Controls.Add(_lblCableTrayType);

                // Cable Tray Type ComboBox
                _cboCableTrayType = new WinForms.ComboBox();
                _cboCableTrayType.Location = new Drawing.Point(20, 85);
                _cboCableTrayType.Size = new Drawing.Size(400, 25);
                _cboCableTrayType.Font = new Drawing.Font("Arial", 9);
                _cboCableTrayType.DropDownStyle = WinForms.ComboBoxStyle.DropDownList;
                foreach (string trayName in _cableTrayTypesDict.Keys)
                {
                    _cboCableTrayType.Items.Add(trayName);
                }
                if (_cboCableTrayType.Items.Count > 0)
                {
                    _cboCableTrayType.SelectedIndex = 0;
                }
                Controls.Add(_cboCableTrayType);

                // Info Label
                _lblInfo = new WinForms.Label();
                _lblInfo.Text = "Thông số mặc định:\n• Height (Chiều cao): 100mm\n• Middle Elevation: 3000mm\n• Width (Chiều rộng): Theo khoảng cách CAD";
                _lblInfo.Font = new Drawing.Font("Arial", 9);
                _lblInfo.Location = new Drawing.Point(20, 130);
                _lblInfo.Size = new Drawing.Size(400, 80);
                Controls.Add(_lblInfo);

                // OK Button
                _btnOK = new WinForms.Button();
                _btnOK.Text = "TIẾP TỤC";
                _btnOK.Location = new Drawing.Point(150, 260);
                _btnOK.Size = new Drawing.Size(100, 30);
                _btnOK.BackColor = Drawing.Color.LightGreen;
                _btnOK.Click += BtnOK_Click;
                Controls.Add(_btnOK);

                // Cancel Button
                _btnCancel = new WinForms.Button();
                _btnCancel.Text = "HỦY";
                _btnCancel.Location = new Drawing.Point(270, 260);
                _btnCancel.Size = new Drawing.Size(80, 30);
                _btnCancel.Click += BtnCancel_Click;
                Controls.Add(_btnCancel);
            }

            private void BtnOK_Click(object sender, EventArgs e)
            {
                if (_cboCableTrayType.SelectedItem != null)
                {
                    SelectedCableTrayType = _cboCableTrayType.SelectedItem.ToString();
                    DialogResult = WinForms.DialogResult.OK;
                    Close();
                }
                else
                {
                    WinForms.MessageBox.Show("Vui lòng chọn Cable Tray Type!", "Thông báo");
                }
            }

            private void BtnCancel_Click(object sender, EventArgs e)
            {
                DialogResult = WinForms.DialogResult.Cancel;
                Close();
            }
        }

        // =============================================================================
        // GEOMETRY EXTRACTION WITH LAYER FILTER
        // =============================================================================
        private List<Curve> GetAllCurvesFromImportInstance(ImportInstance importInstance, string targetLayer = null)
        {
            List<Curve> curves = new List<Curve>();
            try
            {
                Options options = new Options
                {
                    DetailLevel = ViewDetailLevel.Fine,
                    ComputeReferences = true
                };

                GeometryElement geometryElement = importInstance.get_Geometry(options);
                if (geometryElement == null)
                    return curves;

                foreach (GeometryObject geoObj in geometryElement)
                {
                    curves.AddRange(ExtractCurvesFromGeometryObject(geoObj, targetLayer));
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Lỗi khi trích xuất geometry: {ex.Message}");
            }
            return curves;
        }

        private List<Curve> ExtractCurvesFromGeometryObject(GeometryObject geoObj, string targetLayer = null)
        {
            List<Curve> curves = new List<Curve>();
            try
            {
                if (geoObj is Line line)
                {
                    if (line.Length > 1e-6)
                    {
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
                                if (targetLayer == null || BelongsToLayer(instGeo, targetLayer))
                                {
                                    if (instGeo is Line instLine)
                                    {
                                        XYZ startPoint = transform.OfPoint(instLine.GetEndPoint(0));
                                        XYZ endPoint = transform.OfPoint(instLine.GetEndPoint(1));
                                        Line transformedLine = Line.CreateBound(startPoint, endPoint);
                                        if (transformedLine.Length > 1e-6)
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
                                                Line polyLineSegment = Line.CreateBound(p1, p2);
                                                if (polyLineSegment != null && polyLineSegment.Length > 1e-6)
                                                {
                                                    curves.Add(polyLineSegment);
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
                        System.Diagnostics.Debug.WriteLine($"Lỗi xử lý GeometryInstance: {ex.Message}");
                    }
                }
                else if (geoObj is PolyLine polyLine2)
                {
                    if (targetLayer == null || BelongsToLayer(geoObj, targetLayer))
                    {
                        IList<XYZ> points = polyLine2.GetCoordinates();
                        for (int i = 0; i < points.Count - 1; i++)
                        {
                            try
                            {
                                Line polyLineSegment = Line.CreateBound(points[i], points[i + 1]);
                                if (polyLineSegment != null && polyLineSegment.Length > 1e-6)
                                {
                                    curves.Add(polyLineSegment);
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
                System.Diagnostics.Debug.WriteLine($"Lỗi trích xuất curves: {ex.Message}");
            }
            return curves;
        }

        private bool BelongsToLayer(GeometryObject geometryObj, string layerName)
        {
            try
            {
                if (geometryObj.GraphicsStyleId != ElementId.InvalidElementId)
                {
                    GraphicsStyle graphicsStyle = _doc.GetElement(geometryObj.GraphicsStyleId) as GraphicsStyle;
                    if (graphicsStyle != null)
                    {
                        Category category = graphicsStyle.GraphicsStyleCategory;
                        if (category != null && category.Name == layerName)
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
        private List<LineData> CollectLinesFromCadAndLayer(ImportInstance cadLink, string layerName)
        {
            List<LineData> allLinesData = new List<LineData>();
            try
            {
                List<Curve> curves = GetAllCurvesFromImportInstance(cadLink, layerName);
                foreach (Curve curve in curves)
                {
                    if (curve is Line line)
                    {
                        if (line.Length < MIN_LINE_LENGTH)
                            continue;
                        LineData lineData = ExtractLineData(line);
                        if (lineData != null)
                        {
                            allLinesData.Add(lineData);
                        }
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
                    return null;
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
        private (double, double, double) PointToKey(XYZ point, double toleranceM)
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
                    precision = Math.Max(0, (int)(-Math.Log10(toleranceM)));
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

        private (List<CoincidentPoint>, string) FindCoincidentPoints(List<LineData> allLinesData)
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

            List<CoincidentPoint> coincidentPoints = new List<CoincidentPoint>();
            foreach (var kvp in pointMap)
            {
                if (kvp.Value.Count > 1)
                {
                    coincidentPoints.Add(new CoincidentPoint
                    {
                        PointKey = kvp.Key,
                        LineList = kvp.Value
                    });
                }
            }

            string report = $"Bước 1: Tìm điểm trùng\nTìm được {coincidentPoints.Count} điểm trùng\n";
            return (coincidentPoints, report);
        }

        private string SplitCoincidentPoints(List<LineData> allLinesData, List<CoincidentPoint> coincidentPoints)
        {
            double epsilonInternal = EPSILON;
            int modificationsCount = 0;

            foreach (CoincidentPoint coincident in coincidentPoints)
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

        private List<ParallelPair> FindParallelPairs(List<LineData> allLinesData)
        {
            List<ParallelPair> parallelPairs = new List<ParallelPair>();

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
                        continue;

                    XYZ vec12 = lineData2.Start - lineData1.Start;
                    double perpDist = vec12.CrossProduct(dir1).GetLength();

                    if (perpDist < MIN_DISTANCE || perpDist > MAX_DISTANCE)
                        continue;

                    XYZ p1End = lineData1.End;
                    XYZ p2Start = lineData2.Start;
                    double distEndStart = PointsDistance(p1End, p2Start);
                    bool isCoincident = distEndStart < COINCIDENT_THRESHOLD;

                    double overlapLen = CalculateOverlap(lineData1, lineData2, dir1);

                    if (!isCoincident && overlapLen < MIN_OVERLAP_LENGTH)
                        continue;

                    parallelPairs.Add(new ParallelPair
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

                var (t2a, _) = GetProjectionOnLine(p2a, p1a, p1b);
                var (t2b, __) = GetProjectionOnLine(p2b, p1a, p1b);

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

        private (double, XYZ) GetProjectionOnLine(XYZ point, XYZ lineStart, XYZ lineEnd)
        {
            XYZ lineVec = lineEnd - lineStart;
            XYZ pointVec = point - lineStart;
            double lineLenSq = lineVec.DotProduct(lineVec);
            if (lineLenSq < 1e-9)
                return (0.0, lineStart);
            double t = pointVec.DotProduct(lineVec) / lineLenSq;
            t = Math.Max(0.0, Math.Min(1.0, t));
            XYZ projPoint = lineStart + lineVec * t;
            return (t, projPoint);
        }

        // =============================================================================
        // CABLE TRAY CREATION
        // =============================================================================
        private Line ComputeCenterlineFromPairCurves(Line line1, Line line2)
        {
            XYZ p1a = line1.GetEndPoint(0);
            XYZ p1b = line1.GetEndPoint(1);
            XYZ p2a = line2.GetEndPoint(0);
            XYZ p2b = line2.GetEndPoint(1);

            var (t2a, proj2a) = GetProjectionOnLine(p2a, p1a, p1b);
            var (t2b, proj2b) = GetProjectionOnLine(p2b, p1a, p1b);

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

            var (_, start2) = GetProjectionOnLine(start1, p2a, p2b);
            var (__, end2) = GetProjectionOnLine(end1, p2a, p2b);

            XYZ centerStart = (start1 + start2) / 2.0;
            XYZ centerEnd = (end1 + end2) / 2.0;

            return Line.CreateBound(centerStart, centerEnd);
        }

        private Level GetActiveViewLevel()
        {
            try
            {
                Autodesk.Revit.DB.View activeView = _uidoc.ActiveView;

                if (activeView is ViewPlan viewPlan)
                {
                    Level level = viewPlan.GenLevel;
                    if (level != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"DEBUG: Sử dụng Level từ active view: {level.Name}");
                        return level;
                    }
                }

                ShowMessage("Không thể xác định Level từ view hiện tại. Sẽ sử dụng Level thấp nhất.", "Thông báo");

                FilteredElementCollector collector = new FilteredElementCollector(_doc).OfClass(typeof(Level));
                List<Level> levels = collector.Cast<Level>().ToList();
                if (levels.Count == 0)
                {
                    ShowError("Không có Level trong project.", "Lỗi");
                    return null;
                }

                Level lowestLevel = levels.OrderBy(x => x.Elevation).First();
                System.Diagnostics.Debug.WriteLine($"DEBUG: Sử dụng Level thấp nhất: {lowestLevel.Name}");
                return lowestLevel;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Lỗi khi lấy active view level: {ex.Message}");
                return null;
            }
        }

        private CableTray CreateCableTrayFromCenterline(Line centerLine, CableTrayType cableTrayTypeObj, 
            double cableTrayWidthInternal, Level level)
        {
            try
            {
                // Chuyển đổi thông số
                double cableTrayHeightInternal = UnitUtils.ConvertToInternalUnits(CABLE_TRAY_HEIGHT_MM, UnitTypeId.Millimeters);
                double middleElevRelative = UnitUtils.ConvertToInternalUnits(CABLE_TRAY_MIDDLE_ELEVATION_MM, UnitTypeId.Millimeters);

                // Tạo Cable Tray object
                CableTray cableTray = CableTray.Create(_doc, cableTrayTypeObj.Id,
                                                      centerLine.GetEndPoint(0),
                                                      centerLine.GetEndPoint(1),
                                                      level.Id);

                if (cableTray != null)
                {
                    // Thiết lập Width (chiều rộng)
                    try
                    {
                        Parameter widthParam = cableTray.LookupParameter("Width");
                        if (widthParam != null && !widthParam.IsReadOnly)
                        {
                            widthParam.Set(cableTrayWidthInternal);
                        }
                    }
                    catch
                    {
                        // Ignore
                    }

                    // Thiết lập Height (chiều cao)
                    try
                    {
                        Parameter heightParam = cableTray.LookupParameter("Height");
                        if (heightParam != null && !heightParam.IsReadOnly)
                        {
                            heightParam.Set(cableTrayHeightInternal);
                        }
                    }
                    catch
                    {
                        // Ignore
                    }

                    // Thiết lập Middle Elevation (tương đối từ Level hiện tại)
                    try
                    {
                        Parameter middleElevParam = cableTray.LookupParameter("Middle Elevation");
                        if (middleElevParam != null && !middleElevParam.IsReadOnly)
                        {
                            middleElevParam.Set(middleElevRelative);
                        }
                    }
                    catch
                    {
                        // Ignore
                    }
                }

                return cableTray;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Lỗi tạo cable tray: {ex.Message}");
                return null;
            }
        }

        // =============================================================================
        // UTILITY FUNCTIONS
        // =============================================================================
        private void ShowMessage(string messageText, string title = "Thông báo")
        {
            WinForms.MessageBox.Show(messageText, title, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
        }

        private void ShowError(string messageText, string title = "Lỗi")
        {
            WinForms.MessageBox.Show(messageText, title, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
        }

        // =============================================================================
        // DATA STRUCTURES
        // =============================================================================
        private class LineData
        {
            public Line Line { get; set; }
            public XYZ Direction { get; set; }
            public XYZ Start { get; set; }
            public XYZ End { get; set; }
            public double Length { get; set; }
        }

        private class CoincidentPoint
        {
            public (double, double, double) PointKey { get; set; }
            public List<(int, string, XYZ)> LineList { get; set; }
        }

        private class ParallelPair
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
    }
}

/* 
**PYREVIT → C# CONVERSIONS APPLIED (FULL LIST):**
1. `__revit__` / `doc` / `uidoc` → Retrieved from ExternalCommandData
2. `forms.alert()` → `WinForms.MessageBox.Show()` / `TaskDialog.Show()`
3. `print()` → `System.Diagnostics.Debug.WriteLine()`
4. `element.Id.IntegerValue` → `element.Id.Value`
5. Python `dict` → C# `Dictionary<TKey, TValue>` and custom classes
6. Python `set` → C# `HashSet<T>`
7. Python tuple → C# ValueTuple `(type1, type2, type3)`
8. `geometry_object.Category` → `doc.GetElement(geometryObject.GraphicsStyleId) as GraphicsStyle` then `.GraphicsStyleCategory`
9. Python Forms (System.Windows.Forms) → C# WinForms với alias
10. Transaction: `with Transaction()` → `using (Transaction t = new Transaction())`
11. Python list comprehension → C# LINQ `.Where().ToList()` or for loops
12. Python string formatting `.format()` → C# string interpolation `$"{}"`

**THAM KHẢO TỪ Commands/As/:**
- IExternalCommand structure  
- CAD link selection pattern
- Layer extraction logic
- Geometry traversal patterns
- GUI forms implementation
- Transaction management

**IMPORTANT NOTES:**
- FULL CONVERSION của 883 dòng Python source
- Geometry category access: PHẢI dùng GraphicsStyleId → GraphicsStyle → GraphicsStyleCategory (KHÔNG thể access trực tiếp)
- Windows Forms: sử dụng alias `WinForms` và `Drawing` để tránh ambiguous references
- Parallel line detection: đầy đủ logic tìm cặp song song với overlap calculation
- Coincident points handling: split points với epsilon offset
- Cable tray creation với full parameters: Width, Height, Middle Elevation
- Level detection: ưu tiên từ active view, fallback là level thấp nhất
- GUI forms: LayerSelectionForm và CableTrayTypeSelectionForm với đầy đủ controls
- Data structures: LineData, CoincidentPoint, ParallelPair thay thế Python dicts
*/
