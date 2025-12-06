using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Drawing = System.Drawing;
// ALIASES TO AVOID AMBIGUOUS REFERENCES
using WinForms = System.Windows.Forms;

namespace SimpleBIM.AS.tab.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class FinishRoomsWall : IExternalCommand
    {
        // CONSTANTS
        private const double MIN_CURVE_LENGTH_M = 0.1;
        private double _minCurveLength;

        // MAIN EXECUTION
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                _minCurveLength = UnitUtils.ConvertToInternalUnits(MIN_CURVE_LENGTH_M, UnitTypeId.Meters);
                Run(uidoc, doc);
                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // User cancelled the operation
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        // MESSAGE FUNCTIONS
        private void ShowMessage(string message, string title = "Thong bao")
        {
            WinForms.MessageBox.Show(message, title,
                WinForms.MessageBoxButtons.OK,
                WinForms.MessageBoxIcon.Information);
        }

        private void ShowError(string message, string title = "Loi")
        {
            WinForms.MessageBox.Show(message, title,
                WinForms.MessageBoxButtons.OK,
                WinForms.MessageBoxIcon.Error);
        }

        // ROOM SELECTION FILTER
        private class RoomSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element element)
            {
                try
                {
                    if (element.GetType().Name == "Room")
                        return true;

                    if (element.Category != null &&
                        element.Category.Id.Value == (long)BuiltInCategory.OST_Rooms)
                        return true;

                    if (element is Room)
                        return true;

                    return false;
                }
                catch
                {
                    return false;
                }
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }

        private List<Room> PickRooms(UIDocument uidoc)
        {
            try
            {
                IList<Reference> refs = uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    new RoomSelectionFilter(),
                    "Chon 1 hoac nhieu Rooms (Ctrl+Click) de tao tuong");

                if (refs == null || refs.Count == 0)
                {
                    ShowError("Khong chon room nao!", "Loi");
                    return null;
                }

                List<Room> rooms = new List<Room>();
                Document doc = uidoc.Document;

                foreach (Reference reference in refs)
                {
                    Element element = doc.GetElement(reference.ElementId);
                    if (element is Room room)
                    {
                        rooms.Add(room);
                    }
                }

                if (rooms.Count == 0)
                {
                    ShowError("Khong tim thay room hop le!", "Loi");
                    return null;
                }

                return rooms;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                ShowMessage("Ban da huy chon Rooms.", "Thong bao");
                return null;
            }
            catch (Exception ex)
            {
                ShowError("Loi khi chon Rooms: " + ex.Message, "Loi");
                return null;
            }
        }

        // WALL TYPE FUNCTIONS
        private string GetWallTypeName(WallType wallTypeElement)
        {
            try
            {
                // Method 1: Direct name
                if (!string.IsNullOrEmpty(wallTypeElement.Name))
                    return wallTypeElement.Name;

                // Method 2: SYMBOL_NAME_PARAM
                Parameter param = wallTypeElement.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM);
                if (param != null && param.HasValue && !string.IsNullOrEmpty(param.AsString()))
                    return param.AsString();

                // Method 3: Type Name parameter
                param = wallTypeElement.LookupParameter("Type Name");
                if (param != null && param.HasValue && !string.IsNullOrEmpty(param.AsString()))
                    return param.AsString();

                // Method 4: ToString()
                string elementName = wallTypeElement.ToString();
                if (!string.IsNullOrEmpty(elementName))
                    return elementName;

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
                // Method 1: WALL_ATTR_WIDTH_PARAM
                Parameter param = wallType.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM);
                if (param != null && param.HasValue && param.AsDouble() > 0)
                {
                    return UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters);
                }

                // Method 2: Width parameter
                param = wallType.LookupParameter("Width");
                if (param != null && param.HasValue && param.AsDouble() > 0)
                {
                    return UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters);
                }

                // Method 3: Compound structure
                CompoundStructure compoundStructure = wallType.GetCompoundStructure();
                if (compoundStructure != null)
                {
                    // PYREVIT CONVERSION: GetTotalThickness() → GetWidth()
                    return UnitUtils.ConvertFromInternalUnits(
                        compoundStructure.GetWidth(),
                        UnitTypeId.Millimeters);
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
            try
            {
                Dictionary<string, WallType> wallTypes = new Dictionary<string, WallType>();
                FilteredElementCollector collector = new FilteredElementCollector(doc)
                    .OfClass(typeof(WallType));

                foreach (WallType wallType in collector.Cast<WallType>())
                {
                    try
                    {
                        string typeName = GetWallTypeName(wallType);
                        if (!string.IsNullOrEmpty(typeName) &&
                            typeName != "Unknown" &&
                            !wallTypes.ContainsKey(typeName))
                        {
                            wallTypes[typeName] = wallType;
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

                if (wallTypes.Count == 0)
                {
                    ShowError("Khong tim thay Wall Type nao trong project!", "Loi");
                    return null;
                }

                return wallTypes;
            }
            catch (Exception ex)
            {
                ShowError("Loi lay wall types: " + ex.Message, "Loi");
                return null;
            }
        }

        // FORM: WALL PARAMETERS INPUT
        private class RoomWallForm : WinForms.Form
        {
            private Dictionary<string, WallType> _wallTypesDict;
            public WallType SelectedWallType { get; private set; }
            public double HeightMm { get; private set; }
            public int WallLocationLine { get; private set; }

            private WinForms.ComboBox _cmbWallType;
            private WinForms.TextBox _txtHeight;
            private WinForms.ComboBox _cmbLocation;

            public RoomWallForm(Dictionary<string, WallType> wallTypesDict)
            {
                _wallTypesDict = wallTypesDict;
                SelectedWallType = null;
                HeightMm = 0;
                WallLocationLine = 1;
                InitializeComponent();
            }

            private void InitializeComponent()
            {
                this.Text = "Tao Tuong tu Nhieu Rooms (MULTI-SELECT)";
                this.Width = 500;
                this.Height = 350;
                this.StartPosition = WinForms.FormStartPosition.CenterScreen;
                this.BackColor = Drawing.Color.White;
                this.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;

                // TITLE
                WinForms.Label lblTitle = new WinForms.Label();
                lblTitle.Text = "TAO TUONG TU NHIEU ROOMS";
                lblTitle.Font = new Drawing.Font("Arial", 12, Drawing.FontStyle.Bold);
                lblTitle.ForeColor = Drawing.Color.DarkBlue;
                lblTitle.Location = new Drawing.Point(20, 20);
                lblTitle.Size = new Drawing.Size(460, 25);
                this.Controls.Add(lblTitle);

                // WALL TYPE
                WinForms.Label lblWallType = new WinForms.Label();
                lblWallType.Text = "Loai Tuong:";
                lblWallType.Font = new Drawing.Font("Arial", 10);
                lblWallType.Location = new Drawing.Point(20, 60);
                lblWallType.Size = new Drawing.Size(100, 20);
                this.Controls.Add(lblWallType);

                _cmbWallType = new WinForms.ComboBox();
                _cmbWallType.Location = new Drawing.Point(130, 60);
                _cmbWallType.Size = new Drawing.Size(280, 25);
                _cmbWallType.Font = new Drawing.Font("Arial", 9);
                _cmbWallType.DropDownStyle = WinForms.ComboBoxStyle.DropDownList;

                List<string> wallTypeNames = _wallTypesDict.Keys.ToList();
                wallTypeNames.Sort();

                foreach (string name in wallTypeNames)
                {
                    WallType wallType = _wallTypesDict[name];
                    double thickness = GetWallTypeThickness(wallType);
                    string displayName = thickness > 0
                        ? string.Format("{0} ({1:F0}mm)", name, thickness)
                        : name;
                    _cmbWallType.Items.Add(displayName);
                }

                if (wallTypeNames.Count > 0)
                    _cmbWallType.SelectedIndex = 0;

                this.Controls.Add(_cmbWallType);

                // HEIGHT
                WinForms.Label lblHeight = new WinForms.Label();
                lblHeight.Text = "Chieu Cao (mm):";
                lblHeight.Font = new Drawing.Font("Arial", 10);
                lblHeight.Location = new Drawing.Point(20, 110);
                lblHeight.Size = new Drawing.Size(100, 20);
                this.Controls.Add(lblHeight);

                _txtHeight = new WinForms.TextBox();
                _txtHeight.Location = new Drawing.Point(130, 110);
                _txtHeight.Size = new Drawing.Size(280, 25);
                _txtHeight.Font = new Drawing.Font("Arial", 9);
                _txtHeight.Text = "3000";
                this.Controls.Add(_txtHeight);

                // INFO
                WinForms.Label lblInfo = new WinForms.Label();
                lblInfo.Text = "(Vi du: 2800, 3000, 3200, ...)";
                lblInfo.Font = new Drawing.Font("Arial", 8);
                lblInfo.ForeColor = Drawing.Color.Gray;
                lblInfo.Location = new Drawing.Point(130, 135);
                lblInfo.Size = new Drawing.Size(280, 15);
                this.Controls.Add(lblInfo);

                // WALL LOCATION LINE
                WinForms.Label lblLocation = new WinForms.Label();
                lblLocation.Text = "Location Line:";
                lblLocation.Font = new Drawing.Font("Arial", 10);
                lblLocation.Location = new Drawing.Point(20, 160);
                lblLocation.Size = new Drawing.Size(100, 20);
                this.Controls.Add(lblLocation);

                _cmbLocation = new WinForms.ComboBox();
                _cmbLocation.Location = new Drawing.Point(130, 160);
                _cmbLocation.Size = new Drawing.Size(280, 25);
                _cmbLocation.Font = new Drawing.Font("Arial", 9);
                _cmbLocation.DropDownStyle = WinForms.ComboBoxStyle.DropDownList;
                _cmbLocation.Items.Add("Wall Centerline (Mac dinh)");
                _cmbLocation.Items.Add("Core Face: Interior");
                _cmbLocation.Items.Add("Core Face: Exterior");
                _cmbLocation.SelectedIndex = 0;
                this.Controls.Add(_cmbLocation);

                // LOCATION INFO
                WinForms.Label lblLocationInfo = new WinForms.Label();
                lblLocationInfo.Text = "SU DUNG LOCATION LINE DE TRANH CHONG LAN TUONG";
                lblLocationInfo.Font = new Drawing.Font("Arial", 8, Drawing.FontStyle.Bold);
                lblLocationInfo.ForeColor = Drawing.Color.DarkGreen;
                lblLocationInfo.Location = new Drawing.Point(130, 185);
                lblLocationInfo.Size = new Drawing.Size(350, 30);
                this.Controls.Add(lblLocationInfo);

                // BUTTONS
                WinForms.Button btnOk = new WinForms.Button();
                btnOk.Text = "TAO TUONG";
                btnOk.Location = new Drawing.Point(120, 230);
                btnOk.Size = new Drawing.Size(130, 35);
                btnOk.BackColor = Drawing.Color.LightBlue;
                btnOk.Font = new Drawing.Font("Arial", 10, Drawing.FontStyle.Bold);
                btnOk.Click += BtnOK_Click;
                this.Controls.Add(btnOk);

                WinForms.Button btnCancel = new WinForms.Button();
                btnCancel.Text = "HUY";
                btnCancel.Location = new Drawing.Point(260, 230);
                btnCancel.Size = new Drawing.Size(90, 35);
                btnCancel.Font = new Drawing.Font("Arial", 10);
                btnCancel.Click += BtnCancel_Click;
                this.Controls.Add(btnCancel);
            }

            private void BtnOK_Click(object sender, EventArgs e)
            {
                if (_cmbWallType.SelectedIndex < 0)
                {
                    WinForms.MessageBox.Show("Vui long chon loai tuong!", "Canh bao");
                    return;
                }

                try
                {
                    if (!double.TryParse(_txtHeight.Text, out double heightValue) || heightValue <= 0)
                    {
                        WinForms.MessageBox.Show("Chieu cao phai > 0!", "Canh bao");
                        return;
                    }

                    HeightMm = heightValue;
                    WallLocationLine = _cmbLocation.SelectedIndex + 1;

                    string displayName = _cmbWallType.SelectedItem.ToString();
                    string wallTypeName = displayName.Split('(')[0].Trim();

                    if (_wallTypesDict.ContainsKey(wallTypeName))
                    {
                        SelectedWallType = _wallTypesDict[wallTypeName];
                        this.DialogResult = WinForms.DialogResult.OK;
                        this.Close();
                    }
                    else
                    {
                        WinForms.MessageBox.Show("Khong tim thay wall type: " + wallTypeName, "Loi");
                    }
                }
                catch (FormatException)
                {
                    WinForms.MessageBox.Show("Chieu cao phai la so!", "Canh bao");
                }
                catch (Exception ex)
                {
                    WinForms.MessageBox.Show("Loi: " + ex.Message, "Loi");
                }
            }

            private void BtnCancel_Click(object sender, EventArgs e)
            {
                this.DialogResult = WinForms.DialogResult.Cancel;
                this.Close();
            }

            private double GetWallTypeThickness(WallType wallType)
            {
                try
                {
                    Parameter param = wallType.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM);
                    if (param != null && param.HasValue && param.AsDouble() > 0)
                    {
                        return UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters);
                    }

                    param = wallType.LookupParameter("Width");
                    if (param != null && param.HasValue && param.AsDouble() > 0)
                    {
                        return UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters);
                    }

                    CompoundStructure compoundStructure = wallType.GetCompoundStructure();
                    if (compoundStructure != null)
                    {
                        return UnitUtils.ConvertFromInternalUnits(
                            compoundStructure.GetWidth(),
                            UnitTypeId.Millimeters);
                    }

                    return 0;
                }
                catch
                {
                    return 0;
                }
            }
        }

        // ROOM BOUNDARY EXTRACTION
        private List<Curve> GetRoomBoundaryCurves(Room room)
        {
            List<Curve> curves = new List<Curve>();

            try
            {
                SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions();
                options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish;
                IList<IList<BoundarySegment>> boundarySegments = room.GetBoundarySegments(options);

                if (boundarySegments == null || boundarySegments.Count == 0)
                    return null;

                foreach (IList<BoundarySegment> segmentLoop in boundarySegments)
                {
                    foreach (BoundarySegment boundarySegment in segmentLoop)
                    {
                        try
                        {
                            Curve curve = boundarySegment.GetCurve();
                            if (curve != null && curve.Length >= _minCurveLength)
                            {
                                curves.Add(curve);
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }

                return curves.Count > 0 ? curves : null;
            }
            catch (Exception)
            {
                return null;
            }
        }

        // WALL CREATION WITH LOCATION LINE
        private List<Wall> CreateWallsFromCurves(
            List<Curve> curves,
            WallType wallType,
            Level level,
            double heightInternal,
            int wallLocationLine,
            Room room)
        {
            List<Wall> walls = new List<Wall>();

            try
            {
                double wallThickness = GetWallTypeThickness(wallType);
                double wallThicknessInternal = UnitUtils.ConvertToInternalUnits(wallThickness, UnitTypeId.Millimeters);

                for (int idx = 0; idx < curves.Count; idx++)
                {
                    try
                    {
                        Curve curve = curves[idx];
                        Curve offsetCurve = CalculateOffsetCurve(curve, wallLocationLine, wallThicknessInternal, room);

                        Wall wall = Wall.Create(
                            wallType.Document,
                            offsetCurve,
                            wallType.Id,
                            level.Id,
                            heightInternal,
                            0.0,
                            false,
                            false);

                        if (wall != null)
                        {
                            walls.Add(wall);
                            double curveLen = UnitUtils.ConvertFromInternalUnits(curve.Length, UnitTypeId.Meters);
                            string locationName = GetLocationLineName(wallLocationLine);
                        }
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }

                return walls;
            }
            catch (Exception)
            {
                return walls;
            }
        }

        private Curve CalculateOffsetCurve(Curve originalCurve, int locationLine, double wallThickness, Room room)
        {
            try
            {
                if (!(originalCurve is Line))
                    return originalCurve;

                XYZ startPoint = originalCurve.GetEndPoint(0);
                XYZ endPoint = originalCurve.GetEndPoint(1);

                XYZ direction = (endPoint - startPoint).Normalize();
                XYZ perpendicular = new XYZ(-direction.Y, direction.X, 0).Normalize();

                double offsetDirection = GetOffsetDirection(originalCurve, room, perpendicular);
                double offsetDistance = 0.0;

                if (locationLine == 1)
                {
                    offsetDistance = 0.0;
                }
                else if (locationLine == 2)
                {
                    offsetDistance = wallThickness / 2.0;
                }
                else if (locationLine == 3)
                {
                    offsetDistance = -wallThickness / 2.0;
                }

                XYZ offsetVector = perpendicular * offsetDirection * offsetDistance;
                XYZ newStart = startPoint + offsetVector;
                XYZ newEnd = endPoint + offsetVector;

                return Line.CreateBound(newStart, newEnd);
            }
            catch (Exception)
            {
                return originalCurve;
            }
        }

        private double GetOffsetDirection(Curve curve, Room room, XYZ perpendicularVector)
        {
            try
            {
                XYZ midPoint = curve.Evaluate(0.5, true);
                XYZ roomPoint;

                Location roomLocation = room.Location;
                if (roomLocation is LocationPoint locationPoint)
                {
                    roomPoint = locationPoint.Point;
                }
                else
                {
                    BoundingBoxXYZ roomBb = room.get_BoundingBox(null);
                    if (roomBb != null)
                    {
                        roomPoint = (roomBb.Max + roomBb.Min) / 2.0;
                    }
                    else
                    {
                        return 1.0;
                    }
                }

                XYZ toRoom = roomPoint - midPoint;
                double dotProduct = perpendicularVector.DotProduct(toRoom);

                return dotProduct > 0 ? -1.0 : 1.0;
            }
            catch (Exception)
            {
                return 1.0;
            }
        }

        private string GetLocationLineName(int locationLine)
        {
            Dictionary<int, string> names = new Dictionary<int, string>
            {
                {1, "Wall Centerline"},
                {2, "Core Face: Interior"},
                {3, "Core Face: Exterior"}
            };

            return names.ContainsKey(locationLine) ? names[locationLine] : "Unknown";
        }

        // MAIN EXECUTION
        private void Run(UIDocument uidoc, Document doc)
        {
            try
            {
                ShowMessage(
                    "Vui long chon NHIEU Rooms (Ctrl+Click) de tao tuong\n\n" +
                    "CHU Y QUAN TRONG:\n" +
                    "• Script nay tao tuong chinh xac theo Location Line\n" +
                    "• Cac tuong se duoc tao LIEN KE voi tuong cu\n" +
                    "• DE TRANH CHONG LAN: Su dung Core Face: Interior/Exterior\n" +
                    "• Sau khi tao, co the can Join Geometry thu cong trong Revit",
                    "Huong dan");

                // CHON NHIEU ROOMS
                List<Room> rooms = PickRooms(uidoc);
                if (rooms == null || rooms.Count == 0)
                    return;

                ShowMessage(
                    string.Format("Da chon {0} rooms\n\nTiep theo: Chon loai tuong va nhap tham so", rooms.Count),
                    "Thong bao");

                // GET WALL TYPES
                Dictionary<string, WallType> wallTypesDict = GetAllWallTypes(doc);
                if (wallTypesDict == null)
                    return;

                // SHOW FORM
                using (RoomWallForm form = new RoomWallForm(wallTypesDict))
                {
                    if (form.ShowDialog() != WinForms.DialogResult.OK)
                    {
                        ShowMessage("Ban da huy tao tuong.", "Thong bao");
                        return;
                    }

                    WallType wallType = form.SelectedWallType;
                    double heightMm = form.HeightMm;
                    int wallLocationLine = form.WallLocationLine;

                    double heightInternal = UnitUtils.ConvertToInternalUnits(heightMm, UnitTypeId.Millimeters);

                    // PROCESS NHIEU ROOMS
                    ShowMessage(
                        string.Format("Dang tao walls cho {0} rooms...\n\nVui long cho...", rooms.Count),
                        "Thong bao");

                    using (Transaction t = new Transaction(doc, string.Format("Create Walls from {0} Rooms (MULTI)", rooms.Count)))
                    {
                        t.Start();

                        try
                        {
                            int totalWalls = 0;
                            int successRooms = 0;
                            int failedRooms = 0;

                            foreach (Room room in rooms)
                            {
                                try
                                {
                                    string roomNumber = room.Number;
                                    Level level = room.Level;

                                    if (level == null)
                                    {
                                        failedRooms++;
                                        continue;
                                    }

                                    // LAY BOUNDARY CURVES
                                    List<Curve> curves = GetRoomBoundaryCurves(room);
                                    if (curves == null || curves.Count == 0)
                                    {
                                        failedRooms++;
                                        continue;
                                    }

                                    // TAO WALLS
                                    List<Wall> walls = CreateWallsFromCurves(
                                        curves,
                                        wallType,
                                        level,
                                        heightInternal,
                                        wallLocationLine,
                                        room);

                                    if (walls != null && walls.Count > 0)
                                    {
                                        totalWalls += walls.Count;
                                        successRooms++;
                                    }
                                    else
                                    {
                                        failedRooms++;
                                    }
                                }
                                catch (Exception)
                                {
                                    failedRooms++;
                                    continue;
                                }
                            }

                            t.Commit();

                            // RESULT
                            Dictionary<int, string> locationNames = new Dictionary<int, string>
                            {
                                {1, "Wall Centerline"},
                                {2, "Core Face: Interior"},
                                {3, "Core Face: Exterior"}
                            };

                            string msg = "========== HOAN THANH ==========\n\n";
                            msg += string.Format("Tong rooms da chon: {0}\n", rooms.Count);
                            msg += string.Format("Rooms tao thanh cong: {0}\n", successRooms);
                            msg += string.Format("Rooms loi: {0}\n\n", failedRooms);
                            msg += string.Format("Wall Type: {0} ({1:F0}mm)\n",
                                GetWallTypeName(wallType),
                                GetWallTypeThickness(wallType));
                            msg += string.Format("Height: {0:F0}mm\n", heightMm);
                            msg += string.Format("Location Line: {0}\n\n",
                                locationNames.ContainsKey(wallLocationLine) ? locationNames[wallLocationLine] : "Unknown");
                            msg += string.Format("TONG WALLS TAO: {0} walls\n\n", totalWalls);
                            msg += "KHUYEN NGHI:\n";
                            msg += "• Su dung Core Face: Interior/Exterior de tranh chong lan\n";
                            msg += "• Co the can Join Geometry thu cong trong Revit\n";
                            msg += "• Kiem tra vi tri tuong da chinh xac chua";

                            ShowMessage(msg, "Hoan thanh");
                        }
                        catch (Exception ex)
                        {
                            t.RollBack();
                            ShowError("Loi tao walls: " + ex.Message, "Loi");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError("Loi: " + ex.Message, "Loi");
            }
        }
    }
}