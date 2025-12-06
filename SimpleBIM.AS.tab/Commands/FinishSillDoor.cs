using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.Attributes;

//===================== HOÀN CHỈNH =====================

namespace SimpleBIM.AS.tab.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class FinishSillDoor : IExternalCommand
    {
        private Document doc;
        private UIDocument uidoc;

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
                // Khởi tạo document
                uidoc = commandData.Application.ActiveUIDocument;
                doc = uidoc.Document;

                // Gọi hàm main
                return Run() ? Result.Succeeded : Result.Failed;
            }
            catch (Exception ex)
            {
                message = $"{ex.Message}\n{ex.StackTrace}";
                return Result.Failed;
            }
        }

        // ==================== HÀM HIỂN THỊ THÔNG BÁO ====================
        private void ShowMessage(string message, string title = "Thong bao")
        {
            System.Windows.Forms.MessageBox.Show(message, title,
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }

        private void ShowError(string message, string title = "Loi")
        {
            System.Windows.Forms.MessageBox.Show(message, title,
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Error);
        }

        private bool ShowYesNo(string message, string title = "Xac nhan")
        {
            var result = System.Windows.Forms.MessageBox.Show(message, title,
                System.Windows.Forms.MessageBoxButtons.YesNo,
                System.Windows.Forms.MessageBoxIcon.Question);
            return result == System.Windows.Forms.DialogResult.Yes;
        }

        // ==================== DOOR SELECTION FILTER ====================
        private class DoorSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                try
                {
                    if (elem is FamilyInstance)
                    {
                        if (elem.Category != null &&
                            elem.Category.Id.Value == (long)BuiltInCategory.OST_Doors)
                        {
                            return true;
                        }
                    }
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

        // ==================== PICK MULTIPLE DOORS ====================
        private List<Element> PickMultipleDoors()
        {
            List<Element> selectedDoors = new List<Element>();
            HashSet<long> doorIds = new HashSet<long>();

            try
            {
                while (true)
                {
                    try
                    {
                        string promptMsg = $"Chon cua trong Floor Plans (Da chon: {selectedDoors.Count} cua) - Nhan ESC de dung";

                        Reference refElem = uidoc.Selection.PickObject(
                            ObjectType.Element,
                            new DoorSelectionFilter(),
                            promptMsg);

                        Element element = doc.GetElement(refElem.ElementId);

                        if (element != null)
                        {
                            // Check if door already selected
                            if (doorIds.Contains(element.Id.Value))  // ← SỬA: IntegerValue → Value
                            {
                                continue;
                            }

                            // Add to list
                            selectedDoors.Add(element);
                            doorIds.Add(element.Id.Value);  // ← SỬA: IntegerValue → Value

                            string doorName = element.Name ?? $"Door_{element.Id.Value}";  // ← SỬA: IntegerValue → Value
                        }
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                    {
                        // User pressed ESC
                        break;
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = ex.ToString();
                        if (errorMsg.ToLower().Contains("canceled") ||
                            errorMsg.ToLower().Contains("aborted"))
                        {
                            break;
                        }
                        else
                        {
                            ShowError("Loi: " + errorMsg, "Loi");
                            break;
                        }
                    }
                }

                if (selectedDoors.Count == 0)
                {
                    ShowMessage("Khong chon cua nao.", "Thong bao");
                    return null;
                }

                // Confirm selection
                string confirmMsg = $"Da chon {selectedDoors.Count} cua. Ban co muon tao san cho cac cua nay khong?";
                if (ShowYesNo(confirmMsg, "Xac nhan"))
                {
                    return selectedDoors;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                ShowError("Loi khi chon cua: " + ex.ToString(), "Loi");
                return null;
            }
        }

        // ==================== DOOR SILL CALCULATION ====================
        private Tuple<double, double> GetDoorSillDimensions(FamilyInstance door)
        {
            try
            {
                // Get door width
                double doorWidthInternal = 0;

                // Method 1: LookupParameter "Width"
                Parameter doorWidthParam = door.LookupParameter("Width");
                if (doorWidthParam != null && doorWidthParam.AsDouble() > 0)
                {
                    doorWidthInternal = doorWidthParam.AsDouble();
                }

                // Method 2: Get from Symbol
                if (doorWidthInternal == 0 && door.Symbol != null)
                {
                    Parameter widthParam = door.Symbol.LookupParameter("Width");
                    if (widthParam != null && widthParam.AsDouble() > 0)
                    {
                        doorWidthInternal = widthParam.AsDouble();
                    }
                }

                // Method 3: Get from BoundingBox
                if (doorWidthInternal == 0)
                {
                    BoundingBoxXYZ bb = door.get_BoundingBox(null);
                    if (bb != null)
                    {
                        double width = bb.Max.X - bb.Min.X;
                        if (width > 0)
                        {
                            doorWidthInternal = width;
                        }
                    }
                }

                if (doorWidthInternal == 0)
                {
                    ShowError($"Khong tim thay kich thuoc cua (Width) cho: {door.Name}", "Loi");
                    return null;
                }

                double doorWidthMm = UnitUtils.ConvertFromInternalUnits(doorWidthInternal, UnitTypeId.Millimeters);

                // Get wall
                Element hostElement = door.Host;
                if (hostElement == null || !(hostElement is Wall))
                {
                    ShowError($"Cua {door.Name} khong nam tren tuong!", "Loi");
                    return null;
                }

                Wall wall = hostElement as Wall;
                double wallThicknessInternal = 0;

                // Method 1: Get from wall parameter "Width"
                Parameter widthParamWall = wall.LookupParameter("Width");
                if (widthParamWall != null && widthParamWall.AsDouble() > 0)
                {
                    wallThicknessInternal = widthParamWall.AsDouble();
                }

                // Method 2: Get from wall parameter "Thickness"
                if (wallThicknessInternal == 0)
                {
                    Parameter thicknessParam = wall.LookupParameter("Thickness");
                    if (thicknessParam != null && thicknessParam.AsDouble() > 0)
                    {
                        wallThicknessInternal = thicknessParam.AsDouble();
                    }
                }

                // Method 3: Get from BoundingBox
                if (wallThicknessInternal == 0)
                {
                    BoundingBoxXYZ bb = wall.get_BoundingBox(null);
                    if (bb != null)
                    {
                        double thicknessX = bb.Max.X - bb.Min.X;
                        double thicknessY = bb.Max.Y - bb.Min.Y;
                        if (thicknessX > 0 && thicknessY > 0)
                        {
                            wallThicknessInternal = Math.Min(thicknessX, thicknessY);
                        }
                    }
                }

                // Method 4: Get from WallType (không dùng GetOverallThickness)
                if (wallThicknessInternal == 0)
                {
                    WallType wallType = wall.WallType;
                    if (wallType != null)
                    {
                        // Lấy thickness từ parameter của WallType
                        Parameter typeThicknessParam = wallType.LookupParameter("Thickness");
                        if (typeThicknessParam != null && typeThicknessParam.AsDouble() > 0)
                        {
                            wallThicknessInternal = typeThicknessParam.AsDouble();
                        }
                    }
                }

                // Method 5: User input (not implemented in C# version - using default value)
                if (wallThicknessInternal == 0)
                {
                    ShowError($"Khong the lay chieu rong cua tuong cho cua {door.Name}!", "Loi");
                    return null;
                }

                double wallThicknessMm = UnitUtils.ConvertFromInternalUnits(wallThicknessInternal, UnitTypeId.Millimeters);

                return new Tuple<double, double>(doorWidthMm, wallThicknessMm);
            }
            catch (Exception ex)
            {
                ShowError("Loi tinh toan kich thuoc gach cua: " + ex.ToString(), "Loi");
                return null;
            }
        }

        private Tuple<XYZ, XYZ, double> GetDoorLocationAndDirection(FamilyInstance door)
        {
            try
            {
                Location location = door.Location;
                if (location == null || !(location is LocationPoint))
                {
                    ShowError("Khong the lay Location Point cua cua!", "Loi");
                    return null;
                }

                LocationPoint locationPoint = location as LocationPoint;
                XYZ doorPoint = locationPoint.Point;
                double rotation = locationPoint.Rotation;

                double directionX = Math.Cos(rotation);
                double directionY = Math.Sin(rotation);
                XYZ direction = new XYZ(directionX, directionY, 0).Normalize();

                return new Tuple<XYZ, XYZ, double>(doorPoint, direction, rotation);
            }
            catch (Exception ex)
            {
                ShowError("Loi lay location va direction cua cua: " + ex.ToString(), "Loi");
                return null;
            }
        }

        private List<Line> CreateDoorSillRectangle(XYZ doorPoint, XYZ direction, double doorWidthInternal, double wallThicknessInternal)
        {
            try
            {
                XYZ perpendicular = new XYZ(-direction.Y, direction.X, 0).Normalize();

                double halfWidth = doorWidthInternal / 2.0;
                double halfThickness = wallThicknessInternal / 2.0;

                XYZ p1 = doorPoint - direction * halfWidth - perpendicular * halfThickness;
                XYZ p2 = doorPoint + direction * halfWidth - perpendicular * halfThickness;
                XYZ p3 = doorPoint + direction * halfWidth + perpendicular * halfThickness;
                XYZ p4 = doorPoint - direction * halfWidth + perpendicular * halfThickness;

                Line line1 = Line.CreateBound(p1, p2);
                Line line2 = Line.CreateBound(p2, p3);
                Line line3 = Line.CreateBound(p3, p4);
                Line line4 = Line.CreateBound(p4, p1);

                return new List<Line> { line1, line2, line3, line4 };
            }
            catch (Exception ex)
            {
                ShowError("Loi tao rectangle: " + ex.ToString(), "Loi");
                return null;
            }
        }

        private CurveLoop CreateCurveLoopFromLines(List<Line> lines)
        {
            try
            {
                CurveLoop curveLoop = new CurveLoop();

                foreach (Line line in lines)
                {
                    curveLoop.Append(line);
                }

                if (!curveLoop.IsOpen())
                {
                    return curveLoop;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                ShowError("Loi tao CurveLoop: " + ex.ToString(), "Loi");
                return null;
            }
        }

        // ==================== FLOOR TYPE FUNCTIONS ====================
        private string GetFloorTypeName(Element floorTypeElement)
        {
            try
            {
                if (!string.IsNullOrEmpty(floorTypeElement.Name))
                    return floorTypeElement.Name;

                Parameter param = floorTypeElement.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM);
                if (param != null && !string.IsNullOrEmpty(param.AsString()))
                    return param.AsString();

                return $"FloorType_{floorTypeElement.Id.Value}";
            }
            catch
            {
                return "Unknown";
            }
        }

        private double GetFloorTypeThickness(FloorType floorType)
        {
            try
            {
                // Thử lấy từ parameter "Thickness"
                Parameter param = floorType.LookupParameter("Thickness");
                if (param != null && param.AsDouble() > 0)
                {
                    double thickness = UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters);
                    return thickness;
                }

                // Thử lấy từ BuiltInParameter
                param = floorType.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM);
                if (param != null && param.AsDouble() > 0)
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

        private bool IsFloorTypeStructural(FloorType floorType)
        {
            try
            {
                Parameter param = floorType.LookupParameter("Structural");
                if (param != null)
                {
                    return param.AsInteger() == 1;
                }

                string typeName = GetFloorTypeName(floorType).ToLower();
                return typeName.Contains("structural");
            }
            catch
            {
                return false;
            }
        }

        private Dictionary<string, FloorType> GetAllFloorTypes()
        {
            Dictionary<string, FloorType> floorTypes = new Dictionary<string, FloorType>();

            try
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(FloorType));

                foreach (Element elem in collector)
                {
                    FloorType floorType = elem as FloorType;
                    if (floorType != null)
                    {
                        if (!IsFloorTypeStructural(floorType))
                        {
                            string typeName = GetFloorTypeName(floorType);
                            if (!string.IsNullOrEmpty(typeName) &&
                                typeName != "Unknown" &&
                                !floorTypes.ContainsKey(typeName))
                            {
                                floorTypes.Add(typeName, floorType);
                            }
                        }
                    }
                }

                if (floorTypes.Count == 0)
                {
                    ShowError("Khong tim thay Floor Type Non-structural nao trong project!", "Loi");
                    return null;
                }

                return floorTypes;
            }
            catch (Exception ex)
            {
                ShowError("Loi lay floor types: " + ex.ToString(), "Loi");
                return null;
            }
        }

        // ==================== FORM CLASS ====================
        private class MultiDoorSillFloorForm : System.Windows.Forms.Form
        {
            private Dictionary<string, FloorType> floorTypesDict;
            private List<Dictionary<string, object>> doorsInfo;
            private System.Windows.Forms.ListBox lstDoors;
            private System.Windows.Forms.ComboBox cmbFloorType;
            private System.Windows.Forms.TextBox txtOffset;

            public FloorType SelectedFloorType { get; private set; }
            public double BaseOffsetMm { get; private set; }

            public MultiDoorSillFloorForm(
                Dictionary<string, FloorType> floorTypesDict,
                List<Dictionary<string, object>> doorsInfo)
            {
                this.floorTypesDict = floorTypesDict;
                this.doorsInfo = doorsInfo;
                InitializeComponent();
            }

            private void InitializeComponent()
            {
                this.Text = "Tao San Gach Cua - Nhieu Cua";
                this.Width = 600;
                this.Height = 500;
                this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                this.BackColor = System.Drawing.Color.White;
                this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;

                // TITLE
                System.Windows.Forms.Label lblTitle = new System.Windows.Forms.Label();
                lblTitle.Text = "TAO SAN GACH CUA - NHIEU CUA";
                lblTitle.Font = new System.Drawing.Font("Arial", 12, System.Drawing.FontStyle.Bold);
                lblTitle.ForeColor = System.Drawing.Color.DarkBlue;
                lblTitle.Location = new System.Drawing.Point(20, 20);
                lblTitle.Size = new System.Drawing.Size(560, 25);
                this.Controls.Add(lblTitle);

                // DOOR LIST
                System.Windows.Forms.Label lblDoors = new System.Windows.Forms.Label();
                lblDoors.Text = $"DANH SACH CUA DA CHON ({doorsInfo.Count} cua)";
                lblDoors.Font = new System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold);
                lblDoors.ForeColor = System.Drawing.Color.DarkGreen;
                lblDoors.Location = new System.Drawing.Point(20, 55);
                lblDoors.Size = new System.Drawing.Size(560, 20);
                this.Controls.Add(lblDoors);

                // LISTBOX
                lstDoors = new System.Windows.Forms.ListBox();
                lstDoors.Location = new System.Drawing.Point(20, 80);
                lstDoors.Size = new System.Drawing.Size(560, 150);
                lstDoors.Font = new System.Drawing.Font("Arial", 9);

                foreach (var doorInfo in doorsInfo)
                {
                    string itemText = $"{doorInfo["name"]} - {(int)(double)doorInfo["door_width_mm"]}x{(int)(double)doorInfo["wall_thickness_mm"]} mm";
                    lstDoors.Items.Add(itemText);
                }

                this.Controls.Add(lstDoors);

                // FLOOR TYPE
                System.Windows.Forms.Label lblFloorType = new System.Windows.Forms.Label();
                lblFloorType.Text = "Loai San:";
                lblFloorType.Font = new System.Drawing.Font("Arial", 10);
                lblFloorType.Location = new System.Drawing.Point(20, 245);
                lblFloorType.Size = new System.Drawing.Size(100, 20);
                this.Controls.Add(lblFloorType);

                cmbFloorType = new System.Windows.Forms.ComboBox();
                cmbFloorType.Location = new System.Drawing.Point(130, 245);
                cmbFloorType.Size = new System.Drawing.Size(450, 25);
                cmbFloorType.Font = new System.Drawing.Font("Arial", 9);
                cmbFloorType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;

                List<string> floorTypeNames = floorTypesDict.Keys.ToList();
                floorTypeNames.Sort();

                foreach (string name in floorTypeNames)
                {
                    FloorType floorType = floorTypesDict[name];
                    double thickness = GetFloorTypeThickness(floorType);
                    string displayName = thickness > 0 ? $"{name} ({thickness:F0}mm)" : name;
                    cmbFloorType.Items.Add(displayName);
                }

                if (floorTypeNames.Count > 0)
                {
                    cmbFloorType.SelectedIndex = 0;
                }

                this.Controls.Add(cmbFloorType);

                // BASE OFFSET
                System.Windows.Forms.Label lblOffset = new System.Windows.Forms.Label();
                lblOffset.Text = "Base Offset (mm):";
                lblOffset.Font = new System.Drawing.Font("Arial", 10);
                lblOffset.Location = new System.Drawing.Point(20, 285);
                lblOffset.Size = new System.Drawing.Size(100, 20);
                this.Controls.Add(lblOffset);

                txtOffset = new System.Windows.Forms.TextBox();
                txtOffset.Location = new System.Drawing.Point(130, 285);
                txtOffset.Size = new System.Drawing.Size(450, 25);
                txtOffset.Font = new System.Drawing.Font("Arial", 9);
                txtOffset.Text = "0";
                this.Controls.Add(txtOffset);

                // BUTTONS
                System.Windows.Forms.Button btnOk = new System.Windows.Forms.Button();
                btnOk.Text = $"TAO {doorsInfo.Count} SAN";
                btnOk.Location = new System.Drawing.Point(180, 330);
                btnOk.Size = new System.Drawing.Size(150, 35);
                btnOk.BackColor = System.Drawing.Color.LightBlue;
                btnOk.Font = new System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold);
                btnOk.Click += BtnOK_Click;
                this.Controls.Add(btnOk);

                System.Windows.Forms.Button btnCancel = new System.Windows.Forms.Button();
                btnCancel.Text = "HUY";
                btnCancel.Location = new System.Drawing.Point(340, 330);
                btnCancel.Size = new System.Drawing.Size(90, 35);
                btnCancel.Font = new System.Drawing.Font("Arial", 10);
                btnCancel.Click += BtnCancel_Click;
                this.Controls.Add(btnCancel);
            }

            private double GetFloorTypeThickness(FloorType floorType)
            {
                try
                {
                    // Thử lấy từ parameter "Thickness"
                    Parameter param = floorType.LookupParameter("Thickness");
                    if (param != null && param.AsDouble() > 0)
                    {
                        return UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters);
                    }

                    // Thử lấy từ BuiltInParameter
                    param = floorType.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM);
                    if (param != null && param.AsDouble() > 0)
                    {
                        return UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters);
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
                if (cmbFloorType.SelectedIndex < 0)
                {
                    System.Windows.Forms.MessageBox.Show("Vui long chon loai san!", "Canh bao");
                    return;
                }

                try
                {
                    if (!double.TryParse(txtOffset.Text, out double offsetValue))
                    {
                        System.Windows.Forms.MessageBox.Show("Base Offset phai la so!", "Canh bao");
                        return;
                    }

                    BaseOffsetMm = offsetValue;

                    string displayName = cmbFloorType.SelectedItem.ToString();
                    string floorTypeName = displayName.Contains(" (")
                        ? displayName.Substring(0, displayName.IndexOf(" ("))
                        : displayName;

                    if (floorTypesDict.ContainsKey(floorTypeName))
                    {
                        SelectedFloorType = floorTypesDict[floorTypeName];
                        this.DialogResult = System.Windows.Forms.DialogResult.OK;
                        this.Close();
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("Khong tim thay floor type: " + floorTypeName, "Loi");
                    }
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show("Loi: " + ex.Message, "Loi");
                }
            }

            private void BtnCancel_Click(object sender, EventArgs e)
            {
                this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
                this.Close();
            }
        }

        // ==================== CREATE FLOOR FUNCTION ====================
        private Floor CreateFloorFromCurveLoop(CurveLoop curveLoop, FloorType floorType, Level level, double baseOffsetMm)
        {
            try
            {
                double baseOffsetInternal = UnitUtils.ConvertToInternalUnits(baseOffsetMm, UnitTypeId.Millimeters);

                IList<CurveLoop> curveLoops = new List<CurveLoop> { curveLoop };
                Floor floor = Floor.Create(doc, curveLoops, floorType.Id, level.Id);

                if (floor != null)
                {
                    Parameter param = floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
                    if (param != null && !param.IsReadOnly)
                    {
                        param.Set(baseOffsetInternal);
                    }

                    return floor;
                }

                return null;
            }
            catch (Exception ex)
            {
                ShowError("Loi tao floor: " + ex.ToString(), "Loi");
                return null;
            }
        }

        // ==================== GET LEVEL FROM WALL ====================
        private Level GetLevelFromWall(Wall wall)
        {
            try
            {
                // Cách 1: Lấy từ parameter Level
                Parameter levelParam = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT);
                if (levelParam != null && levelParam.HasValue)
                {
                    ElementId levelId = levelParam.AsElementId();
                    if (levelId != ElementId.InvalidElementId)
                    {
                        return doc.GetElement(levelId) as Level;
                    }
                }

                // Cách 2: Lấy từ Level của host nếu có
                if (wall.LevelId != ElementId.InvalidElementId)
                {
                    return doc.GetElement(wall.LevelId) as Level;
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        // ==================== MAIN EXECUTION ====================
        private bool Run()
        {
            try
            {
                ShowMessage(
                    "TAO SAN GACH CUA - NHIEU CUA\n\n" +
                    "HUONG DAN:\n" +
                    "1. Click OK de bat dau chon cua\n" +
                    "2. Click tung cua trong Floor Plans\n" +
                    "3. Chon 'Yes' de tiep tuc, 'No' de hoan tat chon\n" +
                    "4. Chon loai san va Base Offset\n" +
                    "5. Tao san cho tat ca cua da chon",
                    "Huong dan");

                // PICK MULTIPLE DOORS
                List<Element> doors = PickMultipleDoors();
                if (doors == null || doors.Count == 0)
                {
                    return false;
                }

                ShowMessage($"Da chon {doors.Count} cua\n\nDang tinh toan...", "Thong bao");

                // GET FLOOR TYPES
                Dictionary<string, FloorType> floorTypesDict = GetAllFloorTypes();
                if (floorTypesDict == null)
                {
                    return false;
                }

                // PROCESS EACH DOOR
                List<Dictionary<string, object>> doorsInfo = new List<Dictionary<string, object>>();
                List<Element> validDoors = new List<Element>();

                foreach (Element doorElem in doors)
                {
                    FamilyInstance door = doorElem as FamilyInstance;
                    if (door == null) continue;

                    try
                    {
                        string doorName = door.Name ?? $"Door_{door.Id.Value}";
                        var dimensions = GetDoorSillDimensions(door);

                        if (dimensions != null)
                        {
                            doorsInfo.Add(new Dictionary<string, object>
                            {
                                { "door", door },
                                { "name", doorName },
                                { "door_width_mm", dimensions.Item1 },
                                { "wall_thickness_mm", dimensions.Item2 }
                            });
                            validDoors.Add(door);
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

                if (doorsInfo.Count == 0)
                {
                    ShowError("Khong the xu ly cua nao!", "Loi");
                    return false;
                }

                // SHOW FORM
                using (MultiDoorSillFloorForm form = new MultiDoorSillFloorForm(floorTypesDict, doorsInfo))
                {
                    if (form.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                    {
                        ShowMessage("Ban da huy tao san.", "Thong bao");
                        return false;
                    }

                    FloorType floorType = form.SelectedFloorType;
                    double baseOffsetMm = form.BaseOffsetMm;

                    // CREATE FLOORS
                    ShowMessage($"Dang tao san cho {doorsInfo.Count} cua...", "Thong bao");

                    using (Transaction trans = new Transaction(doc, "Create Floors from Door Sills"))
                    {
                        trans.Start();

                        try
                        {
                            int createdCount = 0;
                            List<string> failedDoors = new List<string>();

                            foreach (var doorInfo in doorsInfo)
                            {
                                try
                                {
                                    FamilyInstance door = doorInfo["door"] as FamilyInstance;
                                    double doorWidthMm = (double)doorInfo["door_width_mm"];
                                    double wallThicknessMm = (double)doorInfo["wall_thickness_mm"];

                                    // Get location and direction
                                    var locationInfo = GetDoorLocationAndDirection(door);
                                    if (locationInfo == null)
                                    {
                                        failedDoors.Add(doorInfo["name"] as string);
                                        continue;
                                    }

                                    XYZ doorPoint = locationInfo.Item1;
                                    XYZ direction = locationInfo.Item2;

                                    // Get level
                                    Level level = null;
                                    Wall hostWall = door.Host as Wall;
                                    if (hostWall != null)
                                    {
                                        level = GetLevelFromWall(hostWall);
                                    }

                                    if (level == null)
                                    {
                                        Parameter levelParam = door.LookupParameter("Level");
                                        if (levelParam != null)
                                        {
                                            ElementId levelId = levelParam.AsElementId();
                                            if (levelId != ElementId.InvalidElementId)
                                            {
                                                level = doc.GetElement(levelId) as Level;
                                            }
                                        }
                                    }

                                    if (level == null)
                                    {
                                        failedDoors.Add(doorInfo["name"] as string);
                                        continue;
                                    }

                                    // Convert to internal units
                                    double doorWidthInternal = UnitUtils.ConvertToInternalUnits(doorWidthMm, UnitTypeId.Millimeters);
                                    double wallThicknessInternal = UnitUtils.ConvertToInternalUnits(wallThicknessMm, UnitTypeId.Millimeters);

                                    // Create rectangle
                                    List<Line> lines = CreateDoorSillRectangle(
                                        doorPoint, direction, doorWidthInternal, wallThicknessInternal);

                                    if (lines == null)
                                    {
                                        failedDoors.Add(doorInfo["name"] as string);
                                        continue;
                                    }

                                    // Create curve loop
                                    CurveLoop curveLoop = CreateCurveLoopFromLines(lines);
                                    if (curveLoop == null)
                                    {
                                        failedDoors.Add(doorInfo["name"] as string);
                                        continue;
                                    }

                                    // Create floor
                                    Floor floor = CreateFloorFromCurveLoop(curveLoop, floorType, level, baseOffsetMm);
                                    if (floor != null)
                                    {
                                        createdCount++;
                                    }
                                    else
                                    {
                                        failedDoors.Add(doorInfo["name"] as string);
                                    }
                                }
                                catch
                                {
                                    failedDoors.Add(doorInfo["name"] as string);
                                    continue;
                                }
                            }

                            trans.Commit();

                            // RESULT MESSAGE
                            StringBuilder msg = new StringBuilder();
                            msg.AppendLine("========== HOAN THANH ==========");
                            msg.AppendLine();
                            msg.AppendLine($"Tao duoc: {createdCount}/{doorsInfo.Count} san");
                            msg.AppendLine();
                            msg.AppendLine("THONG TIN:");
                            msg.AppendLine($"• Floor Type: {GetFloorTypeName(floorType)} ({GetFloorTypeThickness(floorType):F0}mm)");
                            msg.AppendLine($"• Base Offset: {baseOffsetMm}mm");
                            msg.AppendLine();

                            if (createdCount > 0)
                            {
                                msg.AppendLine("San da tao:");
                                foreach (var doorInfo in doorsInfo)
                                {
                                    string doorName = doorInfo["name"] as string;
                                    if (!failedDoors.Contains(doorName))
                                    {
                                        double doorWidth = (double)doorInfo["door_width_mm"];
                                        double wallThickness = (double)doorInfo["wall_thickness_mm"];
                                        msg.AppendLine($"  ✓ {doorName} ({doorWidth:F0}x{wallThickness:F0}mm)");
                                    }
                                }
                            }

                            if (failedDoors.Count > 0)
                            {
                                msg.AppendLine();
                                msg.AppendLine("Loi:");
                                foreach (string doorName in failedDoors)
                                {
                                    msg.AppendLine($"  ✗ {doorName}");
                                }
                            }

                            ShowMessage(msg.ToString(), "Hoan thanh");
                            return true;
                        }
                        catch (Exception ex)
                        {
                            trans.RollBack();
                            ShowError("Loi tao san: " + ex.ToString(), "Loi");
                            return false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError("Loi: " + ex.ToString(), "Loi");
                return false;
            }
        }
    }
}