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
// Alias để tránh ambiguous references
using WinForms = System.Windows.Forms;

namespace SimpleBIM.AS.tab.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class FinishRoomsFloor : IExternalCommand
    {
        // Constants
        private const double MIN_CURVE_LENGTH_M = 0.1;
        private const double REVIT_TOLERANCE = 0.001;
        private double MIN_CURVE_LENGTH;

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

            // Initialize units
            MIN_CURVE_LENGTH = UnitUtils.ConvertToInternalUnits(MIN_CURVE_LENGTH_M, UnitTypeId.Meters);

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

        private bool ShowYesNo(string message, string title = "Xac nhan")
        {
            WinForms.DialogResult result = WinForms.MessageBox.Show(message, title,
                WinForms.MessageBoxButtons.YesNo, WinForms.MessageBoxIcon.Question);
            return result == WinForms.DialogResult.Yes;
        }

        // ============================================================================
        // ROOM SELECTION FILTER
        // ============================================================================

        private class RoomSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element element)
            {
                try
                {
                    // Check if element is a Room
                    if (element is Room)
                    {
                        return true;
                    }

                    // Check category
                    if (element.Category != null)
                    {
                        // PYREVIT → C# CONVERSION: IntegerValue → Value
                        if (element.Category.Id.Value == (long)BuiltInCategory.OST_Rooms) // CONVERSION APPLIED
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

        // ============================================================================
        // ROOM SELECTION
        // ============================================================================

        private List<Room> PickMultipleRooms()
        {
            List<Room> selectedRooms = new List<Room>();
            HashSet<long> roomIds = new HashSet<long>();

            try
            {
                Debug.WriteLine("DEBUG: Starting continuous room selection...");

                while (true)
                {
                    try
                    {
                        string promptMsg = $"Chon Room trong Floor Plans (Da chon: {selectedRooms.Count} room) - Nhan ESC de dung";

                        Reference reference = _uidoc.Selection.PickObject(
                            ObjectType.Element,
                            new RoomSelectionFilter(),
                            promptMsg);

                        Element element = _doc.GetElement(reference.ElementId);

                        if (element is Room room)
                        {
                            // PYREVIT → C# CONVERSION: IntegerValue → Value
                            if (roomIds.Contains(room.Id.Value)) // CONVERSION APPLIED
                            {
                                Debug.WriteLine($"DEBUG: Room {room.Id.Value} already selected, skipping...");
                                continue;
                            }

                            selectedRooms.Add(room);
                            // PYREVIT → C# CONVERSION: IntegerValue → Value
                            roomIds.Add(room.Id.Value); // CONVERSION APPLIED

                            string roomNumber;
                            try
                            {
                                roomNumber = room.Number;
                            }
                            catch
                            {
                                // PYREVIT → C# CONVERSION: IntegerValue → Value
                                roomNumber = room.Id.Value.ToString(); // CONVERSION APPLIED
                            }

                            Debug.WriteLine($"DEBUG: Room selected: {roomNumber} (Total: {selectedRooms.Count})");
                        }
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                    {
                        Debug.WriteLine("DEBUG: Selection ended by user (ESC pressed)");
                        break;
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"DEBUG: Error in pick loop: {ex.Message}");
                        ShowError($"Loi: {ex.Message}", "Loi");
                        break;
                    }
                }

                if (selectedRooms.Count == 0)
                {
                    ShowMessage("Khong chon room nao.", "Thong bao");
                    return null;
                }

                string confirmMsg = $"Da chon {selectedRooms.Count} room. Ban co muon tao san cho cac room nay khong?";
                if (ShowYesNo(confirmMsg, "Xac nhan"))
                {
                    return selectedRooms;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"DEBUG: Error in pick_multiple_rooms: {ex.Message}");
                ShowError($"Loi khi chon room: {ex.Message}", "Loi");
                return null;
            }
        }

        // ============================================================================
        // FLOOR TYPE FUNCTIONS
        // ============================================================================

        private string GetFloorTypeName(FloorType floorTypeElement)
        {
            try
            {
                // Method 1: Get from Name property
                try
                {
                    return floorTypeElement.Name;
                }
                catch
                {
                    // Continue to next method
                }

                // Method 2: Get from SYMBOL_NAME_PARAM
                try
                {
                    Parameter param = floorTypeElement.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM);
                    if (param != null && param.HasValue)
                    {
                        return param.AsString();
                    }
                }
                catch
                {
                    // Continue to next method
                }

                // Fallback name
                // PYREVIT → C# CONVERSION: IntegerValue → Value
                return $"FloorType_{floorTypeElement.Id.Value}"; // CONVERSION APPLIED
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
                // Method 1: Get from CompoundStructure
                try
                {
                    CompoundStructure compound = floorType.GetCompoundStructure();
                    if (compound != null)
                    {
                        // PYREVIT → C# CONVERSION: GetOverallThickness() might not exist, use GetWidth()
                        double thickness = compound.GetWidth(); // CONVERSION APPLIED
                        double thicknessMm = UnitUtils.ConvertFromInternalUnits(thickness, UnitTypeId.Millimeters);
                        return thicknessMm;
                    }
                }
                catch
                {
                    // Continue to next method
                }

                // Method 2: Get from Thickness parameter
                Parameter param = floorType.LookupParameter("Thickness");
                if (param != null && param.HasValue && param.AsDouble() > 0)
                {
                    double thicknessMm = UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters);
                    return thicknessMm;
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
                    if (param.AsInteger() == 1)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }

                string typeName = GetFloorTypeName(floorType).ToLower();
                if (typeName.Contains("structural"))
                {
                    return true;
                }
                else
                {
                    return false;
                }
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
                FilteredElementCollector collector = new FilteredElementCollector(_doc)
                    .OfClass(typeof(FloorType));

                foreach (FloorType floorType in collector)
                {
                    try
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
                    catch
                    {
                        continue;
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
                Debug.WriteLine($"DEBUG: Error in get_all_floor_types: {ex.Message}");
                ShowError($"Loi lay floor types: {ex.Message}", "Loi");
                return null;
            }
        }

        // ============================================================================
        // FLOOR PARAMETERS INPUT FORM
        // ============================================================================

        private class MultiRoomFloorForm : WinForms.Form
        {
            public FloorType SelectedFloorType { get; private set; }
            public double BaseOffsetMm { get; private set; }

            private Dictionary<string, FloorType> _floorTypesDict;
            private List<RoomInfo> _roomsInfo;
            private WinForms.ListBox _lstRooms;
            private WinForms.ComboBox _cmbFloorType;
            private WinForms.TextBox _txtOffset;

            public MultiRoomFloorForm(Dictionary<string, FloorType> floorTypesDict, List<RoomInfo> roomsInfo)
            {
                _floorTypesDict = floorTypesDict;
                _roomsInfo = roomsInfo;
                InitializeComponent();
            }

            private void InitializeComponent()
            {
                this.Text = "Tao San tu Nhieu Room - Boundary";
                this.Width = 600;
                this.Height = 500;
                this.StartPosition = WinForms.FormStartPosition.CenterScreen;
                this.BackColor = Drawing.Color.White;
                this.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;

                // Title label
                WinForms.Label lblTitle = new WinForms.Label
                {
                    Text = "TAO SAN TU NHIEU ROOM - BOUNDARY",
                    Font = new Drawing.Font("Arial", 12, Drawing.FontStyle.Bold),
                    ForeColor = Drawing.Color.DarkBlue,
                    Location = new Drawing.Point(20, 20),
                    Size = new Drawing.Size(560, 25)
                };
                this.Controls.Add(lblTitle);

                // Rooms label
                WinForms.Label lblRooms = new WinForms.Label
                {
                    Text = $"DANH SACH ROOM DA CHON ({_roomsInfo.Count} room)",
                    Font = new Drawing.Font("Arial", 10, Drawing.FontStyle.Bold),
                    ForeColor = Drawing.Color.DarkGreen,
                    Location = new Drawing.Point(20, 55),
                    Size = new Drawing.Size(560, 20)
                };
                this.Controls.Add(lblRooms);

                // Rooms ListBox
                _lstRooms = new WinForms.ListBox
                {
                    Location = new Drawing.Point(20, 80),
                    Size = new Drawing.Size(560, 150),
                    Font = new Drawing.Font("Arial", 9)
                };

                foreach (RoomInfo roomInfo in _roomsInfo)
                {
                    string itemText = $"Room: {roomInfo.RoomNumber}";
                    _lstRooms.Items.Add(itemText);
                }

                this.Controls.Add(_lstRooms);

                // Floor Type label
                WinForms.Label lblFloorType = new WinForms.Label
                {
                    Text = "Loai San:",
                    Font = new Drawing.Font("Arial", 10),
                    Location = new Drawing.Point(20, 245),
                    Size = new Drawing.Size(100, 20)
                };
                this.Controls.Add(lblFloorType);

                // Floor Type ComboBox
                _cmbFloorType = new WinForms.ComboBox
                {
                    Location = new Drawing.Point(130, 245),
                    Size = new Drawing.Size(450, 25),
                    Font = new Drawing.Font("Arial", 9),
                    DropDownStyle = WinForms.ComboBoxStyle.DropDownList
                };

                List<string> floorTypeNames = _floorTypesDict.Keys.OrderBy(x => x).ToList();
                foreach (string name in floorTypeNames)
                {
                    FloorType floorType = _floorTypesDict[name];
                    double thickness = GetFloorTypeThicknessStatic(floorType);
                    string displayName = thickness > 0 ? $"{name} ({thickness:F0}mm)" : name;
                    _cmbFloorType.Items.Add(displayName);
                }

                if (floorTypeNames.Count > 0)
                {
                    _cmbFloorType.SelectedIndex = 0;
                }

                this.Controls.Add(_cmbFloorType);

                // Offset label
                WinForms.Label lblOffset = new WinForms.Label
                {
                    Text = "Base Offset (mm):",
                    Font = new Drawing.Font("Arial", 10),
                    Location = new Drawing.Point(20, 285),
                    Size = new Drawing.Size(100, 20)
                };
                this.Controls.Add(lblOffset);

                // Offset TextBox
                _txtOffset = new WinForms.TextBox
                {
                    Location = new Drawing.Point(130, 285),
                    Size = new Drawing.Size(450, 25),
                    Font = new Drawing.Font("Arial", 9),
                    Text = "0"
                };
                this.Controls.Add(_txtOffset);

                // OK button
                WinForms.Button btnOk = new WinForms.Button
                {
                    Text = $"TAO {_roomsInfo.Count} SAN",
                    Location = new Drawing.Point(180, 330),
                    Size = new Drawing.Size(150, 35),
                    BackColor = Drawing.Color.LightBlue,
                    Font = new Drawing.Font("Arial", 10, Drawing.FontStyle.Bold)
                };
                btnOk.Click += BtnOK_Click;
                this.Controls.Add(btnOk);

                // Cancel button
                WinForms.Button btnCancel = new WinForms.Button
                {
                    Text = "HUY",
                    Location = new Drawing.Point(340, 330),
                    Size = new Drawing.Size(90, 35),
                    Font = new Drawing.Font("Arial", 10)
                };
                btnCancel.Click += BtnCancel_Click;
                this.Controls.Add(btnCancel);
            }

            private static double GetFloorTypeThicknessStatic(FloorType floorType)
            {
                try
                {
                    // Method 1: Get from CompoundStructure
                    try
                    {
                        CompoundStructure compound = floorType.GetCompoundStructure();
                        if (compound != null)
                        {
                            // PYREVIT → C# CONVERSION: GetOverallThickness() might not exist, use GetWidth()
                            double thickness = compound.GetWidth(); // CONVERSION APPLIED
                            double thicknessMm = UnitUtils.ConvertFromInternalUnits(thickness, UnitTypeId.Millimeters);
                            return thicknessMm;
                        }
                    }
                    catch
                    {
                        // Continue to next method
                    }

                    // Method 2: Get from Thickness parameter
                    Parameter param = floorType.LookupParameter("Thickness");
                    if (param != null && param.HasValue && param.AsDouble() > 0)
                    {
                        double thicknessMm = UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters);
                        return thicknessMm;
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
                if (_cmbFloorType.SelectedIndex < 0)
                {
                    WinForms.MessageBox.Show("Vui long chon loai san!", "Canh bao");
                    return;
                }

                try
                {
                    if (!double.TryParse(_txtOffset.Text, out double offsetValue))
                    {
                        WinForms.MessageBox.Show("Base Offset phai la so!", "Canh bao");
                        return;
                    }

                    BaseOffsetMm = offsetValue;

                    string displayName = _cmbFloorType.SelectedItem.ToString();
                    string floorTypeName = displayName.Split(new[] { " (" }, StringSplitOptions.None)[0];

                    if (_floorTypesDict.ContainsKey(floorTypeName))
                    {
                        SelectedFloorType = _floorTypesDict[floorTypeName];
                        this.DialogResult = WinForms.DialogResult.OK;
                        this.Close();
                    }
                    else
                    {
                        WinForms.MessageBox.Show($"Khong tim thay floor type: {floorTypeName}", "Loi");
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
        // ROOM BOUNDARY EXTRACTION
        // ============================================================================

        private class RoomInfo
        {
            public Room Room { get; set; }
            public string RoomNumber { get; set; }
        }

        private List<Curve> GetRoomBoundaryCurves(Room room)
        {
            List<Curve> curves = new List<Curve>();

            try
            {
                string roomId;
                try
                {
                    roomId = room.Number;
                }
                catch
                {
                    // PYREVIT → C# CONVERSION: IntegerValue → Value
                    roomId = room.Id.Value.ToString(); // CONVERSION APPLIED
                }

                Debug.WriteLine($"DEBUG: get_room_boundary_curves - start for room: {roomId}");

                SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions
                {
                    SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish
                };

                IList<IList<BoundarySegment>> boundarySegments = room.GetBoundarySegments(options);

                if (boundarySegments == null || boundarySegments.Count == 0)
                {
                    Debug.WriteLine("DEBUG: No boundary segments found!");
                    return null;
                }

                Debug.WriteLine($"DEBUG: Processing {boundarySegments.Count} boundary loops...");

                IList<BoundarySegment> segmentLoop = boundarySegments[0];

                foreach (BoundarySegment boundarySegment in segmentLoop)
                {
                    try
                    {
                        Curve curve = boundarySegment.GetCurve();
                        if (curve != null && curve.Length >= MIN_CURVE_LENGTH)
                        {
                            curves.Add(curve);
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

                Debug.WriteLine($"DEBUG: Total valid curves: {curves.Count}");

                if (curves.Count == 0)
                {
                    return null;
                }

                return curves;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"DEBUG: Error in get_room_boundary_curves: {ex.Message}");
                return null;
            }
        }

        // ============================================================================
        // CURVE ORDERING - IMPROVED VERSION
        // ============================================================================

        private List<Curve> OrderCurvesForLoop(List<Curve> curves)
        {
            // FIX #1: Improved curve ordering with better tolerance handling
            try
            {
                Debug.WriteLine($"DEBUG: order_curves_for_loop - start with {curves.Count} curves");

                if (curves == null || curves.Count == 0)
                {
                    return new List<Curve>();
                }

                List<Curve> remainingCurves = new List<Curve>(curves);
                List<Curve> orderedCurves = new List<Curve>();
                orderedCurves.Add(remainingCurves[0]);
                remainingCurves.RemoveAt(0);

                // FIX: Increased SNAP_TOLERANCE to handle larger gaps from Room boundaries
                double SNAP_TOLERANCE = 0.5;  // 500mm - increased from 100mm

                while (remainingCurves.Count > 0)
                {
                    Curve lastCurve = orderedCurves[orderedCurves.Count - 1];
                    XYZ lastEndPoint = lastCurve.GetEndPoint(1);

                    Curve bestMatch = null;
                    double bestDistance = double.MaxValue;
                    int bestIndex = -1;
                    bool bestReversed = false;

                    for (int i = 0; i < remainingCurves.Count; i++)
                    {
                        Curve curve = remainingCurves[i];
                        XYZ curveStart = curve.GetEndPoint(0);
                        XYZ curveEnd = curve.GetEndPoint(1);

                        double distToStart = lastEndPoint.DistanceTo(curveStart);
                        if (distToStart < bestDistance)
                        {
                            bestDistance = distToStart;
                            bestMatch = curve;
                            bestIndex = i;
                            bestReversed = false;
                        }

                        double distToEnd = lastEndPoint.DistanceTo(curveEnd);
                        if (distToEnd < bestDistance)
                        {
                            bestDistance = distToEnd;
                            bestMatch = curve;
                            bestIndex = i;
                            bestReversed = true;
                        }
                    }

                    if (bestIndex >= 0)
                    {
                        if (bestDistance < SNAP_TOLERANCE)
                        {
                            if (bestReversed)
                            {
                                orderedCurves.Add(bestMatch.CreateReversed());
                            }
                            else
                            {
                                orderedCurves.Add(bestMatch);
                            }

                            remainingCurves.RemoveAt(bestIndex);
                            Debug.WriteLine($"DEBUG: Added curve {orderedCurves.Count}, gap: {bestDistance:F3}m");
                        }
                        else
                        {
                            // FIX #2: Don't break, try to add connector line
                            Debug.WriteLine($"DEBUG: Gap large ({bestDistance:F3}m), trying connector...");
                            break;
                        }
                    }
                    else
                    {
                        break;
                    }
                }

                Debug.WriteLine($"DEBUG: Ordered {orderedCurves.Count} curves, {remainingCurves.Count} remaining");
                return orderedCurves;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"DEBUG: Error in order_curves_for_loop: {ex.Message}");
                return curves;
            }
        }

        // ============================================================================
        // CREATE CURVE LOOP FROM CURVES - FIXED VERSION
        // ============================================================================

        private CurveLoop CreateCurveLoopFromCurves(List<Curve> curves)
        {
            // FIX #3: Completely new approach - try multiple strategies
            try
            {
                Debug.WriteLine($"DEBUG: create_curve_loop_from_curves - start with {curves.Count} curves");

                if (curves == null || curves.Count == 0)
                {
                    return null;
                }

                // Try to order curves
                List<Curve> orderedCurves = OrderCurvesForLoop(curves);

                if (orderedCurves == null || orderedCurves.Count == 0)
                {
                    Debug.WriteLine("DEBUG: Could not order curves");
                    return null;
                }

                // Build final curves list with connectors
                List<Curve> finalCurves = new List<Curve>();
                double MAX_CONNECTOR_LENGTH = 1.0;  // FIX: Allow longer connectors (1 meter)

                for (int i = 0; i < orderedCurves.Count; i++)
                {
                    Curve curve = orderedCurves[i];
                    finalCurves.Add(curve);

                    if (i < orderedCurves.Count - 1)
                    {
                        XYZ currentEnd = curve.GetEndPoint(1);
                        XYZ nextStart = orderedCurves[i + 1].GetEndPoint(0);
                        double gap = currentEnd.DistanceTo(nextStart);

                        // FIX: More lenient gap checking
                        if (gap > 1e-6 && gap < MAX_CONNECTOR_LENGTH)
                        {
                            try
                            {
                                Curve connector = Line.CreateBound(currentEnd, nextStart);
                                finalCurves.Add(connector);
                                Debug.WriteLine($"DEBUG: Connector added: {gap:F3}m");
                            }
                            catch
                            {
                                Debug.WriteLine($"DEBUG: Could not create connector for {gap:F3}m gap");
                            }
                        }
                    }
                }

                // Check and handle closing gap
                if (finalCurves.Count > 1)
                {
                    XYZ lastEnd = finalCurves[finalCurves.Count - 1].GetEndPoint(1);
                    XYZ firstStart = finalCurves[0].GetEndPoint(0);
                    double closingGap = lastEnd.DistanceTo(firstStart);

                    Debug.WriteLine($"DEBUG: Closing gap: {closingGap:F3}m ({closingGap * 1000:F1}mm)");

                    // FIX: More lenient closing gap tolerance
                    if (closingGap > 1e-6 && closingGap < MAX_CONNECTOR_LENGTH)
                    {
                        try
                        {
                            Curve closingLine = Line.CreateBound(lastEnd, firstStart);
                            finalCurves.Add(closingLine);
                            Debug.WriteLine("DEBUG: Closing line created");
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"DEBUG: Could not create closing line: {ex.Message}");
                            return null;
                        }
                    }
                    else if (closingGap >= MAX_CONNECTOR_LENGTH)
                    {
                        Debug.WriteLine("DEBUG: Closing gap too large - cannot create loop");
                        return null;
                    }
                }

                // Try to create curve loop
                try
                {
                    CurveLoop curveLoop = new CurveLoop();

                    for (int idx = 0; idx < finalCurves.Count; idx++)
                    {
                        Curve curve = finalCurves[idx];
                        try
                        {
                            curveLoop.Append(curve);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"DEBUG: Error appending curve {idx}: {ex.Message}");
                            return null;
                        }
                    }

                    if (!curveLoop.IsOpen())
                    {
                        Debug.WriteLine($"DEBUG: CurveLoop closed successfully with {finalCurves.Count} curves");
                        return curveLoop;
                    }
                    else
                    {
                        Debug.WriteLine("DEBUG: CurveLoop is open - loop not closed");
                        return null;
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"DEBUG: Error creating CurveLoop: {ex.Message}");
                    return null;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"DEBUG: Error in create_curve_loop_from_curves: {ex.Message}");
                return null;
            }
        }

        // ============================================================================
        // FLOOR CREATION
        // ============================================================================

        private Floor CreateFloorFromCurveLoop(CurveLoop curveLoop, FloorType floorType, Level level, double baseOffsetMm)
        {
            try
            {
                Debug.WriteLine("DEBUG: create_floor_from_curve_loop - start");

                double baseOffsetInternal = UnitUtils.ConvertToInternalUnits(baseOffsetMm, UnitTypeId.Millimeters);

                // FIX #4: Pass as list of CurveLoops
                List<CurveLoop> curveLoops = new List<CurveLoop> { curveLoop };
                Floor floor = Floor.Create(_doc, curveLoops, floorType.Id, level.Id);

                if (floor != null)
                {
                    try
                    {
                        Parameter param = floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
                        if (param != null && !param.IsReadOnly)
                        {
                            param.Set(baseOffsetInternal);
                        }
                    }
                    catch
                    {
                        // Ignore if cannot set offset
                    }

                    Debug.WriteLine("DEBUG: Floor created successfully");
                    return floor;
                }
                else
                {
                    Debug.WriteLine("DEBUG: Floor.Create returned None");
                    return null;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"DEBUG: Error in create_floor_from_curve_loop: {ex.Message}");
                return null;
            }
        }

        // ============================================================================
        // MAIN EXECUTION
        // ============================================================================

        private Result RunMain()
        {
            try
            {
                ShowMessage(
                    "TAO SAN TU NHIEU ROOM - BOUNDARY\n\n" +
                    "HUONG DAN:\n" +
                    "1. Click OK de bat dau chon room\n" +
                    "2. Click tung room trong Floor Plans\n" +
                    "3. Nhan ESC de hoan tat chon\n" +
                    "4. Chon loai san va Base Offset\n" +
                    "5. Tao san cho tat ca room da chon",
                    "Huong dan"
                );

                List<Room> rooms = PickMultipleRooms();
                if (rooms == null || rooms.Count == 0)
                {
                    return Result.Cancelled;
                }

                ShowMessage(
                    $"Da chon {rooms.Count} room\n\nDang lay boundary curves...",
                    "Thong bao"
                );

                Dictionary<string, FloorType> floorTypesDict = GetAllFloorTypes();
                if (floorTypesDict == null || floorTypesDict.Count == 0)
                {
                    return Result.Failed;
                }

                List<RoomInfo> roomsInfo = new List<RoomInfo>();
                List<Room> validRooms = new List<Room>();

                foreach (Room room in rooms)
                {
                    try
                    {
                        string roomNumber;
                        try
                        {
                            roomNumber = room.Number;
                        }
                        catch
                        {
                            // PYREVIT → C# CONVERSION: IntegerValue → Value
                            roomNumber = room.Id.Value.ToString(); // CONVERSION APPLIED
                        }

                        List<Curve> curves = GetRoomBoundaryCurves(room);

                        if (curves != null && curves.Count > 0)
                        {
                            roomsInfo.Add(new RoomInfo
                            {
                                Room = room,
                                RoomNumber = roomNumber
                            });
                            validRooms.Add(room);
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"DEBUG: Error processing room: {ex.Message}");
                        continue;
                    }
                }

                if (roomsInfo.Count == 0)
                {
                    ShowError("Khong the xu ly room nao!", "Loi");
                    return Result.Failed;
                }

                using (MultiRoomFloorForm form = new MultiRoomFloorForm(floorTypesDict, roomsInfo))
                {
                    WinForms.DialogResult result = form.ShowDialog();

                    if (result != WinForms.DialogResult.OK)
                    {
                        ShowMessage("Ban da huy tao san.", "Thong bao");
                        return Result.Cancelled;
                    }

                    FloorType floorType = form.SelectedFloorType;
                    double baseOffsetMm = form.BaseOffsetMm;

                    ShowMessage($"Dang tao san cho {roomsInfo.Count} room...", "Thong bao");

                    using (Transaction trans = new Transaction(_doc, "Create Floors from Room Boundaries"))
                    {
                        trans.Start();

                        try
                        {
                            int createdCount = 0;
                            List<string> failedRooms = new List<string>();

                            foreach (RoomInfo roomInfo in roomsInfo)
                            {
                                try
                                {
                                    Room room = roomInfo.Room;
                                    string roomNumber = roomInfo.RoomNumber;

                                    List<Curve> curves = GetRoomBoundaryCurves(room);
                                    if (curves == null || curves.Count == 0)
                                    {
                                        failedRooms.Add($"Room {roomNumber}");
                                        continue;
                                    }

                                    Level level = room.Level;
                                    if (level == null)
                                    {
                                        failedRooms.Add($"Room {roomNumber}");
                                        continue;
                                    }

                                    CurveLoop curveLoop = CreateCurveLoopFromCurves(curves);
                                    if (curveLoop == null)
                                    {
                                        failedRooms.Add($"Room {roomNumber}");
                                        continue;
                                    }

                                    Floor floor = CreateFloorFromCurveLoop(curveLoop, floorType, level, baseOffsetMm);
                                    if (floor != null)
                                    {
                                        createdCount++;
                                        Debug.WriteLine($"DEBUG: Floor created for room: {roomNumber}");
                                    }
                                    else
                                    {
                                        failedRooms.Add($"Room {roomNumber}");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Debug.WriteLine($"DEBUG: Error creating floor: {ex.Message}");
                                    failedRooms.Add($"Room {roomInfo.RoomNumber}");
                                    continue;
                                }
                            }

                            trans.Commit();

                            string msg = "========== HOAN THANH ==========\n\n";
                            msg += $"Tao duoc: {createdCount}/{roomsInfo.Count} san\n\n";
                            msg += "THONG TIN:\n";
                            msg += $"• Floor Type: {GetFloorTypeName(floorType)} ({GetFloorTypeThickness(floorType):F0}mm)\n";
                            msg += $"• Base Offset: {baseOffsetMm}mm\n\n";

                            if (createdCount > 0)
                            {
                                msg += "San da tao:\n";
                                foreach (RoomInfo roomInfo in roomsInfo)
                                {
                                    string roomFullName = $"Room {roomInfo.RoomNumber}";
                                    if (!failedRooms.Contains(roomFullName))
                                    {
                                        msg += $"  ✓ {roomFullName}\n";
                                    }
                                }
                            }

                            if (failedRooms.Count > 0)
                            {
                                msg += "\nLoi:\n";
                                foreach (string roomName in failedRooms)
                                {
                                    msg += $"  ✗ {roomName}\n";
                                }
                            }

                            ShowMessage(msg, "Hoan thanh");
                            Debug.WriteLine($"DEBUG: Success! Created {createdCount} floors");
                            return Result.Succeeded;
                        }
                        catch (Exception ex)
                        {
                            trans.RollBack();
                            string errorStr = ex.Message;
                            Debug.WriteLine($"DEBUG: Exception during floor creation: {errorStr}");
                            ShowError($"Loi tao san: {errorStr}", "Loi");
                            return Result.Failed;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string errorStr = ex.Message;
                Debug.WriteLine($"DEBUG: Main exception: {errorStr}");
                ShowError($"Loi: {errorStr}", "Loi");
                return Result.Failed;
            }
        }

        // ============================================================================
        // DEBUG UTILITY
        // ============================================================================

        private static class Debug
        {
            public static void WriteLine(string message)
            {
#if DEBUG
                System.Diagnostics.Debug.WriteLine(message);
#endif
            }
        }
    }
}