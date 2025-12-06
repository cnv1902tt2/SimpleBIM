using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Drawing = System.Drawing;
// QUAN TRỌNG: Sử dụng alias để tránh ambiguous references
using WinForms = System.Windows.Forms;

namespace SimpleBIM.AS.tab.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class FinishRoomsCeiling : IExternalCommand
    {
        // CLASS NAME: Giữ nguyên ý nghĩa từ Python, thêm suffix "Command" cho C# convention
        // Python: "CREATE CEILINGS FROM MULTIPLE ROOM BOUNDARIES" → C#: "CreateCeilingsFromMultipleRoomBoundariesCommand"

        // PARAMETERS
        private const double MIN_CURVE_LENGTH_M = 0.1;
        private double MIN_CURVE_LENGTH;

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
            WinForms.DialogResult result = WinForms.MessageBox.Show(
                message, title,
                WinForms.MessageBoxButtons.YesNo,
                WinForms.MessageBoxIcon.Question);
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
                    if (element.GetType().Name == "Room")
                        return true;

                    if (element.Category != null)
                    {
                        if (element.Category.Id.Value == (long)BuiltInCategory.OST_Rooms)
                            return true;
                    }

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

        // ============================================================================
        // ROOM INFO STRUCTURE
        // ============================================================================
        private class RoomInfo
        {
            public Room Room { get; set; }
            public string RoomNumber { get; set; }
        }

        // ============================================================================
        // PICK MULTIPLE ROOMS
        // ============================================================================
        private List<Room> PickMultipleRooms(UIDocument uidoc, Document doc)
        {
            List<Room> selectedRooms = new List<Room>();
            HashSet<long> roomIds = new HashSet<long>();

            try
            {
#if DEBUG
                Debug.WriteLine("DEBUG: Starting continuous room selection...");
#endif

                while (true)
                {
                    try
                    {
                        string promptMsg = $"Chon Room trong Floor Plans (Da chon: {selectedRooms.Count} room) - Nhan ESC de dung";

                        Reference reference = uidoc.Selection.PickObject(
                            ObjectType.Element,
                            new RoomSelectionFilter(),
                            promptMsg
                        );

                        Element element = doc.GetElement(reference.ElementId);

                        if (element != null)
                        {
                            // Check if room already selected
                            if (roomIds.Contains(element.Id.Value))
                            {
#if DEBUG
                                Debug.WriteLine($"DEBUG: Room {element.Id.Value} already selected, skipping...");
#endif
                                continue;
                            }

                            // Add to list
                            Room room = element as Room;
                            if (room != null)
                            {
                                selectedRooms.Add(room);
                                roomIds.Add(element.Id.Value);

                                string roomNumber;
                                try
                                {
                                    roomNumber = room.Number;
                                }
                                catch
                                {
                                    roomNumber = element.Id.Value.ToString();
                                }

#if DEBUG
                                Debug.WriteLine($"DEBUG: Room selected: {roomNumber} (Total: {selectedRooms.Count})");
#endif
                            }
                        }
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                    {
#if DEBUG
                        Debug.WriteLine("DEBUG: Selection ended by user (ESC pressed)");
#endif
                        break;
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = ex.Message;
                        if (errorMsg.ToLower().Contains("canceled") || errorMsg.ToLower().Contains("aborted"))
                        {
#if DEBUG
                            Debug.WriteLine("DEBUG: Selection ended by user");
#endif
                            break;
                        }
                        else
                        {
#if DEBUG
                            Debug.WriteLine($"DEBUG: Error in pick loop: {errorMsg}");
#endif
                            ShowError($"Loi: {errorMsg}", "Loi");
                            break;
                        }
                    }
                }

                if (selectedRooms.Count == 0)
                {
                    ShowMessage("Khong chon room nao.", "Thong bao");
                    return null;
                }

                // Confirm number of rooms selected
                string confirmMsg = $"Da chon {selectedRooms.Count} room. Ban co muon tao tran cho cac room nay khong?";
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
#if DEBUG
                Debug.WriteLine($"DEBUG: Error in PickMultipleRooms: {ex.Message}");
#endif
                ShowError($"Loi khi chon room: {ex.Message}", "Loi");
                return null;
            }
        }

        // ============================================================================
        // CEILING TYPE FUNCTIONS
        // ============================================================================
        private string GetCeilingTypeName(CeilingType ceilingTypeElement)
        {
            try
            {
                // Try: Direct Name property
                if (!string.IsNullOrEmpty(ceilingTypeElement.Name))
                    return ceilingTypeElement.Name;

                // Try: SYMBOL_NAME_PARAM
                try
                {
                    Parameter param = ceilingTypeElement.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM);
                    if (param != null && param.HasValue && !string.IsNullOrEmpty(param.AsString()))
                        return param.AsString();
                }
                catch { }

                return "CeilingType_" + ceilingTypeElement.Id.Value.ToString();
            }
            catch
            {
                return "Unknown";
            }
        }

        private double GetCeilingTypeThickness(CeilingType ceilingType)
        {
            try
            {
                // Try: Get thickness from CompoundStructure
                CompoundStructure compoundStructure = ceilingType.GetCompoundStructure();
                if (compoundStructure != null)
                {
                    // QUAN TRỌNG: PyRevit wrapper GetOverallThickness() → C# API GetWidth()
                    double thicknessInternal = compoundStructure.GetWidth();
                    double thicknessMm = UnitUtils.ConvertFromInternalUnits(thicknessInternal, UnitTypeId.Millimeters);
                    return thicknessMm;
                }

                // Try: Thickness parameter
                Parameter param = ceilingType.LookupParameter("Thickness");
                if (param != null && param.HasValue && param.AsDouble() > 0)
                {
                    double thicknessInternal = param.AsDouble();
                    double thicknessMm = UnitUtils.ConvertFromInternalUnits(thicknessInternal, UnitTypeId.Millimeters);
                    return thicknessMm;
                }

                return 0;
            }
            catch
            {
                return 0;
            }
        }

        private Dictionary<string, CeilingType> GetAllCeilingTypes(Document doc)
        {
            var ceilingTypes = new Dictionary<string, CeilingType>();
            try
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                IList<Element> allCeilingTypes = collector.OfClass(typeof(CeilingType)).ToElements();

                foreach (Element element in allCeilingTypes)
                {
                    CeilingType ceilingType = element as CeilingType;
                    if (ceilingType != null)
                    {
                        try
                        {
                            string typeName = GetCeilingTypeName(ceilingType);
                            if (!string.IsNullOrEmpty(typeName) && typeName != "Unknown" && !ceilingTypes.ContainsKey(typeName))
                            {
                                ceilingTypes.Add(typeName, ceilingType);
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }

                if (ceilingTypes.Count == 0)
                {
                    ShowError("Khong tim thay Ceiling Type nao trong project!", "Loi");
                    return null;
                }

                return ceilingTypes;
            }
            catch (Exception ex)
            {
#if DEBUG
                Debug.WriteLine($"DEBUG: Error in GetAllCeilingTypes: {ex.Message}");
#endif
                ShowError($"Loi lay ceiling types: {ex.Message}", "Loi");
                return null;
            }
        }

        // ============================================================================
        // FORM: CEILING PARAMETERS INPUT
        // ============================================================================
        private class MultiRoomCeilingForm : WinForms.Form
        {
            private Dictionary<string, CeilingType> ceilingTypesDict;
            private List<RoomInfo> roomsInfo;
            public CeilingType SelectedCeilingType { get; private set; }
            public double BaseOffsetMm { get; private set; }
            private WinForms.ComboBox cmbCeilingType;
            private WinForms.TextBox txtOffset;
            private WinForms.ListBox lstRooms;

            public MultiRoomCeilingForm(Dictionary<string, CeilingType> ceilingTypesDict, List<RoomInfo> roomsInfo)
            {
                this.ceilingTypesDict = ceilingTypesDict;
                this.roomsInfo = roomsInfo;
                this.SelectedCeilingType = null;
                this.BaseOffsetMm = 0;
                InitializeComponent();
            }

            private void InitializeComponent()
            {
                this.Text = "Tao Tran tu Nhieu Room - Boundary";
                this.Width = 600;
                this.Height = 500;
                this.StartPosition = WinForms.FormStartPosition.CenterScreen;
                this.BackColor = Drawing.Color.White;
                this.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;

                // TITLE
                var lblTitle = new WinForms.Label();
                lblTitle.Text = "TAO TRAN TU NHIEU ROOM - BOUNDARY";
                lblTitle.Font = new Drawing.Font("Arial", 12, Drawing.FontStyle.Bold);
                lblTitle.ForeColor = Drawing.Color.DarkBlue;
                lblTitle.Location = new Drawing.Point(20, 20);
                lblTitle.Size = new Drawing.Size(560, 25);
                this.Controls.Add(lblTitle);

                // ROOM LIST
                var lblRooms = new WinForms.Label();
                lblRooms.Text = $"DANH SACH ROOM DA CHON ({roomsInfo.Count} room)";
                lblRooms.Font = new Drawing.Font("Arial", 10, Drawing.FontStyle.Bold);
                lblRooms.ForeColor = Drawing.Color.DarkGreen;
                lblRooms.Location = new Drawing.Point(20, 55);
                lblRooms.Size = new Drawing.Size(560, 20);
                this.Controls.Add(lblRooms);

                // LISTBOX
                lstRooms = new WinForms.ListBox();
                lstRooms.Location = new Drawing.Point(20, 80);
                lstRooms.Size = new Drawing.Size(560, 150);
                lstRooms.Font = new Drawing.Font("Arial", 9);

                foreach (RoomInfo roomInfo in roomsInfo)
                {
                    string itemText = $"Room: {roomInfo.RoomNumber}";
                    lstRooms.Items.Add(itemText);
                }

                this.Controls.Add(lstRooms);

                // CEILING TYPE
                var lblCeilingType = new WinForms.Label();
                lblCeilingType.Text = "Loai Tran:";
                lblCeilingType.Font = new Drawing.Font("Arial", 10);
                lblCeilingType.Location = new Drawing.Point(20, 245);
                lblCeilingType.Size = new Drawing.Size(100, 20);
                this.Controls.Add(lblCeilingType);

                cmbCeilingType = new WinForms.ComboBox();
                cmbCeilingType.Location = new Drawing.Point(130, 245);
                cmbCeilingType.Size = new Drawing.Size(450, 25);
                cmbCeilingType.Font = new Drawing.Font("Arial", 9);
                cmbCeilingType.DropDownStyle = WinForms.ComboBoxStyle.DropDownList;

                List<string> ceilingTypeNames = ceilingTypesDict.Keys.OrderBy(k => k).ToList();
                foreach (string name in ceilingTypeNames)
                {
                    CeilingType ceilingType = ceilingTypesDict[name];
                    double thickness = GetCeilingTypeThickness(ceilingType);
                    string displayName = thickness > 0 ?
                        $"{name} ({thickness:F0}mm)" :
                        name;
                    cmbCeilingType.Items.Add(displayName);
                }

                if (ceilingTypeNames.Count > 0)
                    cmbCeilingType.SelectedIndex = 0;

                this.Controls.Add(cmbCeilingType);

                // BASE OFFSET
                var lblOffset = new WinForms.Label();
                lblOffset.Text = "Base Offset (mm):";
                lblOffset.Font = new Drawing.Font("Arial", 10);
                lblOffset.Location = new Drawing.Point(20, 285);
                lblOffset.Size = new Drawing.Size(100, 20);
                this.Controls.Add(lblOffset);

                txtOffset = new WinForms.TextBox();
                txtOffset.Location = new Drawing.Point(130, 285);
                txtOffset.Size = new Drawing.Size(450, 25);
                txtOffset.Font = new Drawing.Font("Arial", 9);
                txtOffset.Text = "0";
                this.Controls.Add(txtOffset);

                // BUTTONS
                var btnOk = new WinForms.Button();
                btnOk.Text = $"TAO {roomsInfo.Count} TRAN";
                btnOk.Location = new Drawing.Point(180, 330);
                btnOk.Size = new Drawing.Size(150, 35);
                btnOk.BackColor = Drawing.Color.LightBlue;
                btnOk.Font = new Drawing.Font("Arial", 10, Drawing.FontStyle.Bold);
                btnOk.Click += BtnOK_Click;
                this.Controls.Add(btnOk);

                var btnCancel = new WinForms.Button();
                btnCancel.Text = "HUY";
                btnCancel.Location = new Drawing.Point(340, 330);
                btnCancel.Size = new Drawing.Size(90, 35);
                btnCancel.Font = new Drawing.Font("Arial", 10);
                btnCancel.Click += BtnCancel_Click;
                this.Controls.Add(btnCancel);
            }

            private double GetCeilingTypeThickness(CeilingType ceilingType)
            {
                try
                {
                    CompoundStructure compoundStructure = ceilingType.GetCompoundStructure();
                    if (compoundStructure != null)
                    {
                        // QUAN TRỌNG: PyRevit wrapper GetOverallThickness() → C# API GetWidth()
                        double thicknessInternal = compoundStructure.GetWidth();
                        double thicknessMm = UnitUtils.ConvertFromInternalUnits(thicknessInternal, UnitTypeId.Millimeters);
                        return thicknessMm;
                    }

                    Parameter param = ceilingType.LookupParameter("Thickness");
                    if (param != null && param.HasValue && param.AsDouble() > 0)
                    {
                        double thicknessInternal = param.AsDouble();
                        double thicknessMm = UnitUtils.ConvertFromInternalUnits(thicknessInternal, UnitTypeId.Millimeters);
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
                if (cmbCeilingType.SelectedIndex < 0)
                {
                    WinForms.MessageBox.Show("Vui long chon loai tran!", "Canh bao");
                    return;
                }

                try
                {
                    // Parse offset value
                    if (!double.TryParse(txtOffset.Text, out double offsetValue))
                    {
                        WinForms.MessageBox.Show("Base Offset phai la so!", "Canh bao");
                        return;
                    }

                    BaseOffsetMm = offsetValue;

                    // Parse ceiling type name
                    string displayName = cmbCeilingType.SelectedItem.ToString();
                    string ceilingTypeName;
                    if (displayName.Contains(" ("))
                    {
                        ceilingTypeName = displayName.Split(new string[] { " (" }, StringSplitOptions.None)[0];
                    }
                    else
                    {
                        ceilingTypeName = displayName;
                    }

                    if (ceilingTypesDict.ContainsKey(ceilingTypeName))
                    {
                        SelectedCeilingType = ceilingTypesDict[ceilingTypeName];
                        this.DialogResult = WinForms.DialogResult.OK;
                        this.Close();
                    }
                    else
                    {
                        WinForms.MessageBox.Show($"Khong tim thay ceiling type: {ceilingTypeName}", "Loi");
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
        private List<Curve> GetRoomBoundaryCurves(Room room, Document doc)
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
                    roomId = room.Id.Value.ToString();
                }

#if DEBUG
                Debug.WriteLine($"DEBUG: get_room_boundary_curves - start for room: {roomId}");
#endif

                SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions();
                options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish;
                IList<IList<BoundarySegment>> boundarySegments = room.GetBoundarySegments(options);

                if (boundarySegments == null || boundarySegments.Count == 0)
                {
#if DEBUG
                    Debug.WriteLine("DEBUG: No boundary segments found!");
#endif
                    return null;
                }

#if DEBUG
                Debug.WriteLine($"DEBUG: Processing {boundarySegments.Count} boundary loops...");
#endif

                // Get the first loop (outer boundary)
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

#if DEBUG
                Debug.WriteLine($"DEBUG: Total valid curves: {curves.Count}");
#endif

                if (curves.Count == 0)
                    return null;

                return curves;
            }
            catch (Exception ex)
            {
#if DEBUG
                Debug.WriteLine($"DEBUG: Error in GetRoomBoundaryCurves: {ex.Message}");
#endif
                return null;
            }
        }

        // ============================================================================
        // HEAL GAPS BETWEEN CURVES
        // ============================================================================
        private List<Curve> HealCurveGaps(List<Curve> curves)
        {
            try
            {
#if DEBUG
                Debug.WriteLine($"DEBUG: heal_curve_gaps - start with {curves.Count} curves");
#endif

                List<Curve> healedCurves = new List<Curve>();

                for (int i = 0; i < curves.Count; i++)
                {
                    healedCurves.Add(curves[i]);

                    if (i < curves.Count - 1)
                    {
                        Curve nextCurve = curves[i + 1];
                        XYZ endPoint = curves[i].GetEndPoint(1);
                        XYZ nextStart = nextCurve.GetEndPoint(0);

                        double gapDistance = endPoint.DistanceTo(nextStart);

                        if (gapDistance > 1e-6 && gapDistance < 0.033)
                        {
#if DEBUG
                            Debug.WriteLine($"DEBUG: Gap detected: {gapDistance * 304.8:F6}mm, creating connector line");
#endif
                            Line connectorLine = Line.CreateBound(endPoint, nextStart);
                            healedCurves.Add(connectorLine);
                        }
                    }
                }

                if (healedCurves.Count > 0)
                {
                    XYZ lastEnd = curves[curves.Count - 1].GetEndPoint(1);
                    XYZ firstStart = curves[0].GetEndPoint(0);
                    double closingGap = lastEnd.DistanceTo(firstStart);

                    if (closingGap > 1e-6 && closingGap < 0.033)
                    {
#if DEBUG
                        Debug.WriteLine($"DEBUG: Closing gap detected: {closingGap * 304.8:F6}mm, creating closing line");
#endif
                        Line closingLine = Line.CreateBound(lastEnd, firstStart);
                        healedCurves.Add(closingLine);
                    }
                }

#if DEBUG
                Debug.WriteLine($"DEBUG: heal_curve_gaps - added {healedCurves.Count - curves.Count} connector curves");
#endif
                return healedCurves;
            }
            catch (Exception ex)
            {
#if DEBUG
                Debug.WriteLine($"DEBUG: Error in HealCurveGaps: {ex.Message}");
#endif
                return curves;
            }
        }

        // ============================================================================
        // ORDER CURVES FOR LOOP
        // ============================================================================
        private List<Curve> OrderCurvesForLoop(List<Curve> curves)
        {
            try
            {
#if DEBUG
                Debug.WriteLine($"DEBUG: order_curves_for_loop - start with {curves.Count} curves");
#endif

                if (curves == null || curves.Count == 0)
                    return new List<Curve>();

                List<Curve> remainingCurves = new List<Curve>(curves);
                List<Curve> orderedCurves = new List<Curve>();

                orderedCurves.Add(remainingCurves[0]);
                remainingCurves.RemoveAt(0);

                int maxIterations = curves.Count * 5;
                int iteration = 0;

                while (remainingCurves.Count > 0 && iteration < maxIterations)
                {
                    iteration++;
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

                    double tolerance = 0.01;

                    if (bestDistance < tolerance)
                    {
                        if (bestReversed)
                        {
                            Curve reversedCurve = bestMatch.CreateReversed();
                            orderedCurves.Add(reversedCurve);
                        }
                        else
                        {
                            orderedCurves.Add(bestMatch);
                        }

                        remainingCurves.RemoveAt(bestIndex);
                    }
                    else
                    {
                        break;
                    }
                }

                while (remainingCurves.Count > 0)
                {
                    XYZ lastEnd = orderedCurves[orderedCurves.Count - 1].GetEndPoint(1);

                    double bestDistance = double.MaxValue;
                    int bestIdx = -1;
                    bool bestRev = false;

                    for (int i = 0; i < remainingCurves.Count; i++)
                    {
                        Curve curve = remainingCurves[i];
                        double distStart = lastEnd.DistanceTo(curve.GetEndPoint(0));
                        double distEnd = lastEnd.DistanceTo(curve.GetEndPoint(1));

                        if (distStart < bestDistance)
                        {
                            bestDistance = distStart;
                            bestIdx = i;
                            bestRev = false;
                        }
                        if (distEnd < bestDistance)
                        {
                            bestDistance = distEnd;
                            bestIdx = i;
                            bestRev = true;
                        }
                    }

                    if (bestIdx >= 0)
                    {
                        double minTolerance = 0.001;

                        if (bestDistance > minTolerance)
                        {
                            try
                            {
                                XYZ targetPoint = bestRev ?
                                    remainingCurves[bestIdx].GetEndPoint(1) :
                                    remainingCurves[bestIdx].GetEndPoint(0);
                                Line connector = Line.CreateBound(lastEnd, targetPoint);
                                orderedCurves.Add(connector);
                            }
                            catch { }
                        }

                        if (bestRev)
                        {
                            orderedCurves.Add(remainingCurves[bestIdx].CreateReversed());
                        }
                        else
                        {
                            orderedCurves.Add(remainingCurves[bestIdx]);
                        }

                        remainingCurves.RemoveAt(bestIdx);
                    }
                    else
                    {
                        break;
                    }
                }

                List<Curve> healedCurves = HealCurveGaps(orderedCurves);
                return healedCurves;
            }
            catch (Exception ex)
            {
#if DEBUG
                Debug.WriteLine($"DEBUG: Error in OrderCurvesForLoop: {ex.Message}");
#endif
                return curves;
            }
        }

        // ============================================================================
        // CREATE CURVE LOOP FROM CURVES
        // ============================================================================
        private CurveLoop CreateCurveLoopFromCurves(List<Curve> curves)
        {
            try
            {
#if DEBUG
                Debug.WriteLine($"DEBUG: create_curve_loop_from_curves - start with {curves.Count} curves");
#endif

                if (curves == null || curves.Count == 0)
                    return null;

                List<Curve> orderedCurves = OrderCurvesForLoop(curves);

                if (orderedCurves == null || orderedCurves.Count == 0)
                    return null;

                List<Curve> finalCurves = new List<Curve>(orderedCurves);
                if (finalCurves.Count > 1)
                {
                    XYZ lastEnd = finalCurves[finalCurves.Count - 1].GetEndPoint(1);
                    XYZ firstStart = finalCurves[0].GetEndPoint(0);
                    double closingGap = lastEnd.DistanceTo(firstStart);

                    if (closingGap > 1e-6)
                    {
#if DEBUG
                        Debug.WriteLine("DEBUG: Closing gap detected, creating closing line");
#endif
                        Line closingLine = Line.CreateBound(lastEnd, firstStart);
                        finalCurves.Add(closingLine);
                    }
                }

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
#if DEBUG
                        Debug.WriteLine($"DEBUG: Error appending curve {idx} to loop: {ex.Message}");
#endif
                        return null;
                    }
                }

                if (!curveLoop.IsOpen())
                {
#if DEBUG
                    Debug.WriteLine($"DEBUG: CurveLoop is closed successfully with {finalCurves.Count} curves");
#endif
                    return curveLoop;
                }
                else
                {
#if DEBUG
                    Debug.WriteLine("DEBUG: CurveLoop is still NOT closed!");
#endif
                    return null;
                }
            }
            catch (Exception ex)
            {
#if DEBUG
                Debug.WriteLine($"DEBUG: Error in CreateCurveLoopFromCurves: {ex.Message}");
#endif
                return null;
            }
        }

        // ============================================================================
        // CEILING CREATION
        // ============================================================================
        private Ceiling CreateCeilingFromCurveLoop(CurveLoop curveLoop, CeilingType ceilingType, Level level, double baseOffsetMm, Document doc)
        {
            try
            {
#if DEBUG
                Debug.WriteLine("DEBUG: create_ceiling_from_curve_loop - start");
#endif

                double baseOffsetInternal = UnitUtils.ConvertToInternalUnits(baseOffsetMm, UnitTypeId.Millimeters);

                Ceiling ceiling = Ceiling.Create(doc, new List<CurveLoop> { curveLoop }, ceilingType.Id, level.Id);

                if (ceiling != null)
                {
                    try
                    {
                        Parameter param = ceiling.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM);
                        if (param != null && !param.IsReadOnly)
                        {
                            param.Set(baseOffsetInternal);
                        }
                    }
                    catch { }

                    return ceiling;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
#if DEBUG
                Debug.WriteLine($"DEBUG: Error in CreateCeilingFromCurveLoop: {ex.Message}");
#endif
                return null;
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
                // Initialize MIN_CURVE_LENGTH
                MIN_CURVE_LENGTH = UnitUtils.ConvertToInternalUnits(MIN_CURVE_LENGTH_M, UnitTypeId.Meters);

                ShowMessage(
                    "TAO TRAN TU NHIEU ROOM - BOUNDARY\n\n" +
                    "HUONG DAN:\n" +
                    "1. Click OK de bat dau chon room\n" +
                    "2. Click tung room trong Floor Plans\n" +
                    "3. Nhan ESC de hoan tat chon\n" +
                    "4. Chon loai tran va Base Offset\n" +
                    "5. Tao tran cho tat ca room da chon",
                    "Huong dan"
                );

                // PICK MULTIPLE ROOMS
                List<Room> rooms = PickMultipleRooms(uidoc, doc);
                if (rooms == null || rooms.Count == 0)
                    return Result.Cancelled;

                ShowMessage(
                    $"Da chon {rooms.Count} room\n\nDang lay boundary curves...",
                    "Thong bao"
                );

                // GET CEILING TYPES
                Dictionary<string, CeilingType> ceilingTypesDict = GetAllCeilingTypes(doc);
                if (ceilingTypesDict == null || ceilingTypesDict.Count == 0)
                    return Result.Failed;

                // PROCESS EACH ROOM AND COLLECT INFO
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
                            roomNumber = room.Id.Value.ToString();
                        }

                        List<Curve> curves = GetRoomBoundaryCurves(room, doc);

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
#if DEBUG
                        Debug.WriteLine($"DEBUG: Error processing room: {ex.Message}");
#endif
                        continue;
                    }
                }

                if (roomsInfo.Count == 0)
                {
                    ShowError("Khong the xu ly room nao!", "Loi");
                    return Result.Failed;
                }

                // SHOW FORM
                using (MultiRoomCeilingForm form = new MultiRoomCeilingForm(ceilingTypesDict, roomsInfo))
                {
                    if (form.ShowDialog() != WinForms.DialogResult.OK)
                    {
                        ShowMessage("Ban da huy tao tran.", "Thong bao");
                        return Result.Cancelled;
                    }

                    CeilingType ceilingType = form.SelectedCeilingType;
                    double baseOffsetMm = form.BaseOffsetMm;

                    if (ceilingType == null)
                    {
                        ShowError("Khong chon duoc ceiling type!", "Loi");
                        return Result.Failed;
                    }

                    // CREATE CEILINGS
                    ShowMessage($"Dang tao tran cho {roomsInfo.Count} room...", "Thong bao");

                    using (Transaction t = new Transaction(doc, "Create Ceilings from Room Boundaries"))
                    {
                        t.Start();

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

                                    // Get curves
                                    List<Curve> curves = GetRoomBoundaryCurves(room, doc);
                                    if (curves == null || curves.Count == 0)
                                    {
                                        failedRooms.Add($"Room {roomNumber}");
                                        continue;
                                    }

                                    // Get level
                                    Level level = room.Level;
                                    if (level == null)
                                    {
                                        failedRooms.Add($"Room {roomNumber}");
                                        continue;
                                    }

                                    // Create curve loop
                                    CurveLoop curveLoop = CreateCurveLoopFromCurves(curves);
                                    if (curveLoop == null)
                                    {
                                        failedRooms.Add($"Room {roomNumber}");
                                        continue;
                                    }

                                    // Create ceiling
                                    Ceiling ceiling = CreateCeilingFromCurveLoop(curveLoop, ceilingType, level, baseOffsetMm, doc);
                                    if (ceiling != null)
                                    {
                                        createdCount++;
#if DEBUG
                                        Debug.WriteLine($"DEBUG: Ceiling created for room: {roomNumber}");
#endif
                                    }
                                    else
                                    {
                                        failedRooms.Add($"Room {roomNumber}");
                                    }
                                }
                                catch (Exception ex)
                                {
#if DEBUG
                                    Debug.WriteLine($"DEBUG: Error creating ceiling for room {roomInfo.RoomNumber}: {ex.Message}");
#endif
                                    failedRooms.Add($"Room {roomInfo.RoomNumber}");
                                    continue;
                                }
                            }

                            t.Commit();

                            // RESULT MESSAGE
                            string msg = "========== HOAN THANH ==========\n\n";
                            msg += $"Tao duoc: {createdCount}/{roomsInfo.Count} tran\n\n";
                            msg += "THONG TIN:\n";
                            msg += $"• Ceiling Type: {GetCeilingTypeName(ceilingType)} ({GetCeilingTypeThickness(ceilingType):F0}mm)\n";
                            msg += $"• Base Offset: {baseOffsetMm}mm\n\n";

                            if (createdCount > 0)
                            {
                                msg += "Tran da tao:\n";
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

#if DEBUG
                            Debug.WriteLine($"DEBUG: Success! Created {createdCount} ceilings");
#endif
                        }
                        catch (Exception ex)
                        {
                            t.RollBack();
#if DEBUG
                            Debug.WriteLine($"DEBUG: Exception during ceiling creation: {ex.Message}");
                            Debug.WriteLine(ex.StackTrace);
#endif
                            ShowError($"Loi tao tran: {ex.Message}", "Loi");
                            return Result.Failed;
                        }
                    }
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
#if DEBUG
                Debug.WriteLine($"FATAL ERROR: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);
#endif
                ShowError($"Loi: {ex.Message}", "Loi");
                return Result.Failed;
            }
        }
    }
}