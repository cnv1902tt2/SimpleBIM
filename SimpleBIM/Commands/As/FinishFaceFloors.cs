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

using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;

namespace SimpleBIM.Commands.As
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class FinishFaceFloors : IExternalCommand
    {
        // Constants
        private const double MIN_CURVE_LENGTH_M = 0.01;
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

        // =============================================================================
        // HELPER METHODS
        // =============================================================================

        private void ShowMessage(string message, string title = "Notice")
        {
            WinForms.MessageBox.Show(message, title, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
        }

        private void ShowError(string message, string title = "Error")
        {
            WinForms.MessageBox.Show(message, title, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
        }

        // =============================================================================
        // FACE SELECTION
        // =============================================================================

        private class FaceData
        {
            public Element Element { get; set; }
            public Face Face { get; set; }
            public Reference Reference { get; set; }
        }

        private List<FaceData> PickMultipleFaces()
        {
            List<FaceData> selectedFaces = new List<FaceData>();
            HashSet<(long, double, double, double)> faceIds = new HashSet<(long, double, double, double)>();

            try
            {
                Debug.WriteLine("DEBUG: Starting face selection...");

                while (true)
                {
                    try
                    {
                        string promptMsg = $"Select Face (Selected: {selectedFaces.Count}) - Press ESC to finish";
                        Reference reference = _uidoc.Selection.PickObject(ObjectType.Face, promptMsg);

                        if (reference != null)
                        {
                            Element element = _doc.GetElement(reference.ElementId);
                            GeometryObject geometryObject = element.GetGeometryObjectFromReference(reference);

                            if (geometryObject is Face face)
                            {
                                var faceId = (element.Id.Value, reference.GlobalPoint.X, reference.GlobalPoint.Y, reference.GlobalPoint.Z);

                                if (!faceIds.Contains(faceId))
                                {
                                    selectedFaces.Add(new FaceData
                                    {
                                        Element = element,
                                        Face = face,
                                        Reference = reference
                                    });
                                    faceIds.Add(faceId);
                                    Debug.WriteLine($"DEBUG: Face selected (Total: {selectedFaces.Count})");
                                }
                            }
                        }
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                    {
                        Debug.WriteLine("DEBUG: User pressed ESC - selection ended");
                        break;
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"DEBUG: Error: {ex.Message}");
                        break;
                    }
                }

                if (selectedFaces.Count == 0)
                {
                    ShowMessage("No faces selected.");
                    return null;
                }

                Debug.WriteLine($"DEBUG: Total faces selected: {selectedFaces.Count}");
                return selectedFaces;
            }
            catch (Exception ex)
            {
                ShowError($"Error selecting faces: {ex.Message}");
                return null;
            }
        }

        // =============================================================================
        // FLOOR TYPE MANAGEMENT
        // =============================================================================

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
                        double thickness = compound.GetWidth();
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

                // Method 2: Get from "Thickness" parameter
                try
                {
                    Parameter param = floorType.LookupParameter("Thickness");
                    if (param != null && param.HasValue && param.AsDouble() > 0)
                    {
                        return param.AsDouble();
                    }
                }
                catch
                {
                    // Continue to next method
                }

                // Method 3: Search for thickness in all parameters
                try
                {
                    IList<Parameter> parameters = floorType.GetOrderedParameters();
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

                // Method 4: Sum layer thicknesses
                try
                {
                    CompoundStructure compound = floorType.GetCompoundStructure();
                    if (compound != null && compound.LayerCount > 0)
                    {
                        double total = 0.0;
                        for (int i = 0; i < compound.LayerCount; i++)
                        {
                            try
                            {
                                total += compound.GetLayerWidth(i);
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

        private Dictionary<string, FloorType> GetAllFloorTypes()
        {
            try
            {
                Debug.WriteLine("DEBUG: get_all_floor_types() started");
                Dictionary<string, FloorType> floorTypes = new Dictionary<string, FloorType>();

                try
                {
                    Debug.WriteLine("DEBUG: Creating FilteredElementCollector...");
                    FilteredElementCollector collector = new FilteredElementCollector(_doc)
                        .OfClass(typeof(FloorType));

                    Debug.WriteLine("DEBUG: Collector created, iterating through floor types...");

                    int count = 0;
                    foreach (FloorType floorType in collector)
                    {
                        count++;
                        try
                        {
                            string name = null;

                            // Get name from Name property
                            try
                            {
                                name = floorType.Name;
                            }
                            catch
                            {
                                // Continue to next method
                            }

                            // Get name from SYMBOL_NAME_PARAM
                            if (string.IsNullOrEmpty(name))
                            {
                                try
                                {
                                    Parameter param = floorType.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM);
                                    if (param != null && param.HasValue)
                                    {
                                        name = param.AsString();
                                    }
                                }
                                catch
                                {
                                    // Continue to next method
                                }
                            }

                            // Fallback name
                            if (string.IsNullOrEmpty(name))
                            {
                                name = $"FloorType_{floorType.Id.Value}";
                            }

                            if (!string.IsNullOrEmpty(name) && !floorTypes.ContainsKey(name))
                            {
                                floorTypes.Add(name, floorType);
                            }
                        }
                        catch (Exception e)
                        {
                            Debug.WriteLine($"DEBUG: Error processing floor type #{count}: {e.Message}");
                            continue;
                        }
                    }

                    Debug.WriteLine($"DEBUG: Total floor types found: {floorTypes.Count}");

                    if (floorTypes.Count > 0)
                    {
                        return floorTypes;
                    }
                    else
                    {
                        Debug.WriteLine("DEBUG: WARNING - No floor types found");
                        return null;
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"DEBUG: Error in collector: {ex.Message}");
                    return null;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"DEBUG: Error in get_all_floor_types: {ex.Message}");
                return null;
            }
        }

        // =============================================================================
        // FLOOR TYPE SELECTION FORM
        // =============================================================================

        private class SelectFloorTypeForm : WinForms.Form
        {
            public FloorType SelectedFloorType { get; private set; }
            public string SelectedFloorTypeName { get; private set; }

            private Dictionary<string, FloorType> _floorTypesDict;
            private WinForms.ComboBox _cmbFloorType;

            public SelectFloorTypeForm(Dictionary<string, FloorType> floorTypesDict)
            {
                _floorTypesDict = floorTypesDict;
                InitializeComponent();
            }

            private void InitializeComponent()
            {
                this.Text = "Select Floor Type";
                this.Width = 500;
                this.Height = 300;
                this.StartPosition = WinForms.FormStartPosition.CenterScreen;
                this.BackColor = Drawing.Color.White;
                this.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;

                // Title label
                WinForms.Label lblTitle = new WinForms.Label
                {
                    Text = "SELECT FLOOR TYPE",
                    Font = new Drawing.Font("Arial", 12, Drawing.FontStyle.Bold),
                    ForeColor = Drawing.Color.DarkBlue,
                    Location = new Drawing.Point(20, 20),
                    Size = new Drawing.Size(450, 25)
                };
                this.Controls.Add(lblTitle);

                // Type label
                WinForms.Label lblType = new WinForms.Label
                {
                    Text = "Floor Type:",
                    Font = new Drawing.Font("Arial", 10),
                    Location = new Drawing.Point(20, 60),
                    Size = new Drawing.Size(100, 20)
                };
                this.Controls.Add(lblType);

                // ComboBox for floor types
                _cmbFloorType = new WinForms.ComboBox
                {
                    Location = new Drawing.Point(130, 60),
                    Size = new Drawing.Size(330, 25),
                    Font = new Drawing.Font("Arial", 9),
                    DropDownStyle = WinForms.ComboBoxStyle.DropDownList
                };

                // Populate ComboBox
                List<string> floorTypeNames = _floorTypesDict.Keys.OrderBy(x => x).ToList();
                foreach (string name in floorTypeNames)
                {
                    FloorType floorType = _floorTypesDict[name];
                    double thickness = GetFloorTypeThicknessStatic(floorType);
                    double thicknessMm = UnitUtils.ConvertFromInternalUnits(thickness, UnitTypeId.Millimeters);

                    string displayName;
                    if (thicknessMm > 0)
                    {
                        displayName = $"{name} ({thicknessMm:F0}mm)";
                    }
                    else
                    {
                        displayName = name;
                    }

                    _cmbFloorType.Items.Add(displayName);
                }

                if (floorTypeNames.Count > 0)
                {
                    _cmbFloorType.SelectedIndex = 0;
                }

                this.Controls.Add(_cmbFloorType);

                // Info label
                WinForms.Label lblInfo = new WinForms.Label
                {
                    Text = "Select a floor type to create floors on selected faces",
                    Font = new Drawing.Font("Arial", 9),
                    ForeColor = Drawing.Color.Gray,
                    Location = new Drawing.Point(20, 95),
                    Size = new Drawing.Size(440, 30)
                };
                this.Controls.Add(lblInfo);

                // OK button
                WinForms.Button btnOk = new WinForms.Button
                {
                    Text = "CREATE FLOORS",
                    Location = new Drawing.Point(120, 140),
                    Size = new Drawing.Size(150, 35),
                    BackColor = Drawing.Color.LightBlue,
                    Font = new Drawing.Font("Arial", 10, Drawing.FontStyle.Bold)
                };
                btnOk.Click += BtnOK_Click;
                this.Controls.Add(btnOk);

                // Cancel button
                WinForms.Button btnCancel = new WinForms.Button
                {
                    Text = "CANCEL",
                    Location = new Drawing.Point(280, 140),
                    Size = new Drawing.Size(100, 35),
                    Font = new Drawing.Font("Arial", 10)
                };
                btnCancel.Click += BtnCancel_Click;
                this.Controls.Add(btnCancel);
            }

            private static double GetFloorTypeThicknessStatic(FloorType floorType)
            {
                // Static version for use in form initialization
                try
                {
                    // Method 1: Get from CompoundStructure
                    try
                    {
                        CompoundStructure compound = floorType.GetCompoundStructure();
                        if (compound != null)
                        {
                            double thickness = compound.GetWidth();
                            if (thickness > 0)
                            {
                                return thickness;
                            }
                        }
                    }
                    catch
                    {
                        // Continue
                    }

                    // Method 2: Get from "Thickness" parameter
                    try
                    {
                        Parameter param = floorType.LookupParameter("Thickness");
                        if (param != null && param.HasValue && param.AsDouble() > 0)
                        {
                            return param.AsDouble();
                        }
                    }
                    catch
                    {
                        // Continue
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
                    WinForms.MessageBox.Show("Please select a floor type!", "Warning");
                    return;
                }

                try
                {
                    string displayName = _cmbFloorType.SelectedItem.ToString();
                    string floorTypeName = displayName.Split(new[] { " (" }, StringSplitOptions.None)[0];

                    if (_floorTypesDict.ContainsKey(floorTypeName))
                    {
                        SelectedFloorType = _floorTypesDict[floorTypeName];
                        SelectedFloorTypeName = floorTypeName;
                        this.DialogResult = WinForms.DialogResult.OK;
                        this.Close();
                    }
                    else
                    {
                        WinForms.MessageBox.Show($"Floor type not found: {floorTypeName}", "Error");
                    }
                }
                catch (Exception ex)
                {
                    WinForms.MessageBox.Show($"Error: {ex.Message}", "Error");
                }
            }

            private void BtnCancel_Click(object sender, EventArgs e)
            {
                this.DialogResult = WinForms.DialogResult.Cancel;
                this.Close();
            }
        }

        // =============================================================================
        // GEOMETRY PROCESSING
        // =============================================================================

        private List<Curve> GetFaceBoundaryCurves(Face face)
        {
            List<Curve> curves = new List<Curve>();
            try
            {
                EdgeArrayArray edgeArrayArray = face.EdgeLoops;
                if (edgeArrayArray.Size == 0)
                {
                    return null;
                }

                EdgeArray outerLoop = edgeArrayArray.get_Item(0);

                for (int i = 0; i < outerLoop.Size; i++)
                {
                    try
                    {
                        Edge edge = outerLoop.get_Item(i);
                        Curve curve = edge.AsCurve();
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

                return curves.Count > 0 ? curves : null;
            }
            catch
            {
                return null;
            }
        }

        private int GetFaceNormalDirection(Face face)
        {
            // Return: 1=TOP (UP), -1=BOTTOM (DOWN), 0=UNKNOWN
            try
            {
                BoundingBoxUV bbox = face.GetBoundingBox();
                if (bbox != null)
                {
                    UV centerUV = new UV((bbox.Min.U + bbox.Max.U) / 2, (bbox.Min.V + bbox.Max.V) / 2);
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

        private List<Curve> OrderCurvesForLoop(List<Curve> curves)
        {
            try
            {
                if (curves == null || curves.Count == 0)
                {
                    return new List<Curve>();
                }

                List<Curve> remaining = new List<Curve>(curves);
                List<Curve> ordered = new List<Curve> { remaining[0] };
                remaining.RemoveAt(0);

                while (remaining.Count > 0)
                {
                    XYZ lastEnd = ordered[ordered.Count - 1].GetEndPoint(1);
                    int bestIdx = -1;
                    bool bestReversed = false;
                    double bestDist = double.MaxValue;

                    for (int i = 0; i < remaining.Count; i++)
                    {
                        Curve curve = remaining[i];
                        double dStart = lastEnd.DistanceTo(curve.GetEndPoint(0));
                        double dEnd = lastEnd.DistanceTo(curve.GetEndPoint(1));

                        if (dStart < bestDist)
                        {
                            bestDist = dStart;
                            bestIdx = i;
                            bestReversed = false;
                        }

                        if (dEnd < bestDist)
                        {
                            bestDist = dEnd;
                            bestIdx = i;
                            bestReversed = true;
                        }
                    }

                    if (bestIdx >= 0 && bestDist < 0.01)
                    {
                        if (bestReversed)
                        {
                            ordered.Add(remaining[bestIdx].CreateReversed());
                        }
                        else
                        {
                            ordered.Add(remaining[bestIdx]);
                        }
                        remaining.RemoveAt(bestIdx);
                    }
                    else
                    {
                        break;
                    }
                }

                return ordered;
            }
            catch
            {
                return curves;
            }
        }

        private CurveLoop CreateCurveLoopFromCurves(List<Curve> curves)
        {
            try
            {
                List<Curve> ordered = OrderCurvesForLoop(curves);
                if (ordered == null || ordered.Count == 0)
                {
                    return null;
                }

                List<Curve> final = new List<Curve>(ordered);
                if (final.Count > 1)
                {
                    XYZ lastEnd = final[final.Count - 1].GetEndPoint(1);
                    XYZ firstStart = final[0].GetEndPoint(0);
                    if (lastEnd.DistanceTo(firstStart) > 1e-6)
                    {
                        final.Add(Line.CreateBound(lastEnd, firstStart));
                    }
                }

                CurveLoop curveLoop = CurveLoop.Create(final);
                return curveLoop != null && !curveLoop.IsOpen() ? curveLoop : null;
            }
            catch
            {
                return null;
            }
        }

        // =============================================================================
        // FLOOR CREATION
        // =============================================================================

        private Floor CreateFloorOnFace(FaceData faceData, FloorType floorType, double thicknessInternal)
        {
            try
            {
                Face face = faceData.Face;

                int normalDir = GetFaceNormalDirection(face);
                if (normalDir == 0)
                {
                    Debug.WriteLine("DEBUG: Unknown face orientation");
                    return null;
                }

                List<Curve> curves = GetFaceBoundaryCurves(face);
                if (curves == null || curves.Count == 0)
                {
                    return null;
                }

                CurveLoop curveLoop = CreateCurveLoopFromCurves(curves);
                if (curveLoop == null)
                {
                    return null;
                }

                double faceElev = faceData.Reference.GlobalPoint.Z;

                List<Level> levels = new FilteredElementCollector(_doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .ToList();

                if (levels.Count == 0)
                {
                    return null;
                }

                Level closestLevel = levels.OrderBy(l => Math.Abs(l.Elevation - faceElev)).First();
                double baseOffset = faceElev - closestLevel.Elevation;

                double finalOffset;
                if (normalDir == 1)
                {
                    finalOffset = baseOffset + thicknessInternal;
                    Debug.WriteLine($"DEBUG: TOP FACE - offset = base + thickness = {finalOffset:F6}");
                }
                else
                {
                    finalOffset = baseOffset;
                    Debug.WriteLine($"DEBUG: BOTTOM FACE - offset = base only = {finalOffset:F6}");
                }

                Floor floor = Floor.Create(_doc, new List<CurveLoop> { curveLoop }, floorType.Id, closestLevel.Id);

                if (floor != null)
                {
                    try
                    {
                        Parameter param = floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
                        if (param != null && !param.IsReadOnly)
                        {
                            param.Set(finalOffset);
                        }
                    }
                    catch
                    {
                        // Ignore if cannot set offset
                    }
                    return floor;
                }

                return null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"DEBUG: Error creating floor: {ex.Message}");
                return null;
            }
        }

        // =============================================================================
        // MAIN EXECUTION
        // =============================================================================

        private Result RunMain()
        {
            try
            {
                Debug.WriteLine("DEBUG: run() started");
                ShowMessage("Select Face(s) in 3D view\nTOP FACE: offset = base + thickness\nBOTTOM FACE: offset = base only");

                Debug.WriteLine("DEBUG: About to pick faces...");
                List<FaceData> facesData = PickMultipleFaces();

                if (facesData == null || facesData.Count == 0)
                {
                    Debug.WriteLine("DEBUG: No faces selected");
                    return Result.Cancelled;
                }

                Debug.WriteLine("DEBUG: About to get floor types...");
                Dictionary<string, FloorType> floorTypes = GetAllFloorTypes();

                if (floorTypes == null || floorTypes.Count == 0)
                {
                    ShowError("No floor types found in project");
                    return Result.Failed;
                }

                Debug.WriteLine("DEBUG: Showing floor type selection form...");
                using (SelectFloorTypeForm form = new SelectFloorTypeForm(floorTypes))
                {
                    WinForms.DialogResult result = form.ShowDialog();

                    if (result != WinForms.DialogResult.OK)
                    {
                        Debug.WriteLine("DEBUG: User cancelled floor type selection");
                        ShowMessage("Cancelled");
                        return Result.Cancelled;
                    }

                    FloorType floorType = form.SelectedFloorType;
                    string floorTypeName = form.SelectedFloorTypeName;
                    Debug.WriteLine($"DEBUG: Selected floor type: {floorTypeName}");

                    Debug.WriteLine("DEBUG: Getting thickness...");
                    double thickness = GetFloorTypeThickness(floorType);
                    double thicknessMm = UnitUtils.ConvertFromInternalUnits(thickness, UnitTypeId.Millimeters);
                    Debug.WriteLine($"DEBUG: Thickness: {thicknessMm:F2}mm");

                    ShowMessage($"Floor Type: {floorTypeName}\nThickness: {thicknessMm:F2}mm\n\nCreating {facesData.Count} floors...");

                    Debug.WriteLine("DEBUG: Starting transaction...");
                    using (Transaction trans = new Transaction(_doc, "Create Floors from Faces"))
                    {
                        trans.Start();

                        try
                        {
                            int created = 0;
                            for (int idx = 0; idx < facesData.Count; idx++)
                            {
                                Debug.WriteLine($"DEBUG: Creating floor {idx + 1}...");
                                Floor floor = CreateFloorOnFace(facesData[idx], floorType, thickness);
                                if (floor != null)
                                {
                                    created++;
                                    Debug.WriteLine($"DEBUG: Floor {idx + 1} created successfully");
                                }
                                else
                                {
                                    Debug.WriteLine($"DEBUG: Floor {idx + 1} failed");
                                }
                            }

                            Debug.WriteLine("DEBUG: Committing transaction...");
                            trans.Commit();
                            Debug.WriteLine("DEBUG: Transaction committed");

                            string msg = "========== SUCCESS ==========\n\n";
                            msg += $"Created: {created}/{facesData.Count} floors\n\n";
                            msg += $"Floor Type: {floorTypeName}\n";
                            msg += $"Thickness: {thicknessMm:F2}mm\n\n";
                            msg += "OFFSET LOGIC:\n";
                            msg += "• TOP FACE: offset = base + thickness\n";
                            msg += "• BOTTOM FACE: offset = base only";

                            ShowMessage(msg);
                            Debug.WriteLine("DEBUG: run() completed successfully");
                            return Result.Succeeded;
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"DEBUG: Exception in transaction: {ex.Message}");
                            trans.RollBack();
                            ShowError($"Error: {ex.Message}");
                            return Result.Failed;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"DEBUG: Exception in run(): {ex.Message}");
                ShowError($"Error: {ex.Message}");
                return Result.Failed;
            }
        }

        // =============================================================================
        // DEBUG UTILITY
        // =============================================================================

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