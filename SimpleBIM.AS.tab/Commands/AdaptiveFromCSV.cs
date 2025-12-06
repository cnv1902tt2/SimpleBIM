using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

//===================== HOÀN CHỈNH =====================

namespace SimpleBIM.AS.tab.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class AdaptiveFromCSV : IExternalCommand
    {
        // Constants
        private const string CSV_FILTER = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*";

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            Autodesk.Revit.ApplicationServices.Application revitApp = uiapp.Application;

            try
            {
                // Step 1: Get all adaptive families
                List<Family> adaptiveFamilies = GetAdaptiveFamilies(doc);

                if (adaptiveFamilies == null || adaptiveFamilies.Count == 0)
                {
                    TaskDialog.Show("Error",
                        "No adaptive families found in the project.\n\n" +
                        "Please load at least one adaptive component family before running this command.");
                    return Result.Failed;
                }

                // Step 2: Let user select a family
                FamilySelectionForm familyForm = new FamilySelectionForm(adaptiveFamilies);
                System.Windows.Forms.DialogResult formResult = familyForm.ShowDialog();

                if (formResult != System.Windows.Forms.DialogResult.OK || familyForm.SelectedFamily == null)
                {
                    return Result.Cancelled;
                }

                Family selectedFamily = familyForm.SelectedFamily;

                // Step 3: Get the family symbol
                FamilySymbol symbol = GetFirstFamilySymbol(doc, selectedFamily);
                if (symbol == null)
                {
                    TaskDialog.Show("Error",
                        $"No family symbols found for '{selectedFamily.Name}'.");
                    return Result.Failed;
                }

                // Step 4: Determine number of adaptive points
                int requiredPoints = GetAdaptivePointCount(symbol);
                if (requiredPoints <= 0)
                {
                    TaskDialog.Show("Error",
                        $"Unable to determine the number of adaptive points for '{selectedFamily.Name}'.");
                    return Result.Failed;
                }

                // Step 5: Select CSV file
                string csvPath = SelectCsvFile();
                if (string.IsNullOrEmpty(csvPath))
                {
                    return Result.Cancelled;
                }

                // Step 6: Parse CSV file
                List<List<XYZ>> pointGroups = ParseCsvFile(csvPath, requiredPoints);
                if (pointGroups == null || pointGroups.Count == 0)
                {
                    TaskDialog.Show("Error",
                        $"No valid point groups found in CSV file.\n" +
                        $"Expected {requiredPoints} points per adaptive component.");
                    return Result.Failed;
                }

                // Step 7: Show confirmation dialog
                TaskDialogResult confirmation = TaskDialog.Show("Confirmation",
                    $"Found {pointGroups.Count} point groups in CSV file.\n" +
                    $"Each group has {requiredPoints} points.\n" +
                    $"Family: {selectedFamily.Name}\n\n" +
                    $"Proceed with creating {pointGroups.Count} adaptive components?",
                    TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

                if (confirmation != TaskDialogResult.Yes)
                {
                    return Result.Cancelled;
                }

                // Step 8: Create adaptive components
                int successCount = CreateAdaptiveComponents(doc, symbol, pointGroups);

                // Step 9: Show results
                TaskDialog.Show("Success",
                    $"Successfully created {successCount} out of {pointGroups.Count} adaptive components.");

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // User cancelled operation
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                // Detailed error reporting
                message = $"Error: {ex.Message}\n\nStack Trace:\n{ex.StackTrace}";
                TaskDialog.Show("Critical Error",
                    $"An error occurred:\n\n{ex.Message}\n\n" +
                    "Please check the CSV format and try again.");
                return Result.Failed;
            }
        }

        /// <summary>
        /// Retrieves all adaptive families from the document
        /// </summary>
        private List<Family> GetAdaptiveFamilies(Document doc)
        {
            List<Family> adaptiveFamilies = new List<Family>();

            try
            {
                // Collect all families
                FilteredElementCollector familyCollector = new FilteredElementCollector(doc)
                    .OfClass(typeof(Family));

                foreach (Family family in familyCollector.Cast<Family>())
                {
                    try
                    {
                        // Check if family is adaptive
                        if (family.FamilyPlacementType == FamilyPlacementType.Adaptive)
                        {
                            // Verify it has at least one symbol
                            ISet<ElementId> symbolIds = family.GetFamilySymbolIds();
                            if (symbolIds != null && symbolIds.Count > 0)
                            {
                                adaptiveFamilies.Add(family);
                            }
                        }
                    }
                    catch
                    {
                        // Skip families that cause errors
                        continue;
                    }
                }

                return adaptiveFamilies.OrderBy(f => f.Name).ToList();
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Warning",
                    $"Error collecting adaptive families: {ex.Message}");
                return adaptiveFamilies;
            }
        }

        /// <summary>
        /// Gets the first available family symbol for a family
        /// </summary>
        private FamilySymbol GetFirstFamilySymbol(Document doc, Family family)
        {
            try
            {
                ISet<ElementId> symbolIds = family.GetFamilySymbolIds();
                if (symbolIds == null || symbolIds.Count == 0)
                {
                    return null;
                }

                // Get the first symbol ID
                ElementId firstSymbolId = symbolIds.First();
                FamilySymbol symbol = doc.GetElement(firstSymbolId) as FamilySymbol;

                return symbol;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Determines the number of adaptive points for a family symbol
        /// </summary>
        private int GetAdaptivePointCount(FamilySymbol symbol)
        {
            // Method 1: Try to get from family API
            try
            {
                Document doc = symbol.Document;

                // Create a temporary transaction to examine the symbol
                using (Transaction tempTrans = new Transaction(doc, "Temp Examine Symbol"))
                {
                    tempTrans.Start();

                    // Try to create a temporary instance
                    FamilyInstance tempInstance = null;
                    try
                    {
                        // Activate symbol if needed
                        if (!symbol.IsActive)
                        {
                            symbol.Activate();
                        }

                        // Create adaptive instance
                        tempInstance = AdaptiveComponentInstanceUtils.CreateAdaptiveComponentInstance(doc, symbol);

                        if (tempInstance != null)
                        {
                            // Get placement points
                            IList<ElementId> pointIds = AdaptiveComponentInstanceUtils
                                .GetInstancePlacementPointElementRefIds(tempInstance);

                            int pointCount = pointIds?.Count ?? 0;

                            // Rollback transaction (deletes temp instance)
                            tempTrans.RollBack();

                            return pointCount;
                        }
                    }
                    catch
                    {
                        // Ensure transaction is rolled back
                        if (tempTrans.HasStarted() && !tempTrans.HasEnded())
                        {
                            tempTrans.RollBack();
                        }
                    }
                }
            }
            catch
            {
                // Continue to fallback method
            }

            // Method 2: Fallback - default to 4 points (common for adaptive components)
            return 4;
        }

        /// <summary>
        /// Opens file dialog to select CSV file
        /// </summary>
        private string SelectCsvFile()
        {
            using (OpenFileDialog openDialog = new OpenFileDialog())
            {
                openDialog.Title = "Select CSV File with Point Coordinates";
                openDialog.Filter = CSV_FILTER;
                openDialog.InitialDirectory = Environment.GetFolderPath(
                    Environment.SpecialFolder.MyDocuments);
                openDialog.CheckFileExists = true;
                openDialog.Multiselect = false;

                if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    return openDialog.FileName;
                }
            }

            return null;
        }

        /// <summary>
        /// Parses CSV file and groups points
        /// </summary>
        private List<List<XYZ>> ParseCsvFile(string filePath, int pointsPerGroup)
        {
            List<List<XYZ>> pointGroups = new List<List<XYZ>>();
            List<XYZ> currentGroup = new List<XYZ>();
            int lineNumber = 0;

            try
            {
                // Read all lines
                string[] lines = File.ReadAllLines(filePath);

                for (int i = 0; i < lines.Length; i++)
                {
                    lineNumber = i + 1;
                    string line = lines[i].Trim();

                    // Skip empty lines and potential headers
                    if (string.IsNullOrEmpty(line) ||
                        line.StartsWith("#") ||
                        line.ToLower().Contains("x,y,z"))
                    {
                        continue;
                    }

                    // Split line by comma, handling quoted values
                    string[] parts = line.Split(',');
                    if (parts.Length < 3)
                    {
                        // Try semicolon separator
                        parts = line.Split(';');
                        if (parts.Length < 3)
                        {
                            continue; // Skip invalid lines
                        }
                    }

                    // Parse coordinates
                    if (TryParseCoordinates(parts, out XYZ point))
                    {
                        currentGroup.Add(point);

                        // Check if group is complete
                        if (currentGroup.Count == pointsPerGroup)
                        {
                            pointGroups.Add(new List<XYZ>(currentGroup));
                            currentGroup.Clear();
                        }
                    }
                }

                // Handle any incomplete group at the end
                if (currentGroup.Count > 0 && currentGroup.Count < pointsPerGroup)
                {
                    TaskDialog.Show("Warning",
                        $"Found incomplete point group at end of file.\n" +
                        $"Expected {pointsPerGroup} points, but found {currentGroup.Count}.\n" +
                        "This group will be ignored.");
                }

                return pointGroups;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Parse Error",
                    $"Error parsing CSV file at line {lineNumber}:\n{ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Attempts to parse XYZ coordinates from string array
        /// </summary>
        private bool TryParseCoordinates(string[] parts, out XYZ point)
        {
            point = null;

            if (parts == null || parts.Length < 3)
                return false;

            try
            {
                // Try parsing each coordinate
                if (!double.TryParse(parts[0].Trim(), System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out double x))
                    return false;

                if (!double.TryParse(parts[1].Trim(), System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out double y))
                    return false;

                if (!double.TryParse(parts[2].Trim(), System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out double z))
                    return false;

                // Convert from millimeters to feet (Revit internal units)
                x = UnitUtils.ConvertToInternalUnits(x, UnitTypeId.Millimeters);
                y = UnitUtils.ConvertToInternalUnits(y, UnitTypeId.Millimeters);
                z = UnitUtils.ConvertToInternalUnits(z, UnitTypeId.Millimeters);

                point = new XYZ(x, y, z);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Creates adaptive components from point groups
        /// </summary>
        private int CreateAdaptiveComponents(Document doc, FamilySymbol symbol, List<List<XYZ>> pointGroups)
        {
            int successCount = 0;
            int totalGroups = pointGroups.Count;

            using (Transaction trans = new Transaction(doc, "Create Adaptive Components from CSV"))
            {
                trans.Start();

                try
                {
                    // Ensure symbol is active
                    if (!symbol.IsActive)
                    {
                        symbol.Activate();
                    }

                    // Create components for each point group
                    for (int groupIndex = 0; groupIndex < totalGroups; groupIndex++)
                    {
                        try
                        {
                            List<XYZ> points = pointGroups[groupIndex];

                            // Create adaptive component instance
                            FamilyInstance instance = AdaptiveComponentInstanceUtils
                                .CreateAdaptiveComponentInstance(doc, symbol);

                            if (instance == null)
                            {
                                continue;
                            }

                            // Get placement points
                            IList<ElementId> placementPointIds = AdaptiveComponentInstanceUtils
                                .GetInstancePlacementPointElementRefIds(instance);

                            // Set positions of placement points
                            int pointsToPlace = Math.Min(placementPointIds.Count, points.Count);

                            for (int pointIndex = 0; pointIndex < pointsToPlace; pointIndex++)
                            {
                                ReferencePoint refPoint = doc.GetElement(placementPointIds[pointIndex]) as ReferencePoint;
                                if (refPoint != null)
                                {
                                    refPoint.Position = points[pointIndex];
                                }
                            }

                            successCount++;
                        }
                        catch
                        {
                            // Continue with other groups
                        }
                    }

                    trans.Commit();
                }
                catch (Exception ex)
                {
                    trans.RollBack();
                    throw new Exception($"Failed to create adaptive components: {ex.Message}", ex);
                }
            }

            return successCount;
        }

        /// <summary>
        /// Helper class for debugging
        /// </summary>
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

    /// <summary>
    /// Simple Windows Forms dialog for selecting an adaptive family
    /// </summary>
    internal class FamilySelectionForm : System.Windows.Forms.Form
    {
        public Family SelectedFamily { get; private set; }
        private System.Windows.Forms.ListBox listBoxFamilies;
        private List<Family> families;

        public FamilySelectionForm(List<Family> adaptiveFamilies)
        {
            families = adaptiveFamilies ?? new List<Family>();
            InitializeComponent();
            PopulateListBox();
        }

        private void InitializeComponent()
        {
            // Form properties
            this.Text = "Select Adaptive Family";
            this.Width = 450;
            this.Height = 500;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Title label
            System.Windows.Forms.Label titleLabel = new System.Windows.Forms.Label
            {
                Text = "SELECT ADAPTIVE FAMILY",
                Font = new System.Drawing.Font("Arial", 12, System.Drawing.FontStyle.Bold),
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(400, 30),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };
            this.Controls.Add(titleLabel);

            // Instructions label
            System.Windows.Forms.Label instructionLabel = new System.Windows.Forms.Label
            {
                Text = "Select a family from the list below:",
                Font = new System.Drawing.Font("Arial", 9),
                Location = new System.Drawing.Point(20, 55),
                Size = new System.Drawing.Size(400, 20)
            };
            this.Controls.Add(instructionLabel);

            // ListBox for families
            listBoxFamilies = new System.Windows.Forms.ListBox
            {
                Location = new System.Drawing.Point(20, 80),
                Size = new System.Drawing.Size(400, 300),
                Font = new System.Drawing.Font("Consolas", 9),
                SelectionMode = System.Windows.Forms.SelectionMode.One
            };
            listBoxFamilies.DoubleClick += ListBoxFamilies_DoubleClick;
            this.Controls.Add(listBoxFamilies);

            // OK button
            System.Windows.Forms.Button btnOk = new System.Windows.Forms.Button
            {
                Text = "SELECT",
                Location = new System.Drawing.Point(175, 390),
                Size = new System.Drawing.Size(100, 30)
            };
            btnOk.Click += BtnOk_Click;
            this.Controls.Add(btnOk);

            // Cancel button
            System.Windows.Forms.Button btnCancel = new System.Windows.Forms.Button
            {
                Text = "CANCEL",
                Location = new System.Drawing.Point(175, 430),
                Size = new System.Drawing.Size(100, 30)
            };
            btnCancel.Click += BtnCancel_Click;
            this.Controls.Add(btnCancel);
        }

        private void PopulateListBox()
        {
            listBoxFamilies.Items.Clear();

            foreach (Family family in families.OrderBy(f => f.Name))
            {
                // Get symbol count for display
                int symbolCount = 0;
                try
                {
                    ISet<ElementId> symbolIds = family.GetFamilySymbolIds();
                    symbolCount = symbolIds?.Count ?? 0;
                }
                catch
                {
                    symbolCount = 0;
                }

                string displayText = $"{family.Name.PadRight(40)} (Symbols: {symbolCount})";
                listBoxFamilies.Items.Add(displayText);
            }

            // Select first item if available
            if (listBoxFamilies.Items.Count > 0)
            {
                listBoxFamilies.SelectedIndex = 0;
            }
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            if (listBoxFamilies.SelectedIndex >= 0 && listBoxFamilies.SelectedIndex < families.Count)
            {
                SelectedFamily = families[listBoxFamilies.SelectedIndex];
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
                this.Close();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Please select a family from the list.",
                    "Selection Required",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }

        private void ListBoxFamilies_DoubleClick(object sender, EventArgs e)
        {
            BtnOk_Click(sender, e);
        }
    }
}