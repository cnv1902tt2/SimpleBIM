using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Text;

// WINDOWS FORMS NAMESPACES
using System.Windows.Forms;
using System.Drawing;

// AUTODESK REVIT API NAMESPACES
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.Exceptions;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;

// ============================================================================
// ALIASES ĐỂ TRÁNH AMBIGUOUS REFERENCES
// ============================================================================

using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;

using RevitApplication = Autodesk.Revit.ApplicationServices.Application;
using RevitDocument = Autodesk.Revit.DB.Document;
using RevitUIApplication = Autodesk.Revit.UI.UIApplication;
using RevitUIDocument = Autodesk.Revit.UI.UIDocument;
using RevitElement = Autodesk.Revit.DB.Element;
using RevitTransaction = Autodesk.Revit.DB.Transaction;
using OperationCanceledException = Autodesk.Revit.Exceptions.OperationCanceledException;

namespace SimpleBIM.Commands.Qs
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class DeleteDuplicates : IExternalCommand
    {
        private RevitDocument _doc;
        private RevitUIDocument _uidoc;
        private Dictionary<string, List<RevitElement>> _elementsBySignature = new Dictionary<string, List<RevitElement>>();

        private static readonly List<string> TARGET_CATEGORY_NAMES = new List<string>
        {
            "Walls", "Floors", "Structural Framing", "Structural Columns", "Columns",
            "Pipes", "Ducts", "Conduits", "Cable Trays", "Generic Models",
            "Furniture", "Plumbing Fixtures", "Lighting Fixtures",
            "Mechanical Equipment", "Electrical Equipment"
        };

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            RevitUIApplication uiapp = commandData.Application;
            _uidoc = uiapp.ActiveUIDocument;
            _doc = _uidoc.Document;

            try
            {
                Debug.WriteLine("🔄 Đang thu thập đối tượng...");
                Stopwatch timer = Stopwatch.StartNew();

                // Thu thập elements
                List<RevitElement> elementsWithLocation = GatherElementsWithLocation();
                timer.Stop();

                Debug.WriteLine($"✅ Đã thu thập {elementsWithLocation.Count} đối tượng (trong {timer.ElapsedMilliseconds}ms)");

                if (elementsWithLocation.Count == 0)
                {
                    TaskDialog.Show("Warning", "Không có đối tượng nào có Location để kiểm tra!");
                    return Result.Cancelled;
                }

                // Phân nhóm by signature
                Debug.WriteLine("🔍 Đang phân tích duplicates...");
                foreach (RevitElement element in elementsWithLocation)
                {
                    string signature = GetElementSignature(element);
                    if (!_elementsBySignature.ContainsKey(signature))
                    {
                        _elementsBySignature[signature] = new List<RevitElement>();
                    }
                    _elementsBySignature[signature].Add(element);
                }

                // Tìm duplicates
                List<RevitElement> duplicatesToDelete = new List<RevitElement>();
                foreach (var group in _elementsBySignature)
                {
                    if (group.Value.Count > 1)
                    {
                        // Giữ cái đầu tiên, xóa những cái còn lại
                        for (int i = 1; i < group.Value.Count; i++)
                        {
                            duplicatesToDelete.Add(group.Value[i]);
                        }
                    }
                }

                if (duplicatesToDelete.Count == 0)
                {
                    TaskDialog.Show("Info", "Không phát hiện đối tượng trùng lặp!");
                    return Result.Succeeded;
                }

                // Hiển thị summary
                string summary = $"Phát hiện {duplicatesToDelete.Count} đối tượng trùng lặp.\n\n" +
                    $"Bạn có muốn xóa chúng không?";

                TaskDialogResult result = TaskDialog.Show(
                    "Delete Duplicates",
                    summary,
                    TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

                if (result != TaskDialogResult.Yes)
                {
                    return Result.Cancelled;
                }

                // Xóa duplicates
                using (RevitTransaction trans = new RevitTransaction(_doc, "Delete Duplicates"))
                {
                    trans.Start();

                    try
                    {
                        int successCount = 0;
                        List<string> errors = new List<string>();

                        foreach (RevitElement element in duplicatesToDelete)
                        {
                            try
                            {
                                _doc.Delete(element.Id);
                                successCount++;
                            }
                            catch (Exception ex)
                            {
                                errors.Add($"Element {element.Id}: {ex.Message}");
                            }
                        }

                        trans.Commit();

                        string resultMsg = $"Đã xóa {successCount} đối tượng trùng lặp.";
                        if (errors.Count > 0)
                        {
                            resultMsg += $"\n\nLỗi: {errors.Count}";
                        }

                        TaskDialog.Show("Delete Complete", resultMsg);
                        return Result.Succeeded;
                    }
                    catch (Exception ex)
                    {
                        trans.RollBack();
                        TaskDialog.Show("Error", $"Lỗi khi xóa:\n{ex.Message}");
                        return Result.Failed;
                    }
                }
            }
            catch (OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                Debug.WriteLine($"Error in DeleteDuplicates: {ex.Message}\n{ex.StackTrace}");
                return Result.Failed;
            }
        }

        private List<RevitElement> GatherElementsWithLocation()
        {
            FilteredElementCollector collector = new FilteredElementCollector(_doc)
                .WhereElementIsNotElementType();

            List<RevitElement> result = new List<RevitElement>();

            foreach (RevitElement element in collector)
            {
                try
                {
                    if (element.Category == null || !TARGET_CATEGORY_NAMES.Contains(element.Category.Name))
                    {
                        continue;
                    }

                    if (element.Location != null && element.Id.IntegerValue > 0)
                    {
                        result.Add(element);
                    }
                }
                catch
                {
                    continue;
                }
            }

            return result;
        }

        private string GetElementSignature(RevitElement element)
        {
            try
            {
                // Lấy location
                LocationPoint locPoint = element.Location as LocationPoint;
                LocationCurve locCurve = element.Location as LocationCurve;

                XYZ basePoint = null;
                if (locPoint != null)
                {
                    basePoint = locPoint.Point;
                }
                else if (locCurve != null)
                {
                    basePoint = locCurve.Curve.GetEndPoint(0);
                }

                if (basePoint == null)
                {
                    return $"UNKNOWN_{element.Id}";
                }

                // Signature cơ bản
                string signature = $"{element.Category.Name}|" +
                    $"{Math.Round(basePoint.X, 3)}|" +
                    $"{Math.Round(basePoint.Y, 3)}|" +
                    $"{Math.Round(basePoint.Z, 3)}";

                // Thêm thông tin loại element
                if (element is Wall)
                {
                    Wall wall = (Wall)element;
                    if (wall.WallType != null)
                    {
                        signature += $"|{wall.WallType.Name}";
                    }
                    Parameter heightParam = element.LookupParameter("Unconnected Height");
                    if (heightParam != null && heightParam.HasValue)
                    {
                        signature += $"|{Math.Round(heightParam.AsDouble(), 3)}";
                    }
                }
                else if (element is Floor)
                {
                    Floor floor = (Floor)element;
                    if (floor.FloorType != null)
                    {
                        signature += $"|{floor.FloorType.Name}";
                    }
                    Parameter elevParam = element.LookupParameter("Elevation");
                    if (elevParam != null && elevParam.HasValue)
                    {
                        signature += $"|{Math.Round(elevParam.AsDouble(), 3)}";
                    }
                }
                else if (element is Pipe)
                {
                    Pipe pipe = (Pipe)element;
                    if (pipe.PipeType != null)
                    {
                        signature += $"|{pipe.PipeType.Name}";
                    }
                    Parameter diamParam = element.LookupParameter("Diameter");
                    if (diamParam != null && diamParam.HasValue)
                    {
                        signature += $"|{Math.Round(diamParam.AsDouble(), 3)}";
                    }
                }
                else if (element is Duct)
                {
                    Duct duct = (Duct)element;
                    if (duct.DuctType != null)
                    {
                        signature += $"|{duct.DuctType.Name}";
                    }
                    Parameter widthParam = element.LookupParameter("Width");
                    Parameter heightParam = element.LookupParameter("Height");
                    if (widthParam != null && widthParam.HasValue && heightParam != null && heightParam.HasValue)
                    {
                        signature += $"|{Math.Round(widthParam.AsDouble(), 3)}|{Math.Round(heightParam.AsDouble(), 3)}";
                    }
                }
                else
                {
                    // Generic elements
                    Parameter typeParam = element.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM);
                    if (typeParam != null && typeParam.HasValue)
                    {
                        signature += $"|{typeParam.AsValueString()}";
                    }
                }

                return signature;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error getting signature for element {element.Id}: {ex.Message}");
                return $"ERROR_{element.Id}_{DateTime.Now.Ticks}";
            }
        }
    }
}
