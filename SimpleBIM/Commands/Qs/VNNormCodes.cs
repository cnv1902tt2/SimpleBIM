using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Media.Media3D;
using static System.Net.Mime.MediaTypeNames;


using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;
using RevitApplication = Autodesk.Revit.ApplicationServices;
namespace GanMaDinhMuc_Hybrid


{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class MainCommand : IExternalCommand
    {
        private static UIApplication _uiapp;
        private static Document _doc;
        private static UIDocument _uidoc;
        private static RevitApplication.Application _app;
        private static int _revitVersion;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            _uiapp = commandData.Application;
            _doc = _uiapp.ActiveUIDocument.Document;
            _uidoc = _uiapp.ActiveUIDocument;
            _app = _doc.Application;
            _revitVersion = int.Parse(_app.VersionNumber);

            try
            {
                Debug.WriteLine("=".PadRight(70, '='));
                Debug.WriteLine("TOOL GÁN MÃ ĐỊNH MỨC v5.0 + v6.0 HYBRID - FIXED VERSION");
                Debug.WriteLine("=".PadRight(70, '='));

                // STEP 1: SELECT CSV FILE
                Debug.WriteLine("\n[STEP 1] Selecting CSV file...");
                string csvFile = SelectCsvFile();
                if (string.IsNullOrEmpty(csvFile))
                {
                    Debug.WriteLine("User cancelled CSV file selection");
                    TaskDialog.Show("Thông báo", "Bạn đã hủy chọn file!");
                    return Result.Cancelled;
                }

                Debug.WriteLine($"CSV file selected: {csvFile}");

                // STEP 2: LOAD CSV DATA
                Debug.WriteLine("\n[STEP 2] Loading CSV data...");
                try
                {
                    List<Dictionary<string, string>> csvData = ReadCsvData(csvFile);
                    if (csvData == null || csvData.Count == 0)
                    {
                        Debug.WriteLine("ERROR: CSV file is empty or invalid");
                        TaskDialog.Show("Lỗi", "File CSV rỗng hoặc không có dữ liệu!");
                        return Result.Failed;
                    }
                    Debug.WriteLine($"CSV data loaded: {csvData.Count} rows");

                    // Print first row as sample
                    if (csvData.Count > 0)
                    {
                        var firstRow = csvData[0];
                        Debug.WriteLine($"Sample row: {string.Join(", ", firstRow.Select(kv => $"{kv.Key}:{kv.Value}"))}");
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"ERROR loading CSV: {ex.Message}");
                    TaskDialog.Show("Lỗi", $"Lỗi khi đọc CSV: {ex.Message}");
                    return Result.Failed;
                }

                // STEP 3: GET CATEGORIES FROM PROJECT
                Debug.WriteLine("\n[STEP 3] Loading categories from project...");
                try
                {
                    Dictionary<string, CategoryInfo> categoriesDict = GetAllCategoriesFromProject();
                    Debug.WriteLine($"Categories found: {categoriesDict.Count}");

                    foreach (var kvp in categoriesDict)
                    {
                        Debug.WriteLine($"  - {kvp.Key}: {kvp.Value.Count} elements ({kvp.Value.Type})");
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"ERROR loading categories: {ex.Message}");
                    TaskDialog.Show("Lỗi", $"Lỗi khi tải categories: {ex.Message}");
                    return Result.Failed;
                }

                // STEP 4: CREATE AND SHOW FORM
                Debug.WriteLine("\n[STEP 4] Creating main form...");
                try
                {
                    var csvData = ReadCsvData(csvFile);
                    var categoriesDict = GetAllCategoriesFromProject();

                    using (MarkAssignmentForm form = new MarkAssignmentForm(csvData, categoriesDict))
                    {
                        Debug.WriteLine("Form created successfully");
                        Debug.WriteLine($"Form size: {form.Width} x {form.Height}");
                        Debug.WriteLine("Showing form...");

                        DialogResult result = form.ShowDialog();
                        Debug.WriteLine($"Form closed with result: {result}");
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"ERROR creating form: {ex.Message}");
                    Debug.WriteLine($"Traceback: {ex.StackTrace}");
                    TaskDialog.Show("Lỗi", $"Lỗi form: {ex.Message}");
                    return Result.Failed;
                }

                Debug.WriteLine("\n" + "=".PadRight(70, '='));
                Debug.WriteLine("TOOL COMPLETED");
                Debug.WriteLine("=".PadRight(70, '='));

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"FATAL ERROR: {ex.Message}");
                Debug.WriteLine($"Traceback: {ex.StackTrace}");
                TaskDialog.Show("Lỗi", $"Lỗi: {ex.Message}");
                return Result.Failed;
            }
        }

        // =============================================================================
        // HELPER FUNCTIONS
        // =============================================================================

        private static string SelectCsvFile()
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "CSV Files (*.csv)|*.csv|All Files|*.*";
                dlg.Title = "Chon file CSV";
                dlg.CheckFileExists = true;
                dlg.CheckPathExists = true;

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    return dlg.FileName;
                }
                return null;
            }
        }

        private static List<Dictionary<string, string>> ReadCsvData(string filePath)
        {
            List<Dictionary<string, string>> data = new List<Dictionary<string, string>>();
            Encoding[] encodings = {
                new UTF8Encoding(true), // utf-8-sig
                Encoding.UTF8,          // utf-8
                Encoding.GetEncoding(1258), // cp1258
                Encoding.GetEncoding("iso-8859-1") // iso-8859-1
            };

            foreach (var encoding in encodings)
            {
                try
                {
                    string content = File.ReadAllText(filePath, encoding).Trim();
                    if (string.IsNullOrEmpty(content)) continue;

                    string[] lines = content.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    if (lines.Length == 0) continue;

                    // Parse CSV
                    string[] headers = ParseCsvLine(lines[0]);
                    for (int i = 1; i < lines.Length; i++)
                    {
                        string[] values = ParseCsvLine(lines[i]);
                        if (values.Length > 0)
                        {
                            Dictionary<string, string> row = new Dictionary<string, string>();
                            for (int j = 0; j < Math.Min(headers.Length, values.Length); j++)
                            {
                                string key = headers[j].Trim();
                                string value = values[j].Trim();
                                if (!string.IsNullOrEmpty(key))
                                {
                                    row[key] = value;
                                }
                            }

                            if (row.ContainsKey("Mã Hiệu") && !string.IsNullOrEmpty(row["Mã Hiệu"]))
                            {
                                data.Add(row);
                            }
                        }
                    }

                    if (data.Count > 0) break;
                }
                catch
                {
                    continue;
                }
            }

            return data;
        }

        private static string[] ParseCsvLine(string line)
        {
            List<string> result = new List<string>();
            bool inQuotes = false;
            StringBuilder current = new StringBuilder();

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '"')
                {
                    inQuotes = !inQuotes;
                }
                else if (c == ',' && !inQuotes)
                {
                    result.Add(current.ToString());
                    current.Clear();
                }
                else
                {
                    current.Append(c);
                }
            }

            result.Add(current.ToString());
            return result.ToArray();
        }

        private static Dictionary<string, CategoryInfo> GetAllCategoriesFromProject()
        {
            Dictionary<string, CategoryInfo> categoriesDict = new Dictionary<string, CategoryInfo>();
            HashSet<string> tempCategories = new HashSet<string>();

            FilteredElementCollector collector = new FilteredElementCollector(_doc)
                .WhereElementIsNotElementType();

            foreach (Element elem in collector)
            {
                try
                {
                    if (elem.Category == null) continue;
                    string categoryName = elem.Category.Name;
                    if (string.IsNullOrEmpty(categoryName) || tempCategories.Contains(categoryName)) continue;

                    Parameter maHieuParam = elem.LookupParameter("Mã hiệu");
                    Parameter tenCongParam = elem.LookupParameter("Tên công việc");
                    Parameter donViParam = elem.LookupParameter("Đơn vị");

                    if (maHieuParam != null && !maHieuParam.IsReadOnly &&
                        tenCongParam != null && !tenCongParam.IsReadOnly &&
                        donViParam != null && !donViParam.IsReadOnly)
                    {
                        string catType = GetCategoryTypeByName(categoryName);
                        List<Element> elements = GetElementsByCategoryName(categoryName);

                        if (elements.Count > 0)
                        {
                            categoriesDict[categoryName] = new CategoryInfo
                            {
                                Type = catType,
                                Count = elements.Count,
                                HasParams = true
                            };
                            tempCategories.Add(categoryName);
                        }
                    }
                }
                catch
                {
                    continue;
                }
            }

            return categoriesDict;
        }

        private static string GetCategoryTypeByName(string categoryName)
        {
            Dictionary<string, string> categoryTypeMap = new Dictionary<string, string>
            {
                {"duct", "MEP"}, {"pipe", "MEP"}, {"cable", "MEP"}, {"conduit", "MEP"},
                {"elec", "MEP"}, {"mech", "MEP"}, {"plumb", "MEP"}, {"light", "MEP"},
                {"fixture", "MEP"}, {"sprinkler", "MEP"}, {"equipment", "MEP"},
                {"structural", "Structural"}, {"column", "Structural"}, {"beam", "Structural"},
                {"foundation", "Structural"}, {"truss", "Structural"}, {"frame", "Structural"},
                {"wall", "Architectural"}, {"floor", "Architectural"}, {"door", "Architectural"},
                {"window", "Architectural"}, {"stair", "Architectural"}, {"ramp", "Architectural"},
                {"roof", "Architectural"}, {"railing", "Architectural"}
            };

            string categoryLower = categoryName.ToLower();
            foreach (var kvp in categoryTypeMap)
            {
                if (categoryLower.Contains(kvp.Key))
                {
                    return kvp.Value;
                }
            }

            return "Other";
        }

        private static List<Element> GetElementsByCategoryName(string categoryName)
        {
            List<Element> elements = new List<Element>();
            try
            {
                BuiltInCategory? builtinCat = GetBuiltinCategory(categoryName);

                if (builtinCat.HasValue)
                {
                    FilteredElementCollector collector = new FilteredElementCollector(_doc)
                        .OfCategory(builtinCat.Value)
                        .WhereElementIsNotElementType();
                    elements = collector.ToList();
                }
                else
                {
                    FilteredElementCollector collector = new FilteredElementCollector(_doc)
                        .WhereElementIsNotElementType();
                    foreach (Element elem in collector)
                    {
                        try
                        {
                            if (elem.Category != null && elem.Category.Name == categoryName)
                            {
                                elements.Add(elem);
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }

                return elements.OrderBy(e => e.Id.IntegerValue).ToList();
            }
            catch
            {
                return elements;
            }
        }

        private static BuiltInCategory? GetBuiltinCategory(string categoryName)
        {
            Dictionary<string, BuiltInCategory> catMap = new Dictionary<string, BuiltInCategory>
            {
                {"Doors", BuiltInCategory.OST_Doors},
                {"Windows", BuiltInCategory.OST_Windows},
                {"Walls", BuiltInCategory.OST_Walls},
                {"Floors", BuiltInCategory.OST_Floors},
                {"Stairs", BuiltInCategory.OST_Stairs},
                {"Ramps", BuiltInCategory.OST_Ramps},
                {"Roofs", BuiltInCategory.OST_Roofs},
                {"Columns", BuiltInCategory.OST_Columns}
            };

            if (catMap.ContainsKey(categoryName))
            {
                return catMap[categoryName];
            }

            return null;
        }

        public static Document Doc => _doc;
        public static int RevitVersion => _revitVersion;
    }

    // =============================================================================
    // PARAMETER DEFINITIONS
    // =============================================================================

    public class ParameterDefinitionInfo
    {
        public string Name { get; set; }
        public object ParameterType { get; set; }
        public string Description { get; set; }
    }

    public static class ParameterDefinitions
    {
        public static List<ParameterDefinitionInfo> GetDefinitions()
        {
            return new List<ParameterDefinitionInfo>
            {
                new ParameterDefinitionInfo { Name = "Mã hiệu", ParameterType = SpecTypeId.String.Text, Description = "Ma hieu" },
                new ParameterDefinitionInfo { Name = "Tên công việc", ParameterType = SpecTypeId.String.Text, Description = "Ten cong viec" },
                new ParameterDefinitionInfo { Name = "Đơn vị", ParameterType = SpecTypeId.String.Text, Description = "Don vi tinh" }
            };
        }
    }

    // =============================================================================
    // CATEGORY MAPPING & MATCH STRATEGY
    // =============================================================================

    public static class CategoryMapping
    {
        public static Dictionary<string, string> Mapping = new Dictionary<string, string>
        {
            {"Ducts", "MEP_Duct"}, {"Duct Linings", "MEP_Duct"}, {"Duct Insulations", "MEP_Duct"},
            {"Duct Placeholders", "MEP_Duct"}, {"Pipe", "MEP_Pipe"}, {"Pipes", "MEP_Pipe"},
            {"Pipe Fittings", "MEP_PipeFitting"}, {"Pipe Insulations", "MEP_Pipe"},
            {"Pipe Placeholders", "MEP_Pipe"}, {"Pipe Accessories", "MEP_PipeFitting"},
            {"Conduit", "MEP_Conduit"}, {"Conduits", "MEP_Conduit"},
            {"Conduit Fittings", "MEP_ConduitFitting"}, {"Cable Tray", "MEP_CableTray"},
            {"Cable Trays", "MEP_CableTray"}, {"Cable Tray Fittings", "MEP_CableTrayFitting"},
            {"Generic Models", "MEP_Equipment"}, {"Mechanical Equipment", "MEP_Equipment"},
            {"Plumbing Fixtures", "MEP_Equipment"}, {"Electrical Equipment", "MEP_Equipment"},
            {"Walls", "Architectural_Wall"}, {"Doors", "Architectural_Door"},
            {"Windows", "Architectural_Window"}, {"Floors", "Architectural_Floor"},
            {"Roofs", "Architectural_Roof"}, {"Stairs", "Architectural_Stair"},
            {"Railings", "Architectural_Railing"}, {"Ramps", "Architectural_Ramp"},
            {"Columns", "Structural_Column"}, {"Structural Framing", "Structural_Beam"},
            {"Structural Columns", "Structural_Column"}, {"Structural Foundations", "Structural_Foundation"},
            {"Foundation Slabs", "Structural_Foundation"}, {"Lighting Fixtures", "MEP_LightingFixture"},
            {"Electrical Fixtures", "MEP_LightingFixture"}, {"Specialty Equipment", "MEP_Equipment"},
            {"Furniture", "Architectural_Furniture"}, {"Furniture Systems", "Architectural_Furniture"}
        };

        public static Dictionary<string, List<string>> MatchStrategy = new Dictionary<string, List<string>>
        {
            {"MEP_Pipe", new List<string> {"Family", "Type", "Diameter"}},
            {"MEP_PipeFitting", new List<string> {"Family", "Type", "Diameter"}},
            {"MEP_Duct", new List<string> {"Family", "Type", "Diameter"}},
            {"MEP_DuctRect", new List<string> {"Family", "Type", "Width", "Height"}},
            {"MEP_Conduit", new List<string> {"Family", "Type", "Diameter"}},
            {"MEP_ConduitFitting", new List<string> {"Family", "Type", "Diameter"}},
            {"MEP_CableTray", new List<string> {"Family", "Type", "Width"}},
            {"MEP_CableTrayFitting", new List<string> {"Family", "Type", "Width"}},
            {"Architectural_Wall", new List<string> {"Family", "Type", "Material"}},
            {"Architectural_Floor", new List<string> {"Family", "Type", "Material"}},
            {"Architectural_Roof", new List<string> {"Family", "Type", "Material"}},
            {"Structural_Foundation", new List<string> {"Family", "Type", "Material"}},
            {"Architectural_Door", new List<string> {"Family", "Type"}},
            {"Architectural_Window", new List<string> {"Family", "Type"}},
            {"Architectural_Stair", new List<string> {"Family", "Type"}},
            {"Architectural_Railing", new List<string> {"Family", "Type"}},
            {"Structural_Column", new List<string> {"Family", "Type"}},
            {"Structural_Beam", new List<string> {"Family", "Type"}},
            {"MEP_LightingFixture", new List<string> {"Family", "Type"}},
            {"MEP_Equipment", new List<string> {"Family", "Type"}},
            {"Architectural_Furniture", new List<string> {"Family", "Type"}}
        };
    }

    // =============================================================================
    // SHARED FUNCTIONS
    // =============================================================================

    public static class SharedFunctions
    {
        public static List<Category> GetAllCategoryObjects(Document doc)
        {
            List<Category> allCats = new List<Category>();
            Categories categories = doc.Settings.Categories;
            foreach (Category cat in categories)
            {
                try
                {
                    if (cat.AllowsBoundParameters)
                        allCats.Add(cat);
                }
                catch { }
            }
            return allCats;
        }

        public static DefinitionFile EnsureSharedParamFile(RevitApplication.Application app)
        {
            DefinitionFile sharedFile = app.OpenSharedParameterFile();
            if (sharedFile == null)
            {
                TaskDialog.Show("Thong bao",
                    "Chua chon file Shared Parameters.\n\n" +
                    "Vui long:\n" +
                    "1. Vao: Manage > Shared Parameters > Browse...\n" +
                    "2. Chon hoac tao file .txt\n" +
                    "3. Chay lai script nay.");
                return null;
            }
            return sharedFile;
        }

        public static Dictionary<string, object> CreateAndBindSharedParameters(Document doc, RevitApplication.Application app)
        {
            var result = new Dictionary<string, object>
            {
                { "success", false },
                { "created", new List<string>() },
                { "existing", new List<string>() },
                { "errors", new List<string>() },
                { "definitions", new Dictionary<string, Definition>() }
            };

            try
            {
                DefinitionFile sharedFile = EnsureSharedParamFile(app);
                if (sharedFile == null)
                {
                    ((List<string>)result["errors"]).Add("Khong the mo Shared Parameter file");
                    return result;
                }

                DefinitionGroup group = null;
                foreach (DefinitionGroup g in sharedFile.Groups)
                {
                    if (g.Name == "Shared Parameters")
                    {
                        group = g;
                        break;
                    }
                }

                if (group == null)
                    group = sharedFile.Groups.Create("Shared Parameters");

                var paramDict = new Dictionary<string, Definition>();
                foreach (var paramDef in ParameterDefinitions.GetDefinitions())
                {
                    string paramName = paramDef.Name;
                    object paramType = paramDef.ParameterType;
                    string desc = paramDef.Description;

                    Definition definition = null;
                    foreach (Definition d in group.Definitions)
                    {
                        if (d.Name == paramName)
                        {
                            definition = d;
                            ((List<string>)result["existing"]).Add(paramName);
                            break;
                        }
                    }

                    if (definition == null)
                    {
                        try
                        {
                            ExternalDefinitionCreationOptions opts;
                            opts = new ExternalDefinitionCreationOptions(paramName, (ForgeTypeId)paramType);
                            opts.Description = desc;
                            definition = group.Definitions.Create(opts);
                            paramDict[paramName] = definition;
                            ((Dictionary<string, Definition>)result["definitions"])[paramName] = definition;
                            ((List<string>)result["created"]).Add(paramName);
                        }
                        catch (Exception ex)
                        {
                            ((List<string>)result["errors"]).Add($"{paramName}: {ex.Message}");
                            continue;
                        }
                    }
                    else
                    {
                        paramDict[paramName] = definition;
                        ((Dictionary<string, Definition>)result["definitions"])[paramName] = definition;
                    }
                }

                using (Transaction t = new Transaction(doc, "Bind Shared Parameters"))
                {
                    t.Start();
                    try
                    {
                        BindingMap bindingMap = doc.ParameterBindings;
                        List<Category> allCats = GetAllCategoryObjects(doc);

                        foreach (var kvp in paramDict)
                        {
                            string paramName = kvp.Key;
                            Definition definition = kvp.Value;

                            try
                            {
                                CategorySet catSet = app.Create.NewCategorySet();
                                foreach (Category cat in allCats)
                                {
                                    try
                                    {
                                        catSet.Insert(cat);
                                    }
                                    catch { }
                                }

                                if (catSet.Size == 0)
                                {
                                    ((List<string>)result["errors"]).Add($"{paramName}: CategorySet rong");
                                    continue;
                                }

                                InstanceBinding binding = new InstanceBinding(catSet);
                                if (!bindingMap.Contains(definition))
                                {
                                    try
                                    {
                                        bindingMap.Insert(definition, binding);
                                    }
                                    catch
                                    {
                                        try
                                        {
                                            bindingMap.Insert(definition, binding);
                                        }
                                        catch { }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                ((List<string>)result["errors"]).Add($"{paramName}: {ex.Message}");
                            }
                        }

                        t.Commit();
                        result["success"] = ((List<string>)result["created"]).Count > 0 || ((List<string>)result["existing"]).Count > 0;
                    }
                    catch (Exception ex)
                    {
                        t.RollBack();
                        ((List<string>)result["errors"]).Add($"Shared Param Transaction: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                ((List<string>)result["errors"]).Add(ex.Message);
            }

            return result;
        }

        public static Dictionary<string, object> CreateProjectParameters(Document doc, RevitApplication.Application app)
        {
            var result = new Dictionary<string, object>
            {
                { "success", false },
                { "created", new List<string>() },
                { "existing", new List<string>() },
                { "errors", new List<string>() }
            };

            try
            {
                DefinitionFile sharedFile = EnsureSharedParamFile(app);
                if (sharedFile == null)
                {
                    ((List<string>)result["errors"]).Add("Khong the mo Shared Parameter file");
                    return result;
                }

                DefinitionGroup group = null;
                foreach (DefinitionGroup g in sharedFile.Groups)
                {
                    if (g.Name == "Shared Parameters")
                    {
                        group = g;
                        break;
                    }
                }

                if (group == null)
                {
                    ((List<string>)result["errors"]).Add("Khong tim thay group Shared Parameters");
                    return result;
                }

                Dictionary<string, Definition> sharedDefs = new Dictionary<string, Definition>();
                foreach (var paramDef in ParameterDefinitions.GetDefinitions())
                {
                    string paramName = paramDef.Name;
                    foreach (Definition d in group.Definitions)
                    {
                        if (d.Name == paramName)
                        {
                            sharedDefs[paramName] = d;
                            break;
                        }
                    }
                }

                if (sharedDefs.Count < 3)
                {
                    ((List<string>)result["errors"]).Add("Khong tim thay day du 3 Shared Parameters");
                    return result;
                }

                using (Transaction t = new Transaction(doc, "Create Project Parameters from Shared Parameters"))
                {
                    t.Start();
                    try
                    {
                        BindingMap bindingMap = doc.ParameterBindings;
                        List<Category> allCatsList = GetAllCategoryObjects(doc);

                        if (allCatsList.Count == 0)
                        {
                            ((List<string>)result["errors"]).Add("Khong tim thay categories nao");
                            t.RollBack();
                            return result;
                        }

                        foreach (var kvp in sharedDefs)
                        {
                            string paramName = kvp.Key;
                            Definition sharedDef = kvp.Value;

                            try
                            {
                                CategorySet catSet = app.Create.NewCategorySet();
                                foreach (Category cat in allCatsList)
                                {
                                    try
                                    {
                                        catSet.Insert(cat);
                                    }
                                    catch { }
                                }

                                if (catSet.Size == 0)
                                {
                                    ((List<string>)result["errors"]).Add($"{paramName}: CategorySet rong");
                                    continue;
                                }

                                InstanceBinding binding = new InstanceBinding(catSet);

                                bool isAlreadyBound = false;
                                try
                                {
                                    var iterator = bindingMap.ForwardIterator();
                                    iterator.Reset();
                                    while (iterator.MoveNext())
                                    {
                                        if (iterator.Key == sharedDef)
                                        {
                                            isAlreadyBound = true;
                                            break;
                                        }
                                    }
                                }
                                catch
                                {
                                    try
                                    {
                                        if (bindingMap.Contains(sharedDef))
                                            isAlreadyBound = true;
                                    }
                                    catch { }
                                }

                                if (isAlreadyBound)
                                {
                                    ((List<string>)result["existing"]).Add(paramName);
                                    continue;
                                }

                                try
                                {
                                    bindingMap.Insert(sharedDef, binding);
                                    ((List<string>)result["created"]).Add(paramName);
                                }
                                catch
                                {
                                    try
                                    {
                                        bindingMap.Insert(sharedDef, binding);
                                        ((List<string>)result["created"]).Add(paramName);
                                    }
                                    catch { }
                                }
                            }
                            catch (Exception ex)
                            {
                                ((List<string>)result["errors"]).Add($"{paramName}: {ex.Message}");
                            }
                        }

                        t.Commit();
                        result["success"] = ((List<string>)result["created"]).Count > 0;
                    }
                    catch (Exception ex)
                    {
                        t.RollBack();
                        ((List<string>)result["errors"]).Add($"Transaction error: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                ((List<string>)result["errors"]).Add(ex.Message);
            }

            return result;
        }

        public static bool AssignAllParameters(Element elem, string markValue, string tenCongViecValue, string donViValue)
        {
            try
            {
                bool success = true;

                Parameter maHieuParam = elem.LookupParameter("Mã hiệu");
                if (maHieuParam != null && !maHieuParam.IsReadOnly)
                {
                    maHieuParam.Set(markValue ?? "");
                }
                else
                {
                    success = false;
                }

                Parameter tenCongParam = elem.LookupParameter("Tên công việc");
                if (tenCongParam != null && !tenCongParam.IsReadOnly)
                {
                    tenCongParam.Set(tenCongViecValue ?? "");
                }
                else
                {
                    success = false;
                }

                Parameter donViParam = elem.LookupParameter("Đơn vị");
                if (donViParam != null && !donViParam.IsReadOnly)
                {
                    donViParam.Set(donViValue ?? "");
                }
                else
                {
                    success = false;
                }

                return success;
            }
            catch
            {
                return false;
            }
        }

        // v6.0 Template Helper Functions
        public static string GetElementFamilyName(Element element)
        {
            try
            {
                Parameter familyParam = element.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM);
                if (familyParam != null)
                {
                    string familyName = familyParam.AsValueString();
                    if (!string.IsNullOrEmpty(familyName))
                        return familyName;
                }

                if (element.Name != null)
                    return element.Name;

                if (element.Category != null)
                    return element.Category.Name;

                return "Unknown";
            }
            catch
            {
                return "Unknown";
            }
        }

        public static string GetElementTypeName(Element element, Document doc)
        {
            try
            {
                ElementType typeElem = doc.GetElement(element.GetTypeId()) as ElementType;
                if (typeElem != null && typeElem.Name != null)
                    return typeElem.Name;
                return "Unknown Type";
            }
            catch
            {
                return "Unknown Type";
            }
        }

        public static string CategorizeElement(Element element)
        {
            try
            {
                if (element.Category == null)
                    return "Other";

                string catName = element.Category.Name;

                if (CategoryMapping.Mapping.ContainsKey(catName))
                    return CategoryMapping.Mapping[catName];

                string catNameLower = catName.ToLower();

                // Pipe và fittings
                if (new[] { "pipe fitting", "pipe accessory", "elbow", "tee", "valve", "union", "flange" }
                    .Any(keyword => catNameLower.Contains(keyword)))
                    return "MEP_PipeFitting";
                else if (new[] { "pipe", "piping", "tube" }.Any(keyword => catNameLower.Contains(keyword)))
                    return "MEP_Pipe";

                // Duct và fittings
                else if (new[] { "duct fitting", "duct accessory" }.Any(keyword => catNameLower.Contains(keyword)))
                    return "MEP_Duct";
                else if (new[] { "duct", "air" }.Any(keyword => catNameLower.Contains(keyword)))
                    return "MEP_Duct";

                // Electrical
                else if (new[] { "conduit fitting", "conduit accessory" }.Any(keyword => catNameLower.Contains(keyword)))
                    return "MEP_ConduitFitting";
                else if (new[] { "conduit", "electrical" }.Any(keyword => catNameLower.Contains(keyword)))
                    return "MEP_Conduit";
                else if (new[] { "cable tray fitting", "cable tray accessory" }.Any(keyword => catNameLower.Contains(keyword)))
                    return "MEP_CableTrayFitting";
                else if (new[] { "cable tray", "cabletray" }.Any(keyword => catNameLower.Contains(keyword)))
                    return "MEP_CableTray";

                // Equipment
                else if (new[] { "equipment", "mechanical", "plumbing", "electrical", "generic", "specialty" }
                    .Any(keyword => catNameLower.Contains(keyword)))
                    return "MEP_Equipment";

                // Lighting
                else if (new[] { "lighting", "fixture", "luminair" }.Any(keyword => catNameLower.Contains(keyword)))
                    return "MEP_LightingFixture";

                // Architectural
                else if (new[] { "wall", "curtain" }.Any(keyword => catNameLower.Contains(keyword)))
                    return "Architectural_Wall";
                else if (catNameLower.Contains("door"))
                    return "Architectural_Door";
                else if (catNameLower.Contains("window"))
                    return "Architectural_Window";
                else if (catNameLower.Contains("floor"))
                    return "Architectural_Floor";
                else if (new[] { "roof", "ceiling" }.Any(keyword => catNameLower.Contains(keyword)))
                    return "Architectural_Roof";
                else if (catNameLower.Contains("stair"))
                    return "Architectural_Stair";
                else if (catNameLower.Contains("railing"))
                    return "Architectural_Railing";
                else if (catNameLower.Contains("furniture"))
                    return "Architectural_Furniture";

                // Structural
                else if (new[] { "structural", "column", "beam", "foundation", "framing", "brace" }
                    .Any(keyword => catNameLower.Contains(keyword)))
                {
                    if (catNameLower.Contains("column"))
                        return "Structural_Column";
                    else if (new[] { "beam", "framing", "brace" }.Any(keyword => catNameLower.Contains(keyword)))
                        return "Structural_Beam";
                    else if (catNameLower.Contains("foundation"))
                        return "Structural_Foundation";
                    else
                        return "Structural_Column";
                }

                return "Other";
            }
            catch
            {
                return "Other";
            }
        }

        public static string GetElementDimension(Element element, string paramName)
        {
            try
            {
                Parameter param = element.LookupParameter(paramName);
                if (param != null && param.HasValue)
                {
                    return param.AsValueString();
                }
                return null;
            }
            catch
            {
                return null;
            }
        }

        public static List<Element> GetAllProjectElements(Document doc)
        {
            List<Element> elements = new List<Element>();
            try
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType();

                foreach (Element elem in collector)
                {
                    try
                    {
                        if (elem.Category == null)
                            continue;

                        string category = CategorizeElement(elem);
                        if (category == "Other")
                            continue;

                        Parameter maHieu = elem.LookupParameter("Mã hiệu");
                        Parameter tenCong = elem.LookupParameter("Tên công việc");
                        Parameter donVi = elem.LookupParameter("Đơn vị");

                        if (maHieu != null && tenCong != null && donVi != null)
                        {
                            if (!maHieu.IsReadOnly && !tenCong.IsReadOnly && !donVi.IsReadOnly)
                            {
                                elements.Add(elem);
                            }
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

                return elements;
            }
            catch
            {
                return elements;
            }
        }
    }

    // =============================================================================
    // HELPER CLASSES
    // =============================================================================

    public class CategoryInfo
    {
        public string Type { get; set; }
        public int Count { get; set; }
        public bool HasParams { get; set; }
    }

    public class CSVSearcher
    {
        private List<Dictionary<string, string>> _originalData;
        private Dictionary<string, Dictionary<string, List<int>>> _indexedData;

        public CSVSearcher(List<Dictionary<string, string>> csvData)
        {
            _originalData = csvData;
            _indexedData = BuildIndexes(csvData);
        }

        private Dictionary<string, Dictionary<string, List<int>>> BuildIndexes(List<Dictionary<string, string>> data)
        {
            var indexed = new Dictionary<string, Dictionary<string, List<int>>>
            {
                { "by_ma_hieu", new Dictionary<string, List<int>>() },
                { "by_ten_cong_tac", new Dictionary<string, List<int>>() }
            };

            for (int idx = 0; idx < data.Count; idx++)
            {
                var row = data[idx];
                string maHieu = (row.ContainsKey("Mã Hiệu") ? row["Mã Hiệu"] : "").ToLower().Trim();
                string tenCongTac = (row.ContainsKey("Tên Công Việc") ? row["Tên Công Việc"] : "").ToLower().Trim();

                if (!string.IsNullOrEmpty(maHieu))
                {
                    if (!indexed["by_ma_hieu"].ContainsKey(maHieu))
                        indexed["by_ma_hieu"][maHieu] = new List<int>();
                    indexed["by_ma_hieu"][maHieu].Add(idx);
                }

                if (!string.IsNullOrEmpty(tenCongTac))
                {
                    if (!indexed["by_ten_cong_tac"].ContainsKey(tenCongTac))
                        indexed["by_ten_cong_tac"][tenCongTac] = new List<int>();
                    indexed["by_ten_cong_tac"][tenCongTac].Add(idx);
                }
            }

            return indexed;
        }

        public List<Dictionary<string, string>> Search(string keyword, int searchType = 0)
        {
            if (string.IsNullOrEmpty(keyword))
                return _originalData;

            keyword = keyword.ToLower().Trim();
            List<Dictionary<string, string>> results = new List<Dictionary<string, string>>();
            HashSet<int> seenIndices = new HashSet<int>();

            if (searchType == 0 || searchType == 1)
            {
                if (_indexedData["by_ma_hieu"].ContainsKey(keyword))
                {
                    foreach (int idx in _indexedData["by_ma_hieu"][keyword])
                    {
                        if (!seenIndices.Contains(idx))
                        {
                            seenIndices.Add(idx);
                            results.Add(_originalData[idx]);
                        }
                    }
                }

                foreach (var kvp in _indexedData["by_ma_hieu"])
                {
                    if (kvp.Key.Contains(keyword))
                    {
                        foreach (int idx in kvp.Value)
                        {
                            if (!seenIndices.Contains(idx))
                            {
                                seenIndices.Add(idx);
                                results.Add(_originalData[idx]);
                            }
                        }
                    }
                }
            }

            if (searchType == 0 || searchType == 2)
            {
                if (_indexedData["by_ten_cong_tac"].ContainsKey(keyword))
                {
                    foreach (int idx in _indexedData["by_ten_cong_tac"][keyword])
                    {
                        if (!seenIndices.Contains(idx))
                        {
                            seenIndices.Add(idx);
                            results.Add(_originalData[idx]);
                        }
                    }
                }

                foreach (var kvp in _indexedData["by_ten_cong_tac"])
                {
                    if (kvp.Key.Contains(keyword))
                    {
                        foreach (int idx in kvp.Value)
                        {
                            if (!seenIndices.Contains(idx))
                            {
                                seenIndices.Add(idx);
                                results.Add(_originalData[idx]);
                            }
                        }
                    }
                }
            }

            if (results.Count > 1000)
                results = results.Take(1000).ToList();

            return results;
        }
    }

    public class ElementWrapper
    {
        public Element Element { get; set; }
        public int ElementId { get; set; }
        public string Name { get; set; }
        public string FamilyName { get; set; }
        public string MaHieuValue { get; set; }
        public string TenCongViecValue { get; set; }
        public string DonViValue { get; set; }
        public double QuantityValue { get; set; }
        public string UnitName { get; set; }
        public string DiameterValue { get; set; }
        public string WidthValue { get; set; }
        public string HeightValue { get; set; }

        public ElementWrapper(Element element)
        {
            Element = element;
            ElementId = element.Id.IntegerValue;
            Name = element.Name ?? "Unnamed";
            FamilyName = "";
            MaHieuValue = "";
            TenCongViecValue = "";
            DonViValue = "";
            QuantityValue = 0;
            UnitName = "N/A";
            DiameterValue = "N/A";
            WidthValue = "N/A";
            HeightValue = "N/A";

            try
            {
                Parameter familyParam = element.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM);
                if (familyParam != null)
                {
                    FamilyName = familyParam.AsValueString() ?? "Unknown Family";
                }
                else
                {
                    FamilyName = "Unknown Family";
                }
            }
            catch
            {
                FamilyName = "Unknown Family";
            }

            GetDimensions();
            GetQuantityAndUnit(element);
            GetShareParameters(element);
        }

        private void GetDimensions()
        {
            string[] paramNames = { "Diameter", "Width", "Height" };
            Dictionary<string, string> values = new Dictionary<string, string>();

            foreach (string paramName in paramNames)
            {
                try
                {
                    Parameter param = Element.LookupParameter(paramName);
                    if (param != null && param.HasValue)
                    {
                        values[paramName.ToLower()] = param.AsValueString();
                    }
                    else
                    {
                        values[paramName.ToLower()] = "N/A";
                    }
                }
                catch
                {
                    values[paramName.ToLower()] = "N/A";
                }
            }

            DiameterValue = values["diameter"];
            WidthValue = values["width"];
            HeightValue = values["height"];
        }

        private void GetQuantityAndUnit(Element element)
        {
            var paramPriority = new[]
            {
                ("Length", 0.3048, "Length (m)"),
                ("Area", 0.092903, "Area (m²)"),
                ("Volume", 0.0283168, "Volume (m³)")
            };

            foreach (var (paramName, conversion, unit) in paramPriority)
            {
                try
                {
                    Parameter param = element.LookupParameter(paramName);
                    if (param != null && param.HasValue && param.StorageType == StorageType.Double)
                    {
                        double val = param.AsDouble();
                        if (Math.Abs(val) > 0.0001)
                        {
                            QuantityValue = val * conversion;
                            UnitName = unit;
                            return;
                        }
                    }
                }
                catch
                {
                    continue;
                }
            }

            QuantityValue = 1;
            UnitName = "Count";
        }

        private void GetShareParameters(Element element)
        {
            Dictionary<string, string> paramMap = new Dictionary<string, string>
            {
                {"Mã hiệu", "MaHieuValue"},
                {"Tên công việc", "TenCongViecValue"},
                {"Đơn vị", "DonViValue"}
            };

            foreach (var kvp in paramMap)
            {
                try
                {
                    Parameter param = element.LookupParameter(kvp.Key);
                    if (param != null && !param.IsReadOnly && param.HasValue)
                    {
                        string value = param.AsString() ?? "";
                        switch (kvp.Value)
                        {
                            case "MaHieuValue": MaHieuValue = value; break;
                            case "TenCongViecValue": TenCongViecValue = value; break;
                            case "DonViValue": DonViValue = value; break;
                        }
                    }
                }
                catch
                {
                    continue;
                }
            }
        }
    }

    public class FamilyTypeDimensionGroup
    {
        public string FamilyName { get; set; }
        public string TypeName { get; set; }
        public string DiameterValue { get; set; }
        public string WidthValue { get; set; }
        public string HeightValue { get; set; }
        public List<ElementWrapper> Elements { get; set; }
        public int ElementCount { get; set; }
        public string PropertyToUse { get; set; }
        public double QuantityValue { get; set; }
        public string UnitName { get; set; }
        public string MaHieuValue { get; set; }
        public string TenCongViecValue { get; set; }

        public FamilyTypeDimensionGroup(string familyName, string typeName, string diameterValue,
            string widthValue, string heightValue, List<ElementWrapper> elementsList,
            string propertyToUse = "Count")
        {
            FamilyName = familyName;
            TypeName = typeName;
            DiameterValue = diameterValue;
            WidthValue = widthValue;
            HeightValue = heightValue;
            Elements = elementsList ?? new List<ElementWrapper>();
            ElementCount = Elements.Count;
            PropertyToUse = propertyToUse;

            var quantityInfo = GetQuantityByProperty(elementsList);
            QuantityValue = quantityInfo.Item1;
            UnitName = quantityInfo.Item2;
            MaHieuValue = GetMaHieuFromElements();
            TenCongViecValue = GetTenCongViecFromElements();
        }

        private string GetMaHieuFromElements()
        {
            if (Elements != null && Elements.Count > 0)
            {
                try
                {
                    Element elem = Elements[0].Element;
                    Parameter maHieuParam = elem.LookupParameter("Mã hiệu");
                    if (maHieuParam != null && maHieuParam.HasValue)
                        return maHieuParam.AsString() ?? "";
                }
                catch { }
            }
            return "";
        }

        private string GetTenCongViecFromElements()
        {
            if (Elements != null && Elements.Count > 0)
            {
                try
                {
                    Element elem = Elements[0].Element;
                    Parameter tenCongParam = elem.LookupParameter("Tên công việc");
                    if (tenCongParam != null && tenCongParam.HasValue)
                        return tenCongParam.AsString() ?? "";
                }
                catch { }
            }
            return "";
        }

        private Tuple<double, string> GetQuantityByProperty(List<ElementWrapper> elementsList)
        {
            if (PropertyToUse == "Count")
                return new Tuple<double, string>(ElementCount, "Count");

            Dictionary<string, Tuple<string, double, string>> propertyMap = new Dictionary<string, Tuple<string, double, string>>
            {
                {"Length", new Tuple<string, double, string>("Length", 0.3048, "Length (m)")},
                {"Area", new Tuple<string, double, string>("Area", 0.092903, "Area (m²)")},
                {"Volume", new Tuple<string, double, string>("Volume", 0.0283168, "Volume (m³)")}
            };

            if (propertyMap.ContainsKey(PropertyToUse))
            {
                var paramInfo = propertyMap[PropertyToUse];
                double totalValue = 0.0;

                foreach (ElementWrapper wrapper in elementsList)
                {
                    try
                    {
                        Parameter param = wrapper.Element.LookupParameter(paramInfo.Item1);
                        if (param != null && param.HasValue)
                        {
                            double valueFeet = param.AsDouble();
                            double valueMetric = valueFeet * paramInfo.Item2;
                            totalValue += valueMetric;
                        }
                    }
                    catch { }
                }

                return new Tuple<double, string>(totalValue, paramInfo.Item3);
            }

            return new Tuple<double, string>(ElementCount, "Count");
        }
    }

    // =============================================================================
    // MAIN FORM CLASS - v5.0 + v6.0 HYBRID
    // =============================================================================

    public partial class MarkAssignmentForm : WinForms.Form
    {
        private List<Dictionary<string, string>> _allCsvData;
        private List<Dictionary<string, string>> _displayedCsvData;
        private CSVSearcher _csvSearcher;
        private Dictionary<string, CategoryInfo> _categoriesDict;
        private List<Element> _currentElements;
        private string _currentCategory;
        private string _currentProperty;
        private List<FamilyTypeDimensionGroup> _familyTypeDimensionGroups;
        private Dictionary<int, Dictionary<string, string>> _mapping;
        private string _selectedMarkValue;
        private string _selectedTenCongViec;
        private string _selectedDonVi;

        private Dictionary<int, string> _maHieuEdits;
        private Dictionary<int, string> _tenCongViecEdits;

        private System.Windows.Forms.Timer _searchTimer;
        private Dictionary<string, object> _currentTemplate;
        private string _templateFilePath;
        private List<Element> _lastAutoAssignedElements;

        // UI Controls
        private DataGridView _elementsGrid;
        private DataGridView _markGrid;
        private WinForms.ComboBox _categoryCombo;
        private WinForms.ComboBox _propertyCombo;
        private WinForms.ComboBox _searchTypeCombo;
        private WinForms.TextBox _searchBox;
        private Label _infoLabel;
        private Label _statusLabel;
        private Label _searchResultsLabel;
        private Label _selectedCountLabel;
        private VScrollBar _markScrollbar;

        public MarkAssignmentForm(List<Dictionary<string, string>> csvData, Dictionary<string, CategoryInfo> categoriesDict)
        {
            Text = "TOOL GÁN MÃ ĐỊNH MỨC - THÔNG TƯ 12 - 2021/BXD";
            Width = 1650;
            Height = 800;
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(1200, 600);

            // Initialize attributes
            _elementsGrid = null;
            _markGrid = null;
            _categoryCombo = null;
            _propertyCombo = null;
            _infoLabel = null;
            _statusLabel = null;

            // v5.0 attributes
            _allCsvData = csvData;
            _csvSearcher = new CSVSearcher(csvData);
            _displayedCsvData = new List<Dictionary<string, string>>(csvData);

            _categoriesDict = categoriesDict;
            _currentElements = new List<Element>();
            _currentCategory = null;
            _currentProperty = "Count";
            _familyTypeDimensionGroups = new List<FamilyTypeDimensionGroup>();
            _mapping = new Dictionary<int, Dictionary<string, string>>();
            _selectedMarkValue = null;
            _selectedTenCongViec = null;
            _selectedDonVi = null;

            _maHieuEdits = new Dictionary<int, string>();
            _tenCongViecEdits = new Dictionary<int, string>();

            _searchTimer = new System.Windows.Forms.Timer();
            _searchTimer.Interval = 300;
            _searchTimer.Tick += OnSearchTimerTick;

            // v6.0 attributes
            _currentTemplate = null;
            _templateFilePath = null;
            _lastAutoAssignedElements = new List<Element>();

            AutoSetup();
            CreateUI();
            ShowWelcomeMessage();
        }

        private void AutoSetup()
        {
            try
            {
                var result1 = SharedFunctions.CreateAndBindSharedParameters(MainCommand.Doc, MainCommand.Doc.Application);
                var result2 = SharedFunctions.CreateProjectParameters(MainCommand.Doc, MainCommand.Doc.Application);

                int successCount = ((List<string>)result1["created"]).Count + ((List<string>)result1["existing"]).Count +
                                 ((List<string>)result2["created"]).Count + ((List<string>)result2["existing"]).Count;

                if (successCount > 0)
                {
                    Debug.WriteLine("✅ Đã tự động tạo Shared Parameters thành công!");
                }
                else
                {
                    Debug.WriteLine("⚠️ Shared Parameters đã tồn tại hoặc có lỗi");
                }

                // Reload categories
                var newCategoriesDict = GetAllCategoriesFromProject();
                _categoriesDict = newCategoriesDict;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Lỗi khi tự động thiết lập: " + ex.Message);
            }
        }

        private void ShowWelcomeMessage()
        {
            string welcomeText = @"
🚀 CHÀO MỪNG ĐẾN VỚI TOOL GÁN MÃ ĐỊNH MỨC!

Tool đã sẵn sàng sử dụng. Để bắt đầu:

1. Chọn category từ dropdown
2. Shared Parameters đã được tạo tự động
3. Click 'Hướng Dẫn' nếu cần trợ giúp

Bắt đầu bằng cách:
• Chọn category từ dropdown
• Tick chọn elements trong bảng bên phải
• Chọn mã trong bảng bên trái  
• Click 'Gán/Lưu'

Sau khi gán thủ công vài elements, hãy:
1. Click 'Export Template'
2. Click 'Auto Assign'

Chúc bạn sử dụng hiệu quả!";

            if (_statusLabel != null)
            {
                _statusLabel.Text = "Tool đã sẵn sàng! Chọn category để bắt đầu.";
            }
            Debug.WriteLine(welcomeText);
        }

        private void CreateUI()
        {
            // Main Layout
            TableLayoutPanel mainLayout = new TableLayoutPanel();
            mainLayout.Dock = DockStyle.Fill;
            mainLayout.RowCount = 3;
            mainLayout.ColumnCount = 1;
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 70));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));
            this.Controls.Add(mainLayout);

            // Header Panel
            WinForms.Panel headerPanel = CreateHeaderPanel();
            mainLayout.Controls.Add(headerPanel, 0, 0);

            // Content Panel
            WinForms.Panel contentPanel = new WinForms.Panel();
            contentPanel.Dock = DockStyle.Fill;
            contentPanel.BackColor = Drawing.Color.White;
            mainLayout.Controls.Add(contentPanel, 0, 1);

            // Split Container
            SplitContainer splitContainer = new SplitContainer();
            splitContainer.Dock = DockStyle.Fill;
            splitContainer.Orientation = Orientation.Vertical;
            splitContainer.SplitterDistance = 700;
            splitContainer.SplitterWidth = 8;
            contentPanel.Controls.Add(splitContainer);

            // Left Panel
            WinForms.Panel leftPanel = CreateLeftPanel();
            splitContainer.Panel1.Controls.Add(leftPanel);

            // Right Panel
            WinForms.Panel rightPanel = CreateRightPanel();
            splitContainer.Panel2.Controls.Add(rightPanel);

            // Footer Panel
            WinForms.Panel footerPanel = CreateFooterPanel();
            mainLayout.Controls.Add(footerPanel, 0, 2);
        }

        private WinForms.Panel CreateHeaderPanel()
        {
            WinForms.Panel panel = new WinForms.Panel();
            panel.Dock = DockStyle.Fill;
            panel.BackColor = Drawing.Color.LightGray;
            panel.Padding = new Padding(8, 5, 8, 5);

            // Help Button
            Button btnHelp = new Button();
            btnHelp.Text = "Hướng Dẫn";
            btnHelp.Location = new Drawing.Point(8, 8);
            btnHelp.Size = new Size(100, 23);
            btnHelp.BackColor = Drawing.Color.LightBlue;
            btnHelp.Font = new Font(btnHelp.Font, FontStyle.Bold);
            btnHelp.Click += ShowHelpClicked;
            panel.Controls.Add(btnHelp);

            // Create Shared Parameters Button
            Button btnCreateShared = new Button();
            btnCreateShared.Text = "Create Shared Parameters";
            btnCreateShared.Location = new Drawing.Point(115, 8);
            btnCreateShared.Size = new Size(180, 23);
            btnCreateShared.BackColor = Drawing.Color.Gold;
            btnCreateShared.Font = new Font(btnCreateShared.Font, FontStyle.Bold);
            btnCreateShared.Click += CreateSharedParametersClicked;
            panel.Controls.Add(btnCreateShared);

            // Refresh Categories Button
            Button btnRefresh = new Button();
            btnRefresh.Text = "Refresh Categories";
            btnRefresh.Location = new Drawing.Point(115, 35);
            btnRefresh.Size = new Size(180, 23);
            btnRefresh.BackColor = Drawing.Color.LightYellow;
            btnRefresh.Font = new Font(btnRefresh.Font, FontStyle.Bold);
            btnRefresh.Click += RefreshCategoriesClicked;
            panel.Controls.Add(btnRefresh);

            // Category Label
            Label lblCategory = new Label();
            lblCategory.Text = "Chọn Category:";
            lblCategory.Location = new Drawing.Point(305, 8);
            lblCategory.Size = new Size(100, 18);
            lblCategory.Font = new Font(lblCategory.Font, FontStyle.Bold);
            panel.Controls.Add(lblCategory);

            // Category ComboBox
            _categoryCombo = new WinForms.ComboBox();
            _categoryCombo.Location = new Drawing.Point(410, 8);
            _categoryCombo.Size = new Size(220, 23);
            _categoryCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _categoryCombo.SelectedIndexChanged += CategoryChanged;
            panel.Controls.Add(_categoryCombo);

            // Property Label
            Label lblProperty = new Label();
            lblProperty.Text = "Property:";
            lblProperty.Location = new Drawing.Point(640, 8);
            lblProperty.Size = new Size(60, 18);
            lblProperty.Font = new Font(lblProperty.Font, FontStyle.Bold);
            panel.Controls.Add(lblProperty);

            // Property ComboBox
            _propertyCombo = new WinForms.ComboBox();
            _propertyCombo.Location = new Drawing.Point(705, 8);
            _propertyCombo.Size = new Size(120, 23);
            _propertyCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _propertyCombo.Items.AddRange(new object[] { "Count", "Length", "Area", "Volume" });
            _propertyCombo.SelectedIndex = 0;
            _propertyCombo.SelectedIndexChanged += PropertyChanged;
            panel.Controls.Add(_propertyCombo);

            // Info Label
            _infoLabel = new Label();
            _infoLabel.Text = "Tool đã sẵn sàng! Chọn category để bắt đầu.";
            _infoLabel.Location = new Drawing.Point(835, 8);
            _infoLabel.Size = new Size(600, 18);
            _infoLabel.ForeColor = Drawing.Color.DarkBlue;
            _infoLabel.Font = new Font(_infoLabel.Font, FontStyle.Bold);
            panel.Controls.Add(_infoLabel);

            PopulateCategories();
            return panel;
        }

        private WinForms.Panel CreateLeftPanel()
        {
            TableLayoutPanel leftMain = new TableLayoutPanel();
            leftMain.Dock = DockStyle.Fill;
            leftMain.RowCount = 3;
            leftMain.ColumnCount = 2;
            leftMain.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));
            leftMain.RowStyles.Add(new RowStyle(SizeType.Absolute, 60));
            leftMain.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            leftMain.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20));
            leftMain.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            // Title
            Label leftTitle = new Label();
            leftTitle.Text = "BẢNG TRA MÃ HIỆU ĐỊNH MỨC:";
            leftTitle.Font = new Font(leftTitle.Font, FontStyle.Bold);
            leftTitle.Dock = DockStyle.Fill;
            leftTitle.BackColor = Drawing.Color.LightBlue;
            leftTitle.TextAlign = ContentAlignment.MiddleLeft;
            leftMain.Controls.Add(leftTitle, 0, 0);
            leftMain.SetColumnSpan(leftTitle, 2);

            // Search Panel
            WinForms.Panel searchPanel = CreateOptimizedSearchPanel();
            leftMain.Controls.Add(searchPanel, 0, 1);
            leftMain.SetColumnSpan(searchPanel, 2);

            // Scrollbar
            _markScrollbar = new VScrollBar();
            _markScrollbar.Dock = DockStyle.Fill;
            _markScrollbar.Scroll += MarkScrollbarScroll;
            leftMain.Controls.Add(_markScrollbar, 0, 2);

            // Mark Grid
            _markGrid = new DataGridView();
            _markGrid.Dock = DockStyle.Fill;
            _markGrid.ReadOnly = true;
            _markGrid.AllowUserToAddRows = false;
            _markGrid.AllowUserToDeleteRows = false;
            _markGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _markGrid.MultiSelect = false;
            _markGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            _markGrid.CellClick += MarkGridCellClick;

            _markGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Mã Hiệu", Width = 80 });
            _markGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Tên Công Việc", Width = 200 });
            _markGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Đơn Vị", Width = 80 });

            leftMain.Controls.Add(_markGrid, 1, 2);

            WinForms.Panel panel = new WinForms.Panel();
            panel.Dock = DockStyle.Fill;
            panel.Controls.Add(leftMain);
            return panel;
        }

        private WinForms.Panel CreateOptimizedSearchPanel()
        {
            WinForms.Panel panel = new WinForms.Panel();
            panel.Dock = DockStyle.Fill;
            panel.BackColor = Drawing.Color.FromArgb(240, 240, 240);
            panel.Padding = new Padding(5, 3, 5, 3);

            Label lblSearch = new Label();
            lblSearch.Text = "Tìm:";
            lblSearch.Location = new Drawing.Point(5, 8);
            lblSearch.Size = new Size(40, 18);
            lblSearch.Font = new Font(lblSearch.Font, FontStyle.Bold);
            panel.Controls.Add(lblSearch);

            _searchBox = new WinForms.TextBox();
            _searchBox.Location = new Drawing.Point(50, 8);
            _searchBox.Size = new Size(180, 20);
            _searchBox.TextChanged += SearchTextChangedDebounced;
            panel.Controls.Add(_searchBox);

            Label lblType = new Label();
            lblType.Text = "Loại:";
            lblType.Location = new Drawing.Point(5, 32);
            lblType.Size = new Size(40, 18);
            lblType.Font = new Font(lblType.Font, FontStyle.Bold);
            panel.Controls.Add(lblType);

            _searchTypeCombo = new WinForms.ComboBox();
            _searchTypeCombo.Location = new Drawing.Point(50, 32);
            _searchTypeCombo.Size = new Size(120, 20);
            _searchTypeCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _searchTypeCombo.Items.AddRange(new object[] { "Cả hai", "Theo mã", "Theo tên" });
            _searchTypeCombo.SelectedIndex = 0;
            _searchTypeCombo.SelectedIndexChanged += SearchTypeChanged;
            panel.Controls.Add(_searchTypeCombo);

            Button btnClear = new Button();
            btnClear.Text = "Clear";
            btnClear.Location = new Drawing.Point(175, 32);
            btnClear.Size = new Size(55, 20);
            btnClear.Click += ClearSearchClicked;
            panel.Controls.Add(btnClear);

            _searchResultsLabel = new Label();
            _searchResultsLabel.Text = "Tất cả: 0";
            _searchResultsLabel.Location = new Drawing.Point(240, 8);
            _searchResultsLabel.Size = new Size(120, 18);
            _searchResultsLabel.ForeColor = Drawing.Color.DarkGreen;
            panel.Controls.Add(_searchResultsLabel);

            return panel;
        }

        private WinForms.Panel CreateRightPanel()
        {
            TableLayoutPanel rightMain = new TableLayoutPanel();
            rightMain.Dock = DockStyle.Fill;
            rightMain.RowCount = 3;
            rightMain.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));
            rightMain.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
            rightMain.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            // Title
            Label rightTitle = new Label();
            rightTitle.Text = "BẢNG CHỌN ĐỐI TƯỢNG REVIT:";
            rightTitle.Font = new Font(rightTitle.Font, FontStyle.Bold);
            rightTitle.Dock = DockStyle.Fill;
            rightTitle.BackColor = Drawing.Color.LightGreen;
            rightTitle.TextAlign = ContentAlignment.MiddleLeft;
            rightMain.Controls.Add(rightTitle, 0, 0);

            // Control Panel
            WinForms.Panel controlPanel = new WinForms.Panel();
            controlPanel.Dock = DockStyle.Fill;
            controlPanel.BackColor = Drawing.Color.FromArgb(250, 250, 250);

            Button btnSelectAll = new Button();
            btnSelectAll.Text = "Chọn Tất Cả";
            btnSelectAll.Size = new Size(90, 23);
            btnSelectAll.Location = new Drawing.Point(3, 3);
            btnSelectAll.Click += SelectAllClicked;
            controlPanel.Controls.Add(btnSelectAll);

            Button btnDeselectAll = new Button();
            btnDeselectAll.Text = "Bỏ Chọn";
            btnDeselectAll.Size = new Size(90, 23);
            btnDeselectAll.Location = new Drawing.Point(98, 3);
            btnDeselectAll.Click += DeselectAllClicked;
            controlPanel.Controls.Add(btnDeselectAll);

            _selectedCountLabel = new Label();
            _selectedCountLabel.Text = "Đã chọn: 0";
            _selectedCountLabel.Size = new Size(400, 23);
            _selectedCountLabel.Location = new Drawing.Point(200, 3);
            _selectedCountLabel.ForeColor = Drawing.Color.DarkGreen;
            _selectedCountLabel.Font = new Font(_selectedCountLabel.Font, FontStyle.Bold);
            controlPanel.Controls.Add(_selectedCountLabel);

            rightMain.Controls.Add(controlPanel, 0, 1);

            // Elements Grid
            _elementsGrid = new DataGridView();
            _elementsGrid.Dock = DockStyle.Fill;
            _elementsGrid.AllowUserToAddRows = false;
            _elementsGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _elementsGrid.MultiSelect = true;
            _elementsGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            _elementsGrid.CellValueChanged += ElementGridCellValueChanged;

            DataGridViewCheckBoxColumn colCheck = new DataGridViewCheckBoxColumn();
            colCheck.Name = "Chọn";
            colCheck.Width = 50;
            _elementsGrid.Columns.Add(colCheck);

            _elementsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Mã hiệu", Width = 100, ReadOnly = false });
            _elementsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Tên công việc", Width = 180, ReadOnly = false });
            _elementsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Family Name", Width = 120, ReadOnly = true });
            _elementsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Type", Width = 100, ReadOnly = true });
            _elementsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Diameter", Width = 90, ReadOnly = true });
            _elementsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Width", Width = 90, ReadOnly = true });
            _elementsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Height", Width = 90, ReadOnly = true });
            _elementsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Quantities", Width = 80, ReadOnly = true });
            _elementsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Unit", Width = 80, ReadOnly = true });

            rightMain.Controls.Add(_elementsGrid, 0, 2);

            WinForms.Panel panel = new WinForms.Panel();
            panel.Dock = DockStyle.Fill;
            panel.Controls.Add(rightMain);
            return panel;
        }

        private WinForms.Panel CreateFooterPanel()
        {
            WinForms.Panel panel = new WinForms.Panel();
            panel.Dock = DockStyle.Fill;
            panel.BackColor = Drawing.Color.FromArgb(240, 240, 240);
            panel.Padding = new Padding(8, 5, 8, 5);

            _statusLabel = new Label();
            _statusLabel.Text = "Trạng thái: Chọn category, tick elements, chọn mã, rồi click 'Gán/Lưu' hoặc 'Auto Assign'";
            _statusLabel.AutoSize = false;
            _statusLabel.Size = new Size(750, 25);
            _statusLabel.Location = new Drawing.Point(8, 8);
            _statusLabel.ForeColor = Drawing.Color.DarkBlue;
            _statusLabel.Font = new Font(_statusLabel.Font, FontStyle.Bold);
            panel.Controls.Add(_statusLabel);

            // v6.0: 3 new buttons
            Button btnExport = new Button();
            btnExport.Text = "Export Template";
            btnExport.Size = new Size(130, 28);
            btnExport.Location = new Drawing.Point(800, 6);
            btnExport.BackColor = Drawing.Color.Gold;
            btnExport.Font = new Font(btnExport.Font, FontStyle.Bold);
            btnExport.Click += ExportTemplateClicked;
            panel.Controls.Add(btnExport);

            Button btnLoad = new Button();
            btnLoad.Text = "Load Template";
            btnLoad.Size = new Size(130, 28);
            btnLoad.Location = new Drawing.Point(938, 6);
            btnLoad.BackColor = Drawing.Color.LightBlue;
            btnLoad.Font = new Font(btnLoad.Font, FontStyle.Bold);
            btnLoad.Click += LoadTemplateClicked;
            panel.Controls.Add(btnLoad);

            Button btnAuto = new Button();
            btnAuto.Text = "Auto Assign";
            btnAuto.Size = new Size(110, 28);
            btnAuto.Location = new Drawing.Point(1076, 6);
            btnAuto.BackColor = Drawing.Color.LightGreen;
            btnAuto.Font = new Font(btnAuto.Font, FontStyle.Bold);
            btnAuto.Click += AutoAssignClicked;
            panel.Controls.Add(btnAuto);

            // v5.0: 2 original buttons
            Button btnAssign = new Button();
            btnAssign.Text = "Gán/Lưu";
            btnAssign.Size = new Size(110, 28);
            btnAssign.Location = new Drawing.Point(1332, 6);
            btnAssign.BackColor = Drawing.Color.LightBlue;
            btnAssign.Font = new Font(btnAssign.Font, FontStyle.Bold);
            btnAssign.Click += AssignMarkClicked;
            panel.Controls.Add(btnAssign);

            Button btnConfirm = new Button();
            btnConfirm.Text = "Hoàn Thành";
            btnConfirm.Size = new Size(110, 28);
            btnConfirm.Location = new Drawing.Point(1450, 6);
            btnConfirm.BackColor = Drawing.Color.LightGreen;
            btnConfirm.Font = new Font(btnConfirm.Font, FontStyle.Bold);
            btnConfirm.Click += ConfirmClicked;
            panel.Controls.Add(btnConfirm);

            return panel;
        }
        private void RefreshCategoriesClicked(object sender, EventArgs e)
        {
            try
            {
                _infoLabel.Text = "Đang refresh categories...";
                WinForms.Application.DoEvents();

                var newCategoriesDict = GetAllCategoriesFromProject();
                _categoriesDict = newCategoriesDict;
                PopulateCategories();

                if (_categoriesDict.Count > 0)
                {
                    _infoLabel.Text = "Refresh xong! Chọn category để bắt đầu.";
                }
                else
                {
                    _infoLabel.Text = "Refresh xong nhưng chưa có categories.";
                }
            }
            catch (Exception ex)
            {
                _infoLabel.Text = "Lỗi refresh: " + ex.Message;
            }
        }

        private Dictionary<string, CategoryInfo> GetAllCategoriesFromProject()
        {
            // Copy code từ MainCommand.GetAllCategoriesFromProject() nhưng sử dụng MainCommand.Doc
            var categoriesDict = new Dictionary<string, CategoryInfo>();
            HashSet<string> tempCategories = new HashSet<string>();

            FilteredElementCollector collector = new FilteredElementCollector(MainCommand.Doc)
                .WhereElementIsNotElementType();

            foreach (Element elem in collector)
            {
                try
                {
                    if (elem.Category == null) continue;
                    string categoryName = elem.Category.Name;
                    if (string.IsNullOrEmpty(categoryName) || tempCategories.Contains(categoryName)) continue;

                    Parameter maHieuParam = elem.LookupParameter("Mã hiệu");
                    Parameter tenCongParam = elem.LookupParameter("Tên công việc");
                    Parameter donViParam = elem.LookupParameter("Đơn vị");

                    if (maHieuParam != null && !maHieuParam.IsReadOnly &&
                        tenCongParam != null && !tenCongParam.IsReadOnly &&
                        donViParam != null && !donViParam.IsReadOnly)
                    {
                        string catType = GetCategoryTypeByName(categoryName);
                        List<Element> elements = GetElementsByCategoryName(categoryName);

                        if (elements.Count > 0)
                        {
                            categoriesDict[categoryName] = new CategoryInfo
                            {
                                Type = catType,
                                Count = elements.Count,
                                HasParams = true
                            };
                            tempCategories.Add(categoryName);
                        }
                    }
                }
                catch
                {
                    continue;
                }
            }

            return categoriesDict;
        }

        private List<Element> GetElementsByCategoryName(string categoryName)
        {
            // Copy code từ MainCommand.GetElementsByCategoryName()
            List<Element> elements = new List<Element>();
            try
            {
                BuiltInCategory? builtinCat = GetBuiltinCategory(categoryName);

                if (builtinCat.HasValue)
                {
                    FilteredElementCollector collector = new FilteredElementCollector(MainCommand.Doc)
                        .OfCategory(builtinCat.Value)
                        .WhereElementIsNotElementType();
                    elements = collector.ToList();
                }
                else
                {
                    FilteredElementCollector collector = new FilteredElementCollector(MainCommand.Doc)
                        .WhereElementIsNotElementType();
                    foreach (Element elem in collector)
                    {
                        try
                        {
                            if (elem.Category != null && elem.Category.Name == categoryName)
                            {
                                elements.Add(elem);
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }

                return elements.OrderBy(e => e.Id.IntegerValue).ToList();
            }
            catch
            {
                return elements;
            }
        }

        private BuiltInCategory? GetBuiltinCategory(string categoryName)
        {
            // Copy code từ MainCommand.GetBuiltinCategory()
            Dictionary<string, BuiltInCategory> catMap = new Dictionary<string, BuiltInCategory>
        {
            {"Doors", BuiltInCategory.OST_Doors},
            {"Windows", BuiltInCategory.OST_Windows},
            {"Walls", BuiltInCategory.OST_Walls},
            {"Floors", BuiltInCategory.OST_Floors},
            {"Stairs", BuiltInCategory.OST_Stairs},
            {"Ramps", BuiltInCategory.OST_Ramps},
            {"Roofs", BuiltInCategory.OST_Roofs},
            {"Columns", BuiltInCategory.OST_Columns}
        };

            if (catMap.ContainsKey(categoryName))
            {
                return catMap[categoryName];
            }

            return null;
        }

        private string GetCategoryTypeByName(string categoryName)
        {
            // Copy code từ MainCommand.GetCategoryTypeByName()
            Dictionary<string, string> categoryTypeMap = new Dictionary<string, string>
        {
            {"duct", "MEP"}, {"pipe", "MEP"}, {"cable", "MEP"}, {"conduit", "MEP"},
            {"elec", "MEP"}, {"mech", "MEP"}, {"plumb", "MEP"}, {"light", "MEP"},
            {"fixture", "MEP"}, {"sprinkler", "MEP"}, {"equipment", "MEP"},
            {"structural", "Structural"}, {"column", "Structural"}, {"beam", "Structural"},
            {"foundation", "Structural"}, {"truss", "Structural"}, {"frame", "Structural"},
            {"wall", "Architectural"}, {"floor", "Architectural"}, {"door", "Architectural"},
            {"window", "Architectural"}, {"stair", "Architectural"}, {"ramp", "Architectural"},
            {"roof", "Architectural"}, {"railing", "Architectural"}
        };

            string categoryLower = categoryName.ToLower();
            foreach (var kvp in categoryTypeMap)
            {
                if (categoryLower.Contains(kvp.Key))
                {
                    return kvp.Value;
                }
            }

            return "Other";
        }
        // Event Handlers
        private void ShowHelpClicked(object sender, EventArgs e)
        {
            string helpText = @"
🎯 HƯỚNG DẪN SỬ DỤNG TOOL GÁN MÃ ĐỊNH MỨC

BƯỚC 1: TẠO SHARED PARAMETERS (ĐÃ TỰ ĐỘNG)
• Tool đã tự động tạo 3 parameters: 'Mã hiệu', 'Tên công việc', 'Đơn vị'

BƯỚC 2: GÁN MÃ THỦ CÔNG (QUAN TRỌNG)
1. Chọn category (ví dụ: Pipes, Pipe Fittings, Lighting Fixtures)
2. Tick chọn elements trong bảng bên phải
3. Click chọn mã trong bảng bên trái
4. Click 'Gán/Lưu'

BƯỚC 3: EXPORT TEMPLATE
• Click 'Export Template' để lưu rules gán mã

BƯỚC 4: AUTO ASSIGN
• Click 'Auto Assign' để gán tự động cho elements còn lại

📝 LƯU Ý: Phải gán thủ công ít nhất 1 element mỗi category trước khi dùng Auto Assign!";

            MessageBox.Show(helpText, "📖 Hướng Dẫn Sử Dụng");
        }

        private void CreateSharedParametersClicked(object sender, EventArgs e)
        {
            try
            {
                _infoLabel.Text = "Đang tạo Shared Parameters...";
                WinForms.Application.DoEvents();

                var result1 = SharedFunctions.CreateAndBindSharedParameters(MainCommand.Doc, MainCommand.Doc.Application);
                var result2 = SharedFunctions.CreateProjectParameters(MainCommand.Doc, MainCommand.Doc.Application);

                int successCount = ((List<string>)result1["created"]).Count + ((List<string>)result1["existing"]).Count +
                                 ((List<string>)result2["created"]).Count + ((List<string>)result2["existing"]).Count;

                if (successCount > 0)
                {
                    MessageBox.Show("Done! Shared Parameters đã được tạo thành công.", "Thành công");
                    ReloadCategoriesAuto();
                }
                else
                {
                    MessageBox.Show("Co loi! Vui long kiem tra Shared Parameter file.", "Loi");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi");
            }
        }

        private void ReloadCategoriesAuto()
        {
            try
            {
                _infoLabel.Text = "Đang reload categories...";
                WinForms.Application.DoEvents();

                var newCategoriesDict = GetAllCategoriesFromProject();
                _categoriesDict = newCategoriesDict;
                PopulateCategories();

                if (_categoriesDict.Count > 0)
                {
                    _infoLabel.Text = "Reload xong! Chọn category để bắt đầu gán parameters.";
                }
                else
                {
                    _infoLabel.Text = "Reload xong nhung chua co categories.";
                }
            }
            catch (Exception ex)
            {
                _infoLabel.Text = "Lỗi: " + ex.Message;
            }
        }

        private void PopulateCategories()
        {
            _categoryCombo.Items.Clear();
            var sortedCats = _categoriesDict.Keys.OrderBy(k => k).ToList();

            if (sortedCats.Count == 0)
            {
                _categoryCombo.Items.Add("[Chua co categories - Click 'Create Shared Parameters']");
                _categoryCombo.Enabled = false;
            }
            else
            {
                foreach (var cat in sortedCats)
                {
                    var catInfo = _categoriesDict[cat];
                    string displayText = $"{cat} ({catInfo.Count} elements) - {catInfo.Type}";
                    _categoryCombo.Items.Add(displayText);
                }
                _categoryCombo.Enabled = true;
            }
        }

        private void CategoryChanged(object sender, EventArgs e)
        {
            if (_categoryCombo.SelectedIndex < 0)
                return;

            if (_elementsGrid == null)
            {
                Debug.WriteLine("⚠️ elements_grid chưa được khởi tạo, bỏ qua category_changed");
                return;
            }

            string selectedText = _categoryCombo.SelectedItem.ToString();
            if (selectedText.Contains("[Chua co categories"))
                return;

            string categoryName = selectedText.Split('(')[0].Trim();
            _currentCategory = categoryName;
            _currentProperty = _propertyCombo.SelectedItem.ToString();

            try
            {
                var catInfo = _categoriesDict[categoryName];
                _infoLabel.Text = $"Category: {categoryName} | Type: {catInfo.Type} | Count: {catInfo.Count}";

                _currentElements = GetElementsByCategoryName(categoryName);
                CreateDimensionGroups();
                PopulateElementsGrid();
                PopulateMarkTable();
            }
            catch (Exception ex)
            {
                string errorMsg = $"Loi khi tai elements: {ex.Message}";
                Debug.WriteLine(errorMsg);
                MessageBox.Show(errorMsg, "Loi");
            }
        }

        private void PropertyChanged(object sender, EventArgs e)
        {
            if (_propertyCombo.SelectedIndex < 0)
                return;

            _currentProperty = _propertyCombo.SelectedItem.ToString();

            if (_currentElements != null && _currentElements.Count > 0 && _elementsGrid != null)
            {
                CreateDimensionGroups();
                PopulateElementsGrid();
            }
        }

        private void CreateDimensionGroups()
        {
            _familyTypeDimensionGroups.Clear();
            var familyTypeDimensionDict = new Dictionary<Tuple<string, string, string, string, string>, List<ElementWrapper>>();

            foreach (var elem in _currentElements)
            {
                var wrapper = new ElementWrapper(elem);
                string familyName = wrapper.FamilyName;
                string typeName = wrapper.Name;
                string diameterValue = wrapper.DiameterValue;
                string widthValue = wrapper.WidthValue;
                string heightValue = wrapper.HeightValue;

                var key = new Tuple<string, string, string, string, string>(familyName, typeName, diameterValue, widthValue, heightValue);

                if (!familyTypeDimensionDict.ContainsKey(key))
                    familyTypeDimensionDict[key] = new List<ElementWrapper>();
                familyTypeDimensionDict[key].Add(wrapper);
            }

            var sortedKeys = familyTypeDimensionDict.Keys.OrderBy(k => k.Item1).ThenBy(k => k.Item2)
                .ThenBy(k => k.Item3).ThenBy(k => k.Item4).ThenBy(k => k.Item5).ToList();

            foreach (var key in sortedKeys)
            {
                var group = new FamilyTypeDimensionGroup(
                    key.Item1, key.Item2, key.Item3, key.Item4, key.Item5,
                    familyTypeDimensionDict[key], _currentProperty);
                _familyTypeDimensionGroups.Add(group);
            }
        }

        private void PopulateElementsGrid()
        {
            if (_elementsGrid == null)
            {
                Debug.WriteLine("⚠️ Không thể populate elements_grid: chưa được khởi tạo");
                return;
            }

            _elementsGrid.Rows.Clear();
            _maHieuEdits.Clear();
            _tenCongViecEdits.Clear();

            foreach (var group in _familyTypeDimensionGroups)
            {
                int rowIndex = _elementsGrid.Rows.Add();
                _elementsGrid.Rows[rowIndex].Cells[0].Value = false;
                _elementsGrid.Rows[rowIndex].Cells[1].Value = group.MaHieuValue;
                _elementsGrid.Rows[rowIndex].Cells[2].Value = group.TenCongViecValue;
                _elementsGrid.Rows[rowIndex].Cells[3].Value = group.FamilyName;
                _elementsGrid.Rows[rowIndex].Cells[4].Value = group.TypeName;
                _elementsGrid.Rows[rowIndex].Cells[5].Value = group.DiameterValue;
                _elementsGrid.Rows[rowIndex].Cells[6].Value = group.WidthValue;
                _elementsGrid.Rows[rowIndex].Cells[7].Value = group.HeightValue;

                string qtyStr;
                try
                {
                    double qtyVal = group.QuantityValue;
                    if (_currentProperty == "Count")
                        qtyStr = ((int)qtyVal).ToString();
                    else
                        qtyStr = qtyVal.ToString("F2");
                }
                catch
                {
                    qtyStr = "0";
                }

                _elementsGrid.Rows[rowIndex].Cells[8].Value = qtyStr;
                _elementsGrid.Rows[rowIndex].Cells[9].Value = group.UnitName;
            }

            UpdateSelectionCount();
            UpdateStatus();
        }

        private void PopulateMarkTable()
        {
            _displayedCsvData = new List<Dictionary<string, string>>(_allCsvData);
            UpdateMarkGrid();
            _searchResultsLabel.Text = $"Tất cả: {_allCsvData.Count}";
        }

        private void ElementsGridCellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                int rowIndex = e.RowIndex;
                int colIndex = e.ColumnIndex;

                if (colIndex != 1 && colIndex != 2)
                    return;

                object cellValue = _elementsGrid.Rows[rowIndex].Cells[colIndex].Value;
                string colName = _elementsGrid.Columns[colIndex].Name;

                if (cellValue != null && !string.IsNullOrEmpty(cellValue.ToString().Trim()))
                {
                    string newValue = cellValue.ToString().Trim();

                    if (colName == "Mã hiệu")
                        _maHieuEdits[rowIndex] = newValue;
                    else if (colName == "Tên công việc")
                        _tenCongViecEdits[rowIndex] = newValue;
                }
            }
            catch { }
        }

        private void ElementGridDirtyStateChanged(object sender, EventArgs e)
        {
            if (_elementsGrid.IsCurrentCellDirty)
                _elementsGrid.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void ElementGridCellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            UpdateSelectionCount();
        }

        private void UpdateSelectionCount()
        {
            if (_elementsGrid == null)
                return;

            int selectedElementCount = 0;
            int selectedGroupCount = 0;

            for (int i = 0; i < _elementsGrid.Rows.Count; i++)
            {
                object cellValue = _elementsGrid.Rows[i].Cells[0].Value;
                if (cellValue != null && (bool)cellValue)
                {
                    selectedGroupCount++;
                    var group = _familyTypeDimensionGroups[i];
                    selectedElementCount += group.ElementCount;
                }
            }

            _selectedCountLabel.Text = $"Đã chọn: {selectedGroupCount} groups ({selectedElementCount} elements)";
        }

        private void MarkScrollbarScroll(object sender, ScrollEventArgs e)
        {
            try
            {
                if (_markGrid.Rows.Count > 0)
                {
                    int rowIndex = _markScrollbar.Value;
                    if (rowIndex >= 0 && rowIndex < _markGrid.Rows.Count)
                        _markGrid.FirstDisplayedScrollingRowIndex = rowIndex;
                }
            }
            catch { }
        }

        private void UpdateMarkScrollbar()
        {
            int totalRows = _markGrid.Rows.Count;
            int visibleRows = _markGrid.DisplayedRowCount(false);

            if (totalRows <= visibleRows)
            {
                _markScrollbar.Enabled = false;
                _markScrollbar.Maximum = 0;
                _markScrollbar.Value = 0;
            }
            else
            {
                _markScrollbar.Enabled = true;
                _markScrollbar.Maximum = totalRows - visibleRows;
                _markScrollbar.SmallChange = 1;
                _markScrollbar.LargeChange = visibleRows;
            }
        }

        private void MarkGridCellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= _markGrid.Rows.Count)
                return;

            _markGrid.ClearSelection();
            _markGrid.Rows[e.RowIndex].Selected = true;

            _selectedMarkValue = _markGrid.Rows[e.RowIndex].Cells[0].Value?.ToString();
            _selectedTenCongViec = _markGrid.Rows[e.RowIndex].Cells[1].Value?.ToString();
            _selectedDonVi = _markGrid.Rows[e.RowIndex].Cells[2].Value?.ToString();

            if (_statusLabel != null)
            {
                _statusLabel.Text = $"Đã chọn: Mã hiệu='{_selectedMarkValue}' | Tên công việc='{_selectedTenCongViec}' | Đơn vị='{_selectedDonVi}'";
            }
        }

        private void SearchTextChangedDebounced(object sender, EventArgs e)
        {
            if (_searchTimer.Enabled)
                _searchTimer.Stop();
            _searchTimer.Start();
        }

        private void OnSearchTimerTick(object sender, EventArgs e)
        {
            _searchTimer.Stop();
            PerformOptimizedSearch();
        }

        private void SearchTypeChanged(object sender, EventArgs e)
        {
            PerformOptimizedSearch();
        }

        private void PerformOptimizedSearch()
        {
            string searchKeyword = _searchBox.Text.Trim();
            int searchType = _searchTypeCombo.SelectedIndex;

            if (string.IsNullOrEmpty(searchKeyword))
            {
                _displayedCsvData = new List<Dictionary<string, string>>(_allCsvData);
            }
            else
            {
                _displayedCsvData = _csvSearcher.Search(searchKeyword, searchType);
            }

            UpdateMarkGrid();
            _searchResultsLabel.Text = $"Hiển thị: {_displayedCsvData.Count}/{_allCsvData.Count}";
        }

        private void UpdateMarkGrid()
        {
            _markGrid.SuspendLayout();
            _markGrid.Rows.Clear();

            var displayData = _displayedCsvData;
            if (displayData.Count > 2000)
                displayData = displayData.Take(2000).ToList();

            foreach (var row in displayData)
            {
                _markGrid.Rows.Add(
                    row.ContainsKey("Mã Hiệu") ? row["Mã Hiệu"] : "",
                    row.ContainsKey("Tên Công Việc") ? row["Tên Công Việc"] : "",
                    row.ContainsKey("Đơn Vị") ? row["Đơn Vị"] : ""
                );
            }

            _markGrid.ResumeLayout();
            UpdateMarkScrollbar();
        }

        private void ClearSearchClicked(object sender, EventArgs e)
        {
            _searchBox.Text = "";
            _searchBox.Focus();
            PerformOptimizedSearch();
        }

        private void SelectAllClicked(object sender, EventArgs e)
        {
            if (_elementsGrid == null)
                return;

            for (int i = 0; i < _elementsGrid.Rows.Count; i++)
            {
                _elementsGrid.Rows[i].Cells[0].Value = true;
            }
            UpdateSelectionCount();
        }

        private void DeselectAllClicked(object sender, EventArgs e)
        {
            if (_elementsGrid == null)
                return;

            for (int i = 0; i < _elementsGrid.Rows.Count; i++)
            {
                _elementsGrid.Rows[i].Cells[0].Value = false;
            }
            UpdateSelectionCount();
        }

        private void UpdateStatus()
        {
            int assignedCount = _mapping.Count;
            if (_statusLabel != null)
                _statusLabel.Text = $"Trạng thái: Đã Gán/Lưu cho {assignedCount} elements";
        }

        private void AssignMarkClicked(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_selectedMarkValue))
            {
                MessageBox.Show("Vui long click chon ma hieu ben trai!", "Thong bao");
                return;
            }

            if (_elementsGrid == null)
            {
                MessageBox.Show("Lỗi: Bảng elements chưa được khởi tạo!", "Lỗi");
                return;
            }

            var selectedGroups = new List<FamilyTypeDimensionGroup>();
            var selectedRowIndices = new List<int>();

            for (int i = 0; i < _elementsGrid.Rows.Count; i++)
            {
                object cellValue = _elementsGrid.Rows[i].Cells[0].Value;
                if (cellValue != null && (bool)cellValue)
                {
                    selectedGroups.Add(_familyTypeDimensionGroups[i]);
                    selectedRowIndices.Add(i);
                }
            }

            if (selectedGroups.Count == 0)
            {
                MessageBox.Show("Vui long tick chon it nhat mot group ben phai!", "Thong bao");
                return;
            }

            var maHieuPerGroup = new Dictionary<int, string>();
            var tenCongViecPerGroup = new Dictionary<int, string>();

            foreach (int idx in selectedRowIndices)
            {
                if (_maHieuEdits.ContainsKey(idx))
                    maHieuPerGroup[idx] = _maHieuEdits[idx];
                else
                    maHieuPerGroup[idx] = _selectedMarkValue ?? "";

                if (_tenCongViecEdits.ContainsKey(idx))
                    tenCongViecPerGroup[idx] = _tenCongViecEdits[idx];
                else
                    tenCongViecPerGroup[idx] = _selectedTenCongViec ?? "";
            }

            var allSelectedElements = new List<ElementWrapper>();
            var groupToElements = new Dictionary<int, List<ElementWrapper>>();

            for (int groupIdx = 0; groupIdx < selectedGroups.Count; groupIdx++)
            {
                int actualRowIdx = selectedRowIndices[groupIdx];
                var group = selectedGroups[groupIdx];
                groupToElements[groupIdx] = group.Elements;
                allSelectedElements.AddRange(group.Elements);
            }

            using (Transaction t = new Transaction(MainCommand.Doc, $"Gán/Lưu cho {allSelectedElements.Count} elements"))
            {
                t.Start();

                int successCount = 0;
                try
                {
                    for (int groupIdx = 0; groupIdx < selectedGroups.Count; groupIdx++)
                    {
                        int actualRowIdx = selectedRowIndices[groupIdx];
                        var group = selectedGroups[groupIdx];

                        string maHieuToAssign = maHieuPerGroup.ContainsKey(actualRowIdx) ? maHieuPerGroup[actualRowIdx] : (_selectedMarkValue ?? "");
                        string tenCongViecToAssign = tenCongViecPerGroup.ContainsKey(actualRowIdx) ? tenCongViecPerGroup[actualRowIdx] : (_selectedTenCongViec ?? "");
                        string donViToAssign = _selectedDonVi ?? "";

                        foreach (var wrapper in groupToElements[groupIdx])
                        {
                            if (SharedFunctions.AssignAllParameters(wrapper.Element,
                                maHieuToAssign,
                                tenCongViecToAssign,
                                donViToAssign))
                            {
                                wrapper.MaHieuValue = maHieuToAssign;
                                wrapper.TenCongViecValue = tenCongViecToAssign;
                                wrapper.DonViValue = donViToAssign;
                                _mapping[wrapper.ElementId] = new Dictionary<string, string>
                                    {
                                        { "ma_hieu", maHieuToAssign },
                                        { "ten_cong_viec", tenCongViecToAssign },
                                        { "don_vi", donViToAssign }
                                    };
                                successCount++;
                            }
                        }
                    }

                    t.Commit();
                    UpdateStatus();

                    for (int groupIdx = 0; groupIdx < selectedGroups.Count; groupIdx++)
                    {
                        int actualRowIdx = selectedRowIndices[groupIdx];
                        string maHieuValue = maHieuPerGroup.ContainsKey(actualRowIdx) ? maHieuPerGroup[actualRowIdx] : (_selectedMarkValue ?? "");
                        string tenCongViecValue = tenCongViecPerGroup.ContainsKey(actualRowIdx) ? tenCongViecPerGroup[actualRowIdx] : (_selectedTenCongViec ?? "");

                        _elementsGrid.Rows[actualRowIdx].Cells[1].Value = maHieuValue;
                        _elementsGrid.Rows[actualRowIdx].Cells[2].Value = tenCongViecValue;
                    }

                    MessageBox.Show($"Da Gán/Lưu thanh cong!\n\nSo elements duoc gan: {successCount}\n\nKiem tra Output Window de xem chi tiet!", "Thanh cong");
                }
                catch (Exception ex)
                {
                    t.RollBack();
                    MessageBox.Show("Loi khi Gán/Lưu: " + ex.Message, "Loi");
                }
            }
        }

        // ============ v6.0: EVENT HANDLERS (TEMPLATE - HYBRID) ============

        private void ExportTemplateClicked(object sender, EventArgs e)
        {
            try
            {
                _statusLabel.Text = "Đang export template từ project...";
                WinForms.Application.DoEvents();

                string projectName = "CurrentProject";
                var template = ExportToTemplate(projectName);

                if (template == null)
                {
                    MessageBox.Show("Lỗi: Không thể tạo template!", "Lỗi");
                    return;
                }

                if (template.ContainsKey("categories") && ((Dictionary<string, object>)template["categories"]).Count == 0)
                {
                    MessageBox.Show(
                        "Không có element nào được gán mã định mức!\n\n" +
                        "Vui lòng thực hiện:\n" +
                        "1. Gán mã thủ công từ CSV trước\n" +
                        "2. Sau đó export template",
                        "Thông báo");
                    return;
                }

                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Title = "Chọn nơi lưu Template File";
                saveDialog.Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*";
                saveDialog.DefaultExt = ".json";

                string docPath = MainCommand.Doc.PathName;
                if (!string.IsNullOrEmpty(docPath))
                {
                    string projectFolder = Path.GetDirectoryName(docPath);
                    string projectNameOnly = Path.GetFileNameWithoutExtension(docPath);
                    saveDialog.InitialDirectory = projectFolder;
                    saveDialog.FileName = $"{projectNameOnly}_Template.json";
                }
                else
                {
                    string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    saveDialog.InitialDirectory = desktop;
                    saveDialog.FileName = "Template.json";
                }

                if (saveDialog.ShowDialog() != DialogResult.OK)
                {
                    Debug.WriteLine("User cancelled save dialog");
                    _statusLabel.Text = "Đã hủy export template";
                    return;
                }

                string templatePath = saveDialog.FileName;
                Debug.WriteLine($"Template will be saved to: {templatePath}");

                var categoriesDict = template.ContainsKey("categories") ? (Dictionary<string, object>)template["categories"] : new Dictionary<string, object>();
                int catCount = categoriesDict.Count;
                int ruleCount = 0;
                foreach (var cat in categoriesDict.Values)
                {
                    var catObj = (Dictionary<string, object>)cat;
                    if (catObj.ContainsKey("rules"))
                        ruleCount += ((List<object>)catObj["rules"]).Count;
                }

                _statusLabel.Text = $"Đang lưu template: {catCount} categories, {ruleCount} rules...";
                WinForms.Application.DoEvents();

                if (SaveTemplateFile(template, templatePath))
                {
                    MessageBox.Show(
                        "Export thành công!\n\n" +
                        $"File: {Path.GetFileName(templatePath)}\n" +
                        $"Categories: {catCount}\n" +
                        $"Rules: {ruleCount}\n\n" +
                        $"Đường dẫn:\n{templatePath}",
                        "Thành công");
                    _statusLabel.Text = $"Export xong! Template: {templatePath}";
                    _templateFilePath = templatePath;
                }
                else
                {
                    MessageBox.Show("Lỗi khi lưu file template!", "Lỗi");
                    _statusLabel.Text = "Lỗi khi lưu file template";
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"Lỗi export template: {ex.Message}";
                Debug.WriteLine(errorMsg);
                Debug.WriteLine(ex.StackTrace);
                MessageBox.Show(errorMsg, "Lỗi");
                _statusLabel.Text = "Lỗi export template";
            }
        }

        private Dictionary<string, object> ExportToTemplate(string projectName)
        {
            var template = new Dictionary<string, object>
            {
                ["template_meta"] = new Dictionary<string, object>
                {
                    ["name"] = $"{projectName} Standard Template",
                    ["version"] = "1.0",
                    ["created_date"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    ["description"] = $"Template cho {projectName}"
                },
                ["match_strategies"] = CategoryMapping.MatchStrategy,
                ["categories"] = new Dictionary<string, object>()
            };

            try
            {
                var elements = SharedFunctions.GetAllProjectElements(MainCommand.Doc);

                if (elements == null || elements.Count == 0)
                    return template;

                int elementsWithMaHieu = 0;

                foreach (var elem in elements)
                {
                    try
                    {
                        Parameter maHieuParam = elem.LookupParameter("Mã hiệu");
                        if (maHieuParam == null)
                            continue;

                        string maHieuValue = maHieuParam.AsString() ?? "";
                        if (string.IsNullOrEmpty(maHieuValue.Trim()))
                            continue;

                        elementsWithMaHieu++;
                        string familyName = SharedFunctions.GetElementFamilyName(elem);
                        string typeName = SharedFunctions.GetElementTypeName(elem, MainCommand.Doc);
                        string category = SharedFunctions.CategorizeElement(elem);

                        if (category == "Other")
                            continue;

                        var categoriesDict = (Dictionary<string, object>)template["categories"];
                        if (!categoriesDict.ContainsKey(category))
                        {
                            var matchBy = CategoryMapping.MatchStrategy.ContainsKey(category) ?
                                CategoryMapping.MatchStrategy[category] : new List<string> { "Family", "Type" };

                            categoriesDict[category] = new Dictionary<string, object>
                            {
                                ["family"] = familyName,
                                ["match_by"] = matchBy,
                                ["rules"] = new List<object>()
                            };
                        }

                        Parameter tenCongParam = elem.LookupParameter("Tên công việc");
                        Parameter donViParam = elem.LookupParameter("Đơn vị");

                        string tenCong = tenCongParam != null ? tenCongParam.AsString() ?? "" : "";
                        string donVi = donViParam != null ? donViParam.AsString() ?? "" : "";

                        var rule = new Dictionary<string, object>
                        {
                            ["family"] = familyName,
                            ["type"] = typeName,
                            ["ma_hieu"] = maHieuValue,
                            ["ten_cong_viec"] = tenCong,
                            ["don_vi"] = donVi
                        };

                        var categoryObj = (Dictionary<string, object>)categoriesDict[category];
                        var matchByList = (List<string>)categoryObj["match_by"];

                        if (matchByList.Contains("Diameter"))
                        {
                            string diameter = SharedFunctions.GetElementDimension(elem, "Diameter");
                            if (!string.IsNullOrEmpty(diameter))
                                rule["diameter"] = diameter;
                        }

                        if (matchByList.Contains("Width"))
                        {
                            string width = SharedFunctions.GetElementDimension(elem, "Width");
                            if (!string.IsNullOrEmpty(width))
                                rule["width"] = width;
                        }

                        if (matchByList.Contains("Height"))
                        {
                            string height = SharedFunctions.GetElementDimension(elem, "Height");
                            if (!string.IsNullOrEmpty(height))
                                rule["height"] = height;
                        }

                        var rulesList = (List<object>)categoryObj["rules"];
                        bool isDuplicate = false;

                        foreach (var existingRuleObj in rulesList)
                        {
                            var existingRule = (Dictionary<string, object>)existingRuleObj;
                            bool allMatch = true;

                            foreach (var key in new[] { "family", "type", "diameter", "width", "height" })
                            {
                                if (rule.ContainsKey(key))
                                {
                                    string ruleValue = rule[key]?.ToString() ?? "";
                                    string existingValue = existingRule.ContainsKey(key) ? existingRule[key]?.ToString() ?? "" : "";
                                    if (ruleValue != existingValue)
                                    {
                                        allMatch = false;
                                        break;
                                    }
                                }
                            }

                            if (allMatch)
                            {
                                isDuplicate = true;
                                break;
                            }
                        }

                        if (!isDuplicate)
                            rulesList.Add(rule);
                    }
                    catch { }
                }

                return template;
            }
            catch
            {
                return null;
            }
        }

        private bool SaveTemplateFile(Dictionary<string, object> template, string filePath)
        {
            try
            {
                string directory = Path.GetDirectoryName(filePath);
                if (!Directory.Exists(directory))
                    Directory.CreateDirectory(directory);

                string jsonString = JsonConvert.SerializeObject(template, Formatting.Indented);
                File.WriteAllText(filePath, jsonString, Encoding.UTF8);

                return File.Exists(filePath);
            }
            catch
            {
                return false;
            }
        }

        private void LoadTemplateClicked(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openDialog = new OpenFileDialog();
                openDialog.Title = "Chọn Template File để Load";
                openDialog.Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*";
                openDialog.CheckFileExists = true;
                openDialog.CheckPathExists = true;

                string docPath = MainCommand.Doc.PathName;
                if (!string.IsNullOrEmpty(docPath))
                {
                    string projectFolder = Path.GetDirectoryName(docPath);
                    openDialog.InitialDirectory = projectFolder;
                }
                else
                {
                    string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    openDialog.InitialDirectory = desktop;
                }

                if (openDialog.ShowDialog() != DialogResult.OK)
                {
                    Debug.WriteLine("User cancelled open dialog");
                    _statusLabel.Text = "Đã hủy load template";
                    return;
                }

                string templatePath = openDialog.FileName;
                Debug.WriteLine($"Loading template from: {templatePath}");

                if (!File.Exists(templatePath))
                {
                    MessageBox.Show($"File không tồn tại!\n\nPath: {templatePath}", "Lỗi");
                    return;
                }

                _currentTemplate = LoadTemplateFile(templatePath);

                if (_currentTemplate == null)
                {
                    MessageBox.Show("Lỗi khi load template!", "Lỗi");
                    return;
                }

                _templateFilePath = templatePath;
                var categoriesDict = _currentTemplate.ContainsKey("categories") ?
                    (Dictionary<string, object>)_currentTemplate["categories"] : new Dictionary<string, object>();

                int catCount = categoriesDict.Count;
                int ruleCount = 0;
                foreach (var cat in categoriesDict.Values)
                {
                    var catObj = (Dictionary<string, object>)cat;
                    if (catObj.ContainsKey("rules"))
                        ruleCount += ((List<object>)catObj["rules"]).Count;
                }

                MessageBox.Show(
                    "Load thành công!\n\n" +
                    $"File: {Path.GetFileName(templatePath)}\n" +
                    $"Categories: {catCount}\n" +
                    $"Rules: {ruleCount}\n\n" +
                    $"Đường dẫn:\n{templatePath}\n\n" +
                    "Click 'Auto Assign' để gán tự động",
                    "Thành công");

                _statusLabel.Text = $"Template loaded: {catCount} categories, {ruleCount} rules - {templatePath}";
            }
            catch (Exception ex)
            {
                string errorMsg = $"Lỗi load template: {ex.Message}";
                Debug.WriteLine(errorMsg);
                Debug.WriteLine(ex.StackTrace);
                MessageBox.Show(errorMsg, "Lỗi");
                _statusLabel.Text = "Lỗi khi load template";
            }
        }

        private Dictionary<string, object> LoadTemplateFile(string filePath)
        {
            try
            {
                string jsonString = File.ReadAllText(filePath, Encoding.UTF8);
                return JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonString);
            }
            catch
            {
                return null;
            }
        }

        private void AutoAssignClicked(object sender, EventArgs e)
        {
            try
            {
                if (_currentTemplate == null)
                {
                    MessageBox.Show("Vui lòng load template trước!", "Thông báo");
                    return;
                }

                _statusLabel.Text = "Đang auto-assign...";
                WinForms.Application.DoEvents();

                var result = AutoAssignByTemplate(_currentTemplate);

                int successCount = result.ContainsKey("success") ? ((List<object>)result["success"]).Count : 0;
                int noMatchCount = result.ContainsKey("no_match") ? ((List<object>)result["no_match"]).Count : 0;
                int errorCount = result.ContainsKey("errors") ? ((List<object>)result["errors"]).Count : 0;
                int skippedCount = result.ContainsKey("skipped") ? ((List<object>)result["skipped"]).Count : 0;

                string message = $"Auto-Assign kết thúc!\n\n✅ Gán thành công: {successCount}\n⏭️ Đã có mã (bỏ qua): {skippedCount}\n❌ Không match: {noMatchCount}\n⚠️ Lỗi: {errorCount}";
                MessageBox.Show(message, "Thành công");

                _statusLabel.Text = $"Auto-Assign xong! Gán: {successCount} | Đã có: {skippedCount} | Không match: {noMatchCount}";

                if (_currentCategory != null && successCount > 0)
                {
                    CategoryChanged(null, EventArgs.Empty);
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"Lỗi auto assign: {ex.Message}";
                Debug.WriteLine(errorMsg);
                Debug.WriteLine(ex.StackTrace);
                MessageBox.Show(errorMsg, "Lỗi");
            }
        }

        private Dictionary<string, object> AutoAssignByTemplate(Dictionary<string, object> template)
        {
            var result = new Dictionary<string, object>
            {
                ["success"] = new List<object>(),
                ["no_match"] = new List<object>(),
                ["errors"] = new List<object>(),
                ["skipped"] = new List<object>()
            };

            try
            {
                var elements = SharedFunctions.GetAllProjectElements(MainCommand.Doc);

                using (Transaction t = new Transaction(MainCommand.Doc, "Auto Assign by Template"))
                {
                    t.Start();

                    try
                    {
                        foreach (var elem in elements)
                        {
                            try
                            {
                                Parameter maHieuParam = elem.LookupParameter("Mã hiệu");
                                string maHieu = maHieuParam != null ? maHieuParam.AsString() : "";
                                if (!string.IsNullOrEmpty(maHieu?.Trim()))
                                {
                                    ((List<object>)result["skipped"]).Add(elem);
                                    continue;
                                }

                                string category = SharedFunctions.CategorizeElement(elem);
                                string familyName = SharedFunctions.GetElementFamilyName(elem);
                                string typeName = SharedFunctions.GetElementTypeName(elem, MainCommand.Doc);

                                var categoriesDict = template.ContainsKey("categories") ?
                                    (Dictionary<string, object>)template["categories"] : new Dictionary<string, object>();

                                if (!categoriesDict.ContainsKey(category))
                                {
                                    ((List<object>)result["no_match"]).Add(elem);
                                    continue;
                                }

                                var categoryObj = (Dictionary<string, object>)categoriesDict[category];
                                var matchStrategy = categoryObj.ContainsKey("match_by") ?
                                    (List<string>)categoryObj["match_by"] : new List<string> { "Family", "Type" };

                                var rules = categoryObj.ContainsKey("rules") ?
                                    (List<object>)categoryObj["rules"] : new List<object>();

                                Dictionary<string, object> matchedRule = null;
                                foreach (var ruleObj in rules)
                                {
                                    var rule = (Dictionary<string, object>)ruleObj;
                                    if (MatchElementWithRule(elem, rule, matchStrategy))
                                    {
                                        matchedRule = rule;
                                        break;
                                    }
                                }

                                if (matchedRule == null)
                                {
                                    ((List<object>)result["no_match"]).Add(elem);
                                    continue;
                                }

                                string ruleMaHieu = matchedRule.ContainsKey("ma_hieu") ? matchedRule["ma_hieu"]?.ToString() ?? "" : "";
                                string ruleTenCongViec = matchedRule.ContainsKey("ten_cong_viec") ? matchedRule["ten_cong_viec"]?.ToString() ?? "" : "";
                                string ruleDonVi = matchedRule.ContainsKey("don_vi") ? matchedRule["don_vi"]?.ToString() ?? "" : "";

                                if (SharedFunctions.AssignAllParameters(elem, ruleMaHieu, ruleTenCongViec, ruleDonVi))
                                {
                                    ((List<object>)result["success"]).Add(elem);
                                }
                            }
                            catch (Exception ex)
                            {
                                ((List<object>)result["errors"]).Add(new Tuple<Element, string>(elem, ex.Message));
                            }
                        }

                        t.Commit();
                    }
                    catch (Exception ex)
                    {
                        t.RollBack();
                        ((List<object>)result["errors"]).Add(new Tuple<string, string>("Transaction", ex.Message));
                    }
                }
            }
            catch (Exception ex)
            {
                ((List<object>)result["errors"]).Add(new Tuple<string, string>("Main", ex.Message));
            }

            return result;
        }

        private bool MatchElementWithRule(Element element, Dictionary<string, object> rule, List<string> matchStrategy)
        {
            try
            {
                string familyName = SharedFunctions.GetElementFamilyName(element);
                string typeName = SharedFunctions.GetElementTypeName(element, MainCommand.Doc);

                string ruleFamily = rule.ContainsKey("family") ? rule["family"]?.ToString() ?? "" : "";
                string ruleType = rule.ContainsKey("type") ? rule["type"]?.ToString() ?? "" : "";

                if (ruleFamily != familyName)
                    return false;

                if (ruleType != typeName)
                    return false;

                if (matchStrategy.Contains("Diameter"))
                {
                    string elemDia = SharedFunctions.GetElementDimension(element, "Diameter");
                    string ruleDia = rule.ContainsKey("diameter") ? rule["diameter"]?.ToString() ?? "" : "";
                    if (!string.IsNullOrEmpty(elemDia) && !string.IsNullOrEmpty(ruleDia) && elemDia != ruleDia)
                        return false;
                }

                if (matchStrategy.Contains("Width"))
                {
                    string elemWidth = SharedFunctions.GetElementDimension(element, "Width");
                    string ruleWidth = rule.ContainsKey("width") ? rule["width"]?.ToString() ?? "" : "";
                    if (!string.IsNullOrEmpty(elemWidth) && !string.IsNullOrEmpty(ruleWidth) && elemWidth != ruleWidth)
                        return false;
                }

                if (matchStrategy.Contains("Height"))
                {
                    string elemHeight = SharedFunctions.GetElementDimension(element, "Height");
                    string ruleHeight = rule.ContainsKey("height") ? rule["height"]?.ToString() ?? "" : "";
                    if (!string.IsNullOrEmpty(elemHeight) && !string.IsNullOrEmpty(ruleHeight) && elemHeight != ruleHeight)
                        return false;
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private List<string> ExtractMaterialLayers(Element element)
        {
            var materials = new List<string>();
            try
            {
                if (element.Category != null &&
                    (element.Category.Name == "Walls" || element.Category.Name == "Floors" || element.Category.Name == "Roofs"))
                {
                    ElementType typeElem = MainCommand.Doc.GetElement(element.GetTypeId()) as ElementType;
                    if (typeElem != null)
                    {
                        // Sử dụng reflection để gọi GetCompoundStructure
                        var method = typeElem.GetType().GetMethod("GetCompoundStructure");
                        if (method != null)
                        {
                            var compStruct = method.Invoke(typeElem, null) as CompoundStructure;
                            if (compStruct != null)
                            {
                                var layers = compStruct.GetLayers();
                                foreach (var layer in layers)
                                {
                                    try
                                    {
                                        ElementId materialId = layer.MaterialId;
                                        if (materialId != ElementId.InvalidElementId)
                                        {
                                            var matElem = MainCommand.Doc.GetElement(materialId);
                                            if (matElem != null)
                                                materials.Add(matElem.Name);
                                        }
                                    }
                                    catch { }
                                }
                            }
                        }
                    }
                }
            }
            catch { }
            return materials;
        }

        private void ConfirmClicked(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}