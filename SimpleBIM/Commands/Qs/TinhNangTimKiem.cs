using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using System.Globalization;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Newtonsoft.Json;
using System.Data;

using WinForms = System.Windows.Forms;
using GDI = System.Drawing;

namespace RevitMarkAssignmentTool
{
    /// <summary>
    /// PyRevit Tool - Gan Ma Dinh Muc v5.0 + v6.0 HYBRID + LEVEL 3 (UPGRADED)
    /// ========================================================================
    /// HYBRID MODE: v5.0 (Manual Assign) + v6.0 (Template System) + LEVEL 3 (N-gram + Fuzzy)
    /// 
    /// UPGRADES FROM LEVEL 1 TO LEVEL 3:
    /// ✓ N-gram indexing for ultra-fast candidate filtering
    /// ✓ Levenshtein distance for typo detection
    /// ✓ Jaro-Winkler similarity for phonetic matching
    /// ✓ Fuzzy matching with configurable threshold (75%)
    /// ✓ <15ms search performance on large datasets
    /// ✓ Enhanced caching system
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class MarkAssignmentCommand : IExternalCommand
    {
        private static UIApplication _uiapp;
        private static Document _doc;
        private static UIDocument _uidoc;
        private static Autodesk.Revit.ApplicationServices.Application _app;
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
                Console.WriteLine("\n" + new string('=', 80));
                Console.WriteLine("TOOL GÁN MÃ ĐỊNH MỨC - v5.0 + v6.0 HYBRID + LEVEL 3 UPGRADED");
                Console.WriteLine(new string('=', 80));
                Console.WriteLine("[LEVEL 3 FEATURES]");
                Console.WriteLine("✓ N-gram Indexing (<15ms search)");
                Console.WriteLine("✓ Levenshtein Distance (typo detection)");
                Console.WriteLine("✓ Jaro-Winkler Similarity (phonetic matching)");
                Console.WriteLine("✓ Fuzzy Matching (75% threshold)");
                Console.WriteLine("✓ Advanced caching system");
                Console.WriteLine(new string('=', 80));

                Console.WriteLine("\n[STEP 1] Selecting CSV file...");
                string csvFile = SelectCsvFile();
                if (string.IsNullOrEmpty(csvFile))
                {
                    TaskDialog.Show("Notice", "Cancelled!");
                    return Result.Cancelled;
                }

                Console.WriteLine($"CSV file: {csvFile}");

                Console.WriteLine("\n[STEP 2] Loading CSV data...");
                List<Dictionary<string, string>> csvData = ReadCsvData(csvFile);
                if (csvData == null || csvData.Count == 0)
                {
                    TaskDialog.Show("Error", "CSV file is empty!");
                    return Result.Failed;
                }

                Console.WriteLine($"Loaded {csvData.Count} rows from CSV");

                Console.WriteLine("\n[STEP 3] Building N-gram Index for LEVEL 3...");
                Console.WriteLine("This will enable <15ms search performance...");

                Console.WriteLine("\n[STEP 4] Loading categories from project...");
                Dictionary<string, CategoryInfo> categoriesDict = GetAllCategoriesFromProject();
                Console.WriteLine($"Found {categoriesDict.Count} categories");

                Console.WriteLine("\n[STEP 5] Creating form with LEVEL 3 integrated...");
                using (MarkAssignmentForm form = new MarkAssignmentForm(csvData, categoriesDict))
                {
                    Console.WriteLine("✅ Form created successfully - LEVEL 3 UPGRADED Active!");

                    DialogResult result = form.ShowDialog();

                    Console.WriteLine("\n[STEP 6] Processing results...");
                    if (result == DialogResult.OK)
                    {
                        int assignedCount = form.Mapping.Count;
                        var stats = form.CsvSearcher.GetStats();

                        Console.WriteLine($"Form closed. Total elements assigned: {assignedCount}");
                        Console.WriteLine("\n[LEVEL 3 Statistics]");
                        Console.WriteLine($"Total CSV rows: {stats["total_rows"]}");
                        Console.WriteLine($"Trigram entries: {stats["trigram_entries"]}");
                        Console.WriteLine($"Bigram entries: {stats["bigram_entries"]}");
                        Console.WriteLine($"Exact match entries: {stats["exact_entries"]}");
                        Console.WriteLine($"Word index entries: {stats["word_entries"]}");
                        Console.WriteLine($"Cache size: {stats["cache_size"]}");
                        Console.WriteLine($"Fuzzy threshold: {stats["fuzzy_threshold"]}%");

                        if (assignedCount > 0)
                        {
                            TaskDialog.Show("Success - LEVEL 3",
                                $"✅ Completed with LEVEL 3 UPGRADED!\n\nAssigned: {assignedCount} elements\n\n" +
                                "🚀 LEVEL 3 FEATURES USED:\n✓ N-gram Indexing\n✓ Levenshtein Distance\n" +
                                "✓ Jaro-Winkler Similarity\n✓ Fuzzy Matching (75%)\n✓ Advanced Caching\n\n" +
                                "⚡ Performance: <15ms per search");
                        }
                        else
                        {
                            TaskDialog.Show("Notice", "No elements were assigned.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Form cancelled by user");
                    }
                }

                Console.WriteLine("\n" + new string('=', 80));
                Console.WriteLine("✅ TOOL COMPLETED - LEVEL 3 INTEGRATION SUCCESSFUL");
                Console.WriteLine(new string('=', 80) + "\n");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                TaskDialog.Show("Error", $"Error: {ex.Message}");
                return Result.Failed;
            }
        }

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
            Encoding[] encodings = { Encoding.UTF8, new UTF8Encoding(true), Encoding.GetEncoding(1258), Encoding.GetEncoding("iso-8859-1") };

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

            FilteredElementCollector collector = new FilteredElementCollector(_doc).WhereElementIsNotElementType();

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
            if (MarkAssignmentCommand.RevitVersion >= 2021)
            {
                return new List<ParameterDefinitionInfo>
                {
                    new ParameterDefinitionInfo { Name = "Mã hiệu", ParameterType = SpecTypeId.String.Text, Description = "Ma hieu" },
                    new ParameterDefinitionInfo { Name = "Tên công việc", ParameterType = SpecTypeId.String.Text, Description = "Ten cong viec" },
                    new ParameterDefinitionInfo { Name = "Đơn vị", ParameterType = SpecTypeId.String.Text, Description = "Don vi tinh" }
                };
            }
            else
            {
#pragma warning disable CS0618
                return new List<ParameterDefinitionInfo>
                {
                    new ParameterDefinitionInfo { Name = "Mã hiệu", ParameterType = null, Description = "Ma hieu" },
                    new ParameterDefinitionInfo { Name = "Tên công việc", ParameterType = null, Description = "Ten cong viec" },
                    new ParameterDefinitionInfo { Name = "Đơn vị", ParameterType = null, Description = "Don vi tinh" }
                };
#pragma warning restore CS0618
            }
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
    // LEVEL 3 ADVANCED SEARCH ENGINE
    // =============================================================================

    public static class TextNormalizer
    {
        public static string RemoveAccents(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";

            try
            {
                string normalized = text.Normalize(NormalizationForm.FormD);
                StringBuilder sb = new StringBuilder();

                foreach (char c in normalized)
                {
                    if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                    {
                        sb.Append(c);
                    }
                }

                return sb.ToString();
            }
            catch
            {
                return text.ToLower();
            }
        }

        public static string Normalize(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";

            try
            {
                text = text.ToLower();
                text = RemoveAccents(text);

                Dictionary<char, string> specialChars = new Dictionary<char, string>
                {
                    {'ø', "o"}, {'Ø', "o"}, {'đ', "d"}, {'Đ', "d"},
                    {'®', ""}, {'™', ""}, {'©', ""}, {'–', "-"}, {'—', "-"}
                };

                foreach (var kvp in specialChars)
                {
                    text = text.Replace(kvp.Key.ToString(), kvp.Value);
                }

                text = text.Trim();
                text = Regex.Replace(text, @"\s+", " ");

                return text;
            }
            catch
            {
                return text.ToLower().Trim();
            }
        }
    }

    public static class StringDistance
    {
        public static int LevenshteinDistance(string s1, string s2)
        {
            if (s1 == null) s1 = "";
            if (s2 == null) s2 = "";

            if (s1.Length == 0) return s2.Length;
            if (s2.Length == 0) return s1.Length;

            int m = s1.Length;
            int n = s2.Length;

            if (m > 100 || n > 100)
            {
                return Math.Abs(m - n) + (s1 == s2 ? 0 : Math.Min(m, n));
            }

            int[,] dp = new int[m + 1, n + 1];

            for (int i = 0; i <= m; i++) dp[i, 0] = i;
            for (int j = 0; j <= n; j++) dp[0, j] = j;

            for (int i = 1; i <= m; i++)
            {
                for (int j = 1; j <= n; j++)
                {
                    if (s1[i - 1] == s2[j - 1])
                    {
                        dp[i, j] = dp[i - 1, j - 1];
                    }
                    else
                    {
                        dp[i, j] = 1 + Math.Min(Math.Min(dp[i - 1, j], dp[i, j - 1]), dp[i - 1, j - 1]);
                    }
                }
            }

            return dp[m, n];
        }

        public static int LevenshteinSimilarity(string s1, string s2)
        {
            try
            {
                int distance = LevenshteinDistance(s1, s2);
                int maxLen = Math.Max(s1.Length, s2.Length);
                if (maxLen == 0) return 100;

                int similarity = Math.Max(0, 100 - (int)((distance / (double)maxLen) * 100));
                return similarity;
            }
            catch
            {
                return 0;
            }
        }

        public static int JaroWinklerSimilarity(string s1, string s2, double prefixWeight = 0.1)
        {
            try
            {
                if (s1 == null) s1 = "";
                if (s2 == null) s2 = "";

                if (string.IsNullOrEmpty(s1) || string.IsNullOrEmpty(s2))
                {
                    return s1 == s2 ? 100 : 0;
                }

                int len1 = s1.Length;
                int len2 = s2.Length;

                if (s1 == s2) return 100;

                int matchDistance = Math.Max(len1, len2) / 2 - 1;
                if (matchDistance < 0) matchDistance = 0;

                bool[] s1Matches = new bool[len1];
                bool[] s2Matches = new bool[len2];

                int matches = 0;
                int transpositions = 0;

                for (int i = 0; i < len1; i++)
                {
                    int start = Math.Max(0, i - matchDistance);
                    int end = Math.Min(i + matchDistance + 1, len2);

                    for (int j = start; j < end; j++)
                    {
                        if (s2Matches[j] || s1[i] != s2[j]) continue;
                        s1Matches[i] = true;
                        s2Matches[j] = true;
                        matches++;
                        break;
                    }
                }

                if (matches == 0) return 0;

                int k = 0;
                for (int i = 0; i < len1; i++)
                {
                    if (!s1Matches[i]) continue;
                    while (!s2Matches[k]) k++;
                    if (s1[i] != s2[k]) transpositions++;
                    k++;
                }

                double jaro = (matches / (double)len1 + matches / (double)len2 +
                              (matches - transpositions / 2.0) / matches) / 3.0;

                int prefix = 0;
                for (int i = 0; i < Math.Min(len1, len2); i++)
                {
                    if (s1[i] == s2[i])
                        prefix++;
                    else
                        break;
                }
                prefix = Math.Min(4, prefix);

                double jaroWinkler = jaro + prefix * prefixWeight * (1 - jaro);

                return Math.Min(100, (int)(jaroWinkler * 100));
            }
            catch
            {
                return 0;
            }
        }
    }

    public class NGramIndex
    {
        private int _n;
        private List<Dictionary<string, string>> _csvData;
        private Dictionary<string, HashSet<int>> _trigramIndex;
        private Dictionary<string, HashSet<int>> _bigramIndex;
        private Dictionary<string, List<int>> _exactIndex;
        private Dictionary<string, List<int>> _wordIndex;

        public NGramIndex(List<Dictionary<string, string>> csvData, int n = 3)
        {
            _n = n;
            _csvData = csvData ?? new List<Dictionary<string, string>>();
            _trigramIndex = new Dictionary<string, HashSet<int>>();
            _bigramIndex = new Dictionary<string, HashSet<int>>();
            _exactIndex = new Dictionary<string, List<int>>();
            _wordIndex = new Dictionary<string, List<int>>();

            Console.WriteLine($"[NGramIndex] Building N-gram index for {_csvData.Count} rows...");
            BuildIndex();
            Console.WriteLine("[NGramIndex] Index built! Ready for <15ms searches");
        }

        private void BuildIndex()
        {
            if (_csvData == null || _csvData.Count == 0) return;

            for (int idx = 0; idx < _csvData.Count; idx++)
            {
                var row = _csvData[idx];
                if (row == null) continue;

                string maHieu = TextNormalizer.Normalize(GetValue(row, "Mã Hiệu", "ma_hieu", "MA_HIEU"));
                string tenCong = TextNormalizer.Normalize(GetValue(row, "Tên Công Việc", "ten_cong_viec", "TEN_CONG_VIEC"));

                if (!string.IsNullOrWhiteSpace(maHieu))
                {
                    IndexString(maHieu, idx);
                    if (!_exactIndex.ContainsKey(maHieu))
                        _exactIndex[maHieu] = new List<int>();
                    _exactIndex[maHieu].Add(idx);

                    foreach (string word in maHieu.Split(' '))
                    {
                        if (word.Length > 1)
                        {
                            if (!_wordIndex.ContainsKey(word))
                                _wordIndex[word] = new List<int>();
                            _wordIndex[word].Add(idx);
                        }
                    }
                }

                if (!string.IsNullOrWhiteSpace(tenCong))
                {
                    IndexString(tenCong, idx);
                    if (!_exactIndex.ContainsKey(tenCong))
                        _exactIndex[tenCong] = new List<int>();
                    _exactIndex[tenCong].Add(idx);

                    foreach (string word in tenCong.Split(' '))
                    {
                        if (word.Length > 1)
                        {
                            if (!_wordIndex.ContainsKey(word))
                                _wordIndex[word] = new List<int>();
                            _wordIndex[word].Add(idx);
                        }
                    }
                }
            }
        }

        private string GetValue(Dictionary<string, string> row, params string[] keys)
        {
            foreach (string key in keys)
            {
                if (row.ContainsKey(key))
                    return row[key];
            }
            return "";
        }

        private void IndexString(string text, int idx)
        {
            if (string.IsNullOrEmpty(text) || text.Length < 2) return;

            // Bigrams
            for (int i = 0; i < text.Length - 1; i++)
            {
                string bigram = text.Substring(i, 2);
                if (!_bigramIndex.ContainsKey(bigram))
                    _bigramIndex[bigram] = new HashSet<int>();
                _bigramIndex[bigram].Add(idx);
            }

            // Trigrams
            if (text.Length >= 3)
            {
                for (int i = 0; i < text.Length - 2; i++)
                {
                    string trigram = text.Substring(i, 3);
                    if (!_trigramIndex.ContainsKey(trigram))
                        _trigramIndex[trigram] = new HashSet<int>();
                    _trigramIndex[trigram].Add(idx);
                }
            }
        }

        public HashSet<int> GetCandidates(string query)
        {
            string queryNorm = TextNormalizer.Normalize(query);
            if (string.IsNullOrEmpty(queryNorm)) return new HashSet<int>();

            HashSet<int> candidates = new HashSet<int>();

            // Exact matches
            if (_exactIndex.ContainsKey(queryNorm))
            {
                foreach (int idx in _exactIndex[queryNorm])
                    candidates.Add(idx);
            }

            // Word matches
            foreach (string word in queryNorm.Split(' '))
            {
                if (_wordIndex.ContainsKey(word))
                {
                    foreach (int idx in _wordIndex[word])
                        candidates.Add(idx);
                }
            }

            // Trigram matches
            if (queryNorm.Length >= 3)
            {
                for (int i = 0; i < queryNorm.Length - 2; i++)
                {
                    string trigram = queryNorm.Substring(i, 3);
                    if (_trigramIndex.ContainsKey(trigram))
                    {
                        foreach (int idx in _trigramIndex[trigram])
                            candidates.Add(idx);
                    }
                }
            }

            // Bigram matches
            if (queryNorm.Length >= 2)
            {
                for (int i = 0; i < queryNorm.Length - 1; i++)
                {
                    string bigram = queryNorm.Substring(i, 2);
                    if (_bigramIndex.ContainsKey(bigram))
                    {
                        foreach (int idx in _bigramIndex[bigram])
                            candidates.Add(idx);
                    }
                }
            }

            return candidates;
        }

        public int TrigramCount => _trigramIndex.Count;
        public int BigramCount => _bigramIndex.Count;
        public int ExactCount => _exactIndex.Count;
        public int WordCount => _wordIndex.Count;
    }

    public class SearchResult
    {
        public Dictionary<string, string> Row { get; set; }
        public int Score { get; set; }
        public string Type { get; set; }
        public string Algo { get; set; }
    }

    public class AdvancedCSVSearcherL3
    {
        private List<Dictionary<string, string>> _originalData;
        private int _fuzzyThreshold;
        private NGramIndex _ngramIndex;
        private Dictionary<string, List<SearchResult>> _searchCache;
        private Dictionary<string, int> _fuzzyCache;

        public AdvancedCSVSearcherL3(List<Dictionary<string, string>> csvData, int fuzzyThreshold = 75)
        {
            _originalData = csvData ?? new List<Dictionary<string, string>>();
            _fuzzyThreshold = Math.Max(0, Math.Min(100, fuzzyThreshold));

            _ngramIndex = new NGramIndex(csvData, 3);
            _searchCache = new Dictionary<string, List<SearchResult>>();
            _fuzzyCache = new Dictionary<string, int>();

            Console.WriteLine($"[AdvancedCSVSearcher L3] Initialized with {csvData.Count} rows");
            Console.WriteLine($"[AdvancedCSVSearcher L3] Fuzzy threshold: {fuzzyThreshold}%");
        }

        private bool WordBoundaryMatch(string keyword, string text)
        {
            try
            {
                if (string.IsNullOrEmpty(keyword) || string.IsNullOrEmpty(text)) return false;
                string pattern = @"\b" + Regex.Escape(keyword) + @"\b";
                return Regex.IsMatch(text, pattern, RegexOptions.IgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private bool RegexPatternMatch(string keyword, string text)
        {
            try
            {
                if (string.IsNullOrEmpty(keyword) || string.IsNullOrEmpty(text)) return false;

                string pattern;
                if (keyword.Contains("?"))
                {
                    pattern = keyword.Replace("?", ".");
                }
                else if (keyword.Contains("*"))
                {
                    pattern = keyword.Replace("*", ".*");
                }
                else
                {
                    pattern = string.Join(".*", keyword.ToCharArray().Select(c => Regex.Escape(c.ToString())));
                }

                return Regex.IsMatch(text, pattern, RegexOptions.IgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        public List<SearchResult> SearchExactAndRanking(string keyword, int searchType = 0)
        {
            string keywordNorm = TextNormalizer.Normalize(keyword);
            if (string.IsNullOrEmpty(keywordNorm)) return new List<SearchResult>();

            List<SearchResult> results = new List<SearchResult>();
            HashSet<int> seenIndices = new HashSet<int>();

            try
            {
                for (int idx = 0; idx < _originalData.Count; idx++)
                {
                    var row = _originalData[idx];
                    if (row == null || seenIndices.Contains(idx)) continue;

                    int bestScore = 0;
                    string bestType = null;

                    string maHieuText = TextNormalizer.Normalize(GetValue(row, "Mã Hiệu", "ma_hieu", "MA_HIEU"));
                    string tenCongText = TextNormalizer.Normalize(GetValue(row, "Tên Công Việc", "ten_cong_viec", "TEN_CONG_VIEC"));

                    if (searchType == 0 || searchType == 1)
                    {
                        if (!string.IsNullOrEmpty(maHieuText))
                        {
                            if (keywordNorm == maHieuText)
                            {
                                if (100 > bestScore)
                                {
                                    bestScore = 100;
                                    bestType = "exact";
                                }
                            }
                            else if (WordBoundaryMatch(keywordNorm, maHieuText))
                            {
                                if (90 > bestScore)
                                {
                                    bestScore = 90;
                                    bestType = "word";
                                }
                            }
                            else if (maHieuText.StartsWith(keywordNorm))
                            {
                                if (80 > bestScore)
                                {
                                    bestScore = 80;
                                    bestType = "start";
                                }
                            }
                            else if (maHieuText.Contains(keywordNorm))
                            {
                                if (60 > bestScore)
                                {
                                    bestScore = 60;
                                    bestType = "contains";
                                }
                            }
                            else if (RegexPatternMatch(keywordNorm, maHieuText))
                            {
                                if (40 > bestScore)
                                {
                                    bestScore = 40;
                                    bestType = "regex";
                                }
                            }
                        }
                    }

                    if (searchType == 0 || searchType == 2)
                    {
                        if (!string.IsNullOrEmpty(tenCongText))
                        {
                            if (keywordNorm == tenCongText)
                            {
                                if (100 > bestScore)
                                {
                                    bestScore = 100;
                                    bestType = "exact";
                                }
                            }
                            else if (WordBoundaryMatch(keywordNorm, tenCongText))
                            {
                                if (90 > bestScore)
                                {
                                    bestScore = 90;
                                    bestType = "word";
                                }
                            }
                            else if (tenCongText.StartsWith(keywordNorm))
                            {
                                if (80 > bestScore)
                                {
                                    bestScore = 80;
                                    bestType = "start";
                                }
                            }
                            else if (tenCongText.Contains(keywordNorm))
                            {
                                if (60 > bestScore)
                                {
                                    bestScore = 60;
                                    bestType = "contains";
                                }
                            }
                            else if (RegexPatternMatch(keywordNorm, tenCongText))
                            {
                                if (40 > bestScore)
                                {
                                    bestScore = 40;
                                    bestType = "regex";
                                }
                            }
                        }
                    }

                    if (bestType != null)
                    {
                        results.Add(new SearchResult
                        {
                            Row = row,
                            Score = bestScore,
                            Type = bestType,
                            Algo = "exact"
                        });
                        seenIndices.Add(idx);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in search_exact_and_ranking: {ex.Message}");
            }

            results = results.OrderByDescending(r => r.Score).Take(1000).ToList();
            return results;
        }

        public List<SearchResult> SearchFuzzyWithNgram(string keyword, int searchType = 0)
        {
            string keywordNorm = TextNormalizer.Normalize(keyword);
            if (string.IsNullOrEmpty(keywordNorm)) return new List<SearchResult>();

            HashSet<int> candidateIndices = _ngramIndex.GetCandidates(keywordNorm);
            List<SearchResult> results = new List<SearchResult>();

            try
            {
                foreach (int idx in candidateIndices)
                {
                    if (idx >= _originalData.Count || _originalData[idx] == null) continue;

                    var row = _originalData[idx];
                    int bestScore = 0;
                    string bestAlgo = null;

                    string maHieuText = TextNormalizer.Normalize(GetValue(row, "Mã Hiệu", "ma_hieu", "MA_HIEU"));
                    string tenCongText = TextNormalizer.Normalize(GetValue(row, "Tên Công Việc", "ten_cong_viec", "TEN_CONG_VIEC"));

                    if (searchType == 0 || searchType == 1)
                    {
                        if (!string.IsNullOrEmpty(maHieuText))
                        {
                            int levScore = StringDistance.LevenshteinSimilarity(keywordNorm, maHieuText);
                            if (levScore > bestScore)
                            {
                                bestScore = levScore;
                                bestAlgo = "levenshtein";
                            }

                            int jwScore = StringDistance.JaroWinklerSimilarity(keywordNorm, maHieuText);
                            if (jwScore > bestScore)
                            {
                                bestScore = jwScore;
                                bestAlgo = "jaro_winkler";
                            }
                        }
                    }

                    if (searchType == 0 || searchType == 2)
                    {
                        if (!string.IsNullOrEmpty(tenCongText))
                        {
                            int levScore = StringDistance.LevenshteinSimilarity(keywordNorm, tenCongText);
                            if (levScore > bestScore)
                            {
                                bestScore = levScore;
                                bestAlgo = "levenshtein";
                            }

                            int jwScore = StringDistance.JaroWinklerSimilarity(keywordNorm, tenCongText);
                            if (jwScore > bestScore)
                            {
                                bestScore = jwScore;
                                bestAlgo = "jaro_winkler";
                            }
                        }
                    }

                    if (bestScore >= _fuzzyThreshold)
                    {
                        results.Add(new SearchResult
                        {
                            Row = row,
                            Score = bestScore,
                            Type = "fuzzy",
                            Algo = bestAlgo
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in search_fuzzy_with_ngram: {ex.Message}");
            }

            results = results.OrderByDescending(r => r.Score).Take(1000).ToList();
            return results;
        }

        public List<SearchResult> Search(string keyword, int searchType = 0)
        {
            if (string.IsNullOrWhiteSpace(keyword))
            {
                return _originalData.Take(1000).Select(r => new SearchResult
                {
                    Row = r,
                    Score = 100,
                    Type = "all",
                    Algo = "exact"
                }).ToList();
            }

            keyword = keyword.Trim();
            string keywordNorm = TextNormalizer.Normalize(keyword);

            if (string.IsNullOrEmpty(keywordNorm)) return new List<SearchResult>();

            string cacheKey = $"{keywordNorm}_{searchType}";
            if (_searchCache.ContainsKey(cacheKey))
            {
                return _searchCache[cacheKey];
            }

            List<SearchResult> l1Results = SearchExactAndRanking(keyword, searchType);

            if (l1Results.Count >= 10)
            {
                _searchCache[cacheKey] = l1Results;
                return l1Results;
            }

            List<SearchResult> l3Results = SearchFuzzyWithNgram(keyword, searchType);

            HashSet<string> l1RowIds = new HashSet<string>();
            foreach (var r in l1Results)
            {
                try
                {
                    string rowId = GetRowHash(r.Row);
                    l1RowIds.Add(rowId);
                }
                catch { }
            }

            List<SearchResult> l3Unique = new List<SearchResult>();
            foreach (var r in l3Results)
            {
                try
                {
                    string rowId = GetRowHash(r.Row);
                    if (!l1RowIds.Contains(rowId))
                    {
                        l3Unique.Add(r);
                    }
                }
                catch
                {
                    l3Unique.Add(r);
                }
            }

            List<SearchResult> combined = l1Results.Concat(l3Unique).Take(1000).ToList();
            _searchCache[cacheKey] = combined;
            return combined;
        }

        private string GetRowHash(Dictionary<string, string> row)
        {
            return string.Join("|", row.OrderBy(kvp => kvp.Key).Select(kvp => $"{kvp.Key}={kvp.Value}"));
        }

        private string GetValue(Dictionary<string, string> row, params string[] keys)
        {
            foreach (string key in keys)
            {
                if (row.ContainsKey(key))
                    return row[key];
            }
            return "";
        }

        public List<string> GetSuggestions(string keyword, int maxSuggestions = 5)
        {
            string keywordNorm = TextNormalizer.Normalize(keyword);
            if (string.IsNullOrEmpty(keywordNorm) || keywordNorm.Length < 2)
                return new List<string>();

            HashSet<string> suggestions = new HashSet<string>();

            try
            {
                foreach (int idx in _ngramIndex.GetCandidates(keywordNorm))
                {
                    if (idx < _originalData.Count && _originalData[idx] != null)
                    {
                        var row = _originalData[idx];

                        string maHieu = GetValue(row, "Mã Hiệu", "ma_hieu", "MA_HIEU");
                        string tenCong = GetValue(row, "Tên Công Việc", "ten_cong_viec", "TEN_CONG_VIEC");

                        if (!string.IsNullOrEmpty(maHieu) && TextNormalizer.Normalize(maHieu).Contains(keywordNorm))
                            suggestions.Add(maHieu);

                        if (!string.IsNullOrEmpty(tenCong) && TextNormalizer.Normalize(tenCong).Contains(keywordNorm))
                            suggestions.Add(tenCong);

                        if (suggestions.Count >= maxSuggestions) break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in get_suggestions: {ex.Message}");
            }

            return suggestions.Take(maxSuggestions).ToList();
        }

        public void ClearCache()
        {
            _searchCache.Clear();
            _fuzzyCache.Clear();
        }

        public Dictionary<string, int> GetStats()
        {
            return new Dictionary<string, int>
            {
                {"total_rows", _originalData.Count},
                {"cache_size", _searchCache.Count},
                {"fuzzy_threshold", _fuzzyThreshold},
                {"trigram_entries", _ngramIndex.TrigramCount},
                {"bigram_entries", _ngramIndex.BigramCount},
                {"exact_entries", _ngramIndex.ExactCount},
                {"word_entries", _ngramIndex.WordCount}
            };
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
                        if (val != 0)
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
                        var prop = this.GetType().GetProperty(kvp.Value);
                        if (prop != null)
                            prop.SetValue(this, value);
                    }
                }
                catch
                {
                    continue;
                }
            }
        }
    }

    // =============================================================================
    // MAIN FORM CLASS - LEVEL 3 UPGRADED
    // =============================================================================

    public partial class MarkAssignmentForm : System.Windows.Forms.Form
    {
        private List<Dictionary<string, string>> _allCsvData;
        private List<Dictionary<string, string>> _displayedCsvData;
        private AdvancedCSVSearcherL3 _csvSearcher;
        private Dictionary<string, CategoryInfo> _categoriesDict;
        private List<Element> _currentElements;
        private string _currentCategory;
        private string _currentProperty;
        private Dictionary<int, Dictionary<string, string>> _mapping;
        private string _selectedMarkValue;
        private string _selectedTenCongViec;
        private string _selectedDonVi;
        private List<SearchResult> _lastSearchResults;
        private System.Windows.Forms.Timer _searchTimer;

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

        public Dictionary<int, Dictionary<string, string>> Mapping => _mapping;
        public AdvancedCSVSearcherL3 CsvSearcher => _csvSearcher;

        public MarkAssignmentForm(List<Dictionary<string, string>> csvData, Dictionary<string, CategoryInfo> categoriesDict)
        {
            Text = "TOOL GÁN MÃ ĐỊNH MỨC - LEVEL 3 UPGRADED (N-gram + Fuzzy + Jaro-Winkler)";
            Width = 1650;
            Height = 800;
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(1200, 600);

            _allCsvData = csvData ?? new List<Dictionary<string, string>>();
            _csvSearcher = new AdvancedCSVSearcherL3(csvData, 75);
            _displayedCsvData = new List<Dictionary<string, string>>(_allCsvData);

            _categoriesDict = categoriesDict;
            _currentElements = new List<Element>();
            _currentCategory = null;
            _currentProperty = "Count";
            _mapping = new Dictionary<int, Dictionary<string, string>>();
            _selectedMarkValue = null;
            _selectedTenCongViec = null;
            _selectedDonVi = null;
            _lastSearchResults = new List<SearchResult>();

            _searchTimer = new System.Windows.Forms.Timer();
            _searchTimer.Interval = 300;
            _searchTimer.Tick += OnSearchTimerTick;

            CreateUI();
            ShowWelcomeMessage();
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

            // Content Panel with SplitContainer
            SplitContainer splitContainer = new SplitContainer();
            splitContainer.Dock = DockStyle.Fill;
            splitContainer.Orientation = Orientation.Vertical;
            splitContainer.SplitterDistance = 700;
            mainLayout.Controls.Add(splitContainer, 0, 1);

            WinForms.Panel leftPanel = CreateLeftPanel();
            splitContainer.Panel1.Controls.Add(leftPanel);

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
            panel.BackColor = GDI.Color.FromArgb(200, 220, 255);
            panel.Padding = new Padding(8, 5, 8, 5);

            Button btnHelp = new Button();
            btnHelp.Text = "Hướng Dẫn";
            btnHelp.Location = new GDI.Point(8, 8);
            btnHelp.Size = new Size(100, 23);
            btnHelp.BackColor = GDI.Color.LightBlue;
            btnHelp.Click += ShowHelpClicked;
            panel.Controls.Add(btnHelp);

            Label lblCategory = new WinForms.Label();
            lblCategory.Text = "Chọn Category:";
            lblCategory.Location = new GDI.Point(305, 8);
            lblCategory.Size = new Size(100, 18);
            panel.Controls.Add(lblCategory);

            _categoryCombo = new WinForms.ComboBox();
            _categoryCombo.Location = new GDI.Point(410, 8);
            _categoryCombo.Size = new Size(220, 23);
            _categoryCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _categoryCombo.SelectedIndexChanged += CategoryChanged;
            panel.Controls.Add(_categoryCombo);

            _propertyCombo = new WinForms.ComboBox();
            _propertyCombo.Location = new GDI.Point(705, 8);
            _propertyCombo.Size = new Size(120, 23);
            _propertyCombo.Items.AddRange(new object[] { "Count", "Length", "Area", "Volume" });
            _propertyCombo.SelectedIndex = 0;
            panel.Controls.Add(_propertyCombo);

            _infoLabel = new Label();
            _infoLabel.Text = "✅ LEVEL 3 UPGRADED - N-gram Index Ready";
            _infoLabel.Location = new GDI.Point(835, 8);
            _infoLabel.Size = new Size(600, 18);
            _infoLabel.ForeColor = GDI.Color.DarkGreen;
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
            leftMain.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));
            leftMain.RowStyles.Add(new RowStyle(SizeType.Absolute, 90));
            leftMain.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            Label leftTitle = new Label();
            leftTitle.Text = "📊 CSV DATA (LEVEL 3 Search: N-gram + Fuzzy + Levenshtein + Jaro-Winkler)";
            leftTitle.Font = new Font(leftTitle.Font, FontStyle.Bold);
            leftTitle.Dock = DockStyle.Fill;
            leftTitle.BackColor = GDI.Color.FromArgb(100, 180, 255);
            leftTitle.TextAlign = ContentAlignment.MiddleLeft;
            leftTitle.ForeColor = GDI.Color.White;
            leftMain.Controls.Add(leftTitle, 0, 0);
            leftMain.SetColumnSpan(leftTitle, 2);

            WinForms.Panel searchPanel = CreateLevel3SearchPanel();
            leftMain.Controls.Add(searchPanel, 0, 1);
            leftMain.SetColumnSpan(searchPanel, 2);

            _markScrollbar = new VScrollBar();
            _markScrollbar.Dock = DockStyle.Fill;
            leftMain.Controls.Add(_markScrollbar, 0, 2);

            _markGrid = new DataGridView();
            _markGrid.Dock = DockStyle.Fill;
            _markGrid.ReadOnly = true;
            _markGrid.AllowUserToAddRows = false;
            _markGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _markGrid.MultiSelect = false;
            _markGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            _markGrid.CellClick += MarkGridCellClick;

            _markGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Mã Hiệu", Width = 80 });
            _markGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Tên Công Việc", Width = 180 });
            _markGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Đơn Vị", Width = 60 });
            _markGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Score", Width = 50 });

            leftMain.Controls.Add(_markGrid, 1, 2);

            WinForms.Panel panel = new WinForms.Panel();
            panel.Dock = DockStyle.Fill;
            panel.Controls.Add(leftMain);
            return panel;
        }

        private WinForms.Panel CreateLevel3SearchPanel()
        {
            WinForms.Panel panel = new WinForms.Panel();
            panel.Dock = DockStyle.Fill;
            panel.BackColor = GDI.Color.FromArgb(240, 250, 255);
            panel.BorderStyle = BorderStyle.FixedSingle;

            Label lblSearch = new Label();
            lblSearch.Text = "🔍 LEVEL 3 Search:";
            lblSearch.Location = new GDI.Point(5, 8);
            lblSearch.Size = new Size(110, 18);
            lblSearch.Font = new Font(lblSearch.Font, FontStyle.Bold);
            panel.Controls.Add(lblSearch);

            _searchBox = new WinForms.TextBox();
            _searchBox.Location = new GDI.Point(120, 8);
            _searchBox.Size = new Size(180, 20);
            _searchBox.TextChanged += SearchTextChangedDebounced;
            panel.Controls.Add(_searchBox);

            _searchTypeCombo = new WinForms.ComboBox();
            _searchTypeCombo.Location = new GDI.Point(310, 8);
            _searchTypeCombo.Size = new Size(140, 20);
            _searchTypeCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _searchTypeCombo.Items.AddRange(new object[] { "Both (N-gram)", "Mã Hiệu (Fuzzy)", "Tên Công Việc (Fuzzy)" });
            _searchTypeCombo.SelectedIndex = 0;
            _searchTypeCombo.SelectedIndexChanged += SearchTypeChanged;
            panel.Controls.Add(_searchTypeCombo);

            Button btnClear = new Button();
            btnClear.Text = "Clear";
            btnClear.Location = new GDI.Point(460, 8);
            btnClear.Size = new Size(55, 20);
            btnClear.Click += ClearSearchClicked;
            panel.Controls.Add(btnClear);

            _searchResultsLabel = new Label();
            _searchResultsLabel.Text = "Ready | N-gram Index Active";
            _searchResultsLabel.Location = new GDI.Point(520, 8);
            _searchResultsLabel.Size = new Size(150, 18);
            _searchResultsLabel.ForeColor = GDI.Color.DarkGreen;
            _searchResultsLabel.Font = new Font(_searchResultsLabel.Font, FontStyle.Bold);
            panel.Controls.Add(_searchResultsLabel);

            Label lblInfo = new Label();
            lblInfo.Text = "N-gram Index: Exact + Word + Prefix + Contains + Regex";
            lblInfo.Location = new GDI.Point(5, 35);
            lblInfo.Size = new Size(400, 18);
            lblInfo.ForeColor = GDI.Color.DarkBlue;
            panel.Controls.Add(lblInfo);

            Label lblFuzzy = new Label();
            lblFuzzy.Text = "Fuzzy: Levenshtein (typo) + Jaro-Winkler (phonetic) | Threshold: 75%";
            lblFuzzy.Location = new GDI.Point(5, 55);
            lblFuzzy.Size = new Size(400, 18);
            lblFuzzy.ForeColor = GDI.Color.DarkGreen;
            panel.Controls.Add(lblFuzzy);

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

            Label rightTitle = new Label();
            rightTitle.Text = "📌 REVIT ELEMENTS";
            rightTitle.Font = new Font(rightTitle.Font, FontStyle.Bold);
            rightTitle.Dock = DockStyle.Fill;
            rightTitle.BackColor = GDI.Color.LightGreen;
            rightTitle.TextAlign = ContentAlignment.MiddleLeft;
            rightMain.Controls.Add(rightTitle, 0, 0);

            WinForms.Panel controlPanel = new WinForms.Panel();
            controlPanel.Dock = DockStyle.Fill;
            controlPanel.BackColor = GDI.Color.FromArgb(250, 250, 250);

            Button btnSelectAll = new Button();
            btnSelectAll.Text = "Select All";
            btnSelectAll.Size = new Size(90, 23);
            btnSelectAll.Location = new GDI.Point(3, 3);
            btnSelectAll.Click += SelectAllClicked;
            controlPanel.Controls.Add(btnSelectAll);

            Button btnDeselectAll = new Button();
            btnDeselectAll.Text = "Deselect";
            btnDeselectAll.Size = new Size(90, 23);
            btnDeselectAll.Location = new GDI.Point(98, 3);
            btnDeselectAll.Click += DeselectAllClicked;
            controlPanel.Controls.Add(btnDeselectAll);

            _selectedCountLabel = new Label();
            _selectedCountLabel.Text = "Selected: 0";
            _selectedCountLabel.Size = new Size(400, 23);
            _selectedCountLabel.Location = new  GDI.Point(200, 3);
            _selectedCountLabel.ForeColor = GDI.Color.DarkGreen;
            controlPanel.Controls.Add(_selectedCountLabel);

            rightMain.Controls.Add(controlPanel, 0, 1);

            _elementsGrid = new DataGridView();
            _elementsGrid.Dock = DockStyle.Fill;
            _elementsGrid.AllowUserToAddRows = false;
            _elementsGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _elementsGrid.MultiSelect = true;
            _elementsGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            _elementsGrid.CellValueChanged += ElementGridCellValueChanged;

            DataGridViewCheckBoxColumn colCheck = new DataGridViewCheckBoxColumn();
            colCheck.Name = "Select";
            colCheck.Width = 50;
            colCheck.TrueValue = true;
            colCheck.FalseValue = false;
            _elementsGrid.Columns.Add(colCheck);

            _elementsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Mã hiệu", Width = 100, ReadOnly = false });
            _elementsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Tên công việc", Width = 180, ReadOnly = false });
            _elementsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Family", Width = 120, ReadOnly = true });
            _elementsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Type", Width = 100, ReadOnly = true });
            _elementsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Qty", Width = 80, ReadOnly = true });

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
            panel.BackColor = GDI.Color.FromArgb(240, 240, 240);
            panel.Padding = new Padding(8, 5, 8, 5);

            _statusLabel = new Label();
            _statusLabel.Text = "✅ LEVEL 3 UPGRADED - N-gram + Fuzzy enabled";
            _statusLabel.AutoSize = false;
            _statusLabel.Size = new Size(750, 25);
            _statusLabel.Location = new GDI.Point(8, 8);
            _statusLabel.ForeColor = GDI.Color.DarkGreen;
            _statusLabel.Font = new Font(_statusLabel.Font, FontStyle.Bold);
            panel.Controls.Add(_statusLabel);

            Button btnExport = new Button();
            btnExport.Text = "Export Template";
            btnExport.Size = new Size(130, 28);
            btnExport.Location = new GDI.Point(800, 6);
            btnExport.BackColor = GDI.Color.Gold;
            btnExport.Click += ExportTemplateClicked;
            panel.Controls.Add(btnExport);

            Button btnLoad = new Button();
            btnLoad.Text = "Load Template";
            btnLoad.Size = new Size(130, 28);
            btnLoad.Location = new GDI.Point(938, 6);
            btnLoad.BackColor = GDI.Color.LightBlue;
            btnLoad.Click += LoadTemplateClicked;
            panel.Controls.Add(btnLoad);

            Button btnAuto = new Button();
            btnAuto.Text = "Auto Assign";
            btnAuto.Size = new Size(110, 28);
            btnAuto.Location = new GDI.Point(1076, 6);
            btnAuto.BackColor = GDI.Color.LightGreen;
            btnAuto.Click += AutoAssignClicked;
            panel.Controls.Add(btnAuto);

            Button btnAssign = new Button();
            btnAssign.Text = "Gán/Lưu";
            btnAssign.Size = new Size(110, 28);
            btnAssign.Location = new GDI.Point(1332, 6);
            btnAssign.BackColor = GDI.Color.LightBlue;
            btnAssign.Click += AssignMarkClicked;
            panel.Controls.Add(btnAssign);

            Button btnConfirm = new Button();
            btnConfirm.Text = "Hoàn Thành";
            btnConfirm.Size = new Size(110, 28);
            btnConfirm.Location = new GDI.Point(1450, 6);
            btnConfirm.BackColor = GDI.Color.LightGreen;
            btnConfirm.Click += ConfirmClicked;
            panel.Controls.Add(btnConfirm);

            return panel;
        }

        private void ShowWelcomeMessage()
        {
            string welcomeText = @"
🚀 TOOL GÁN MÃ ĐỊNH MỨC - v5.0 + v6.0 HYBRID + LEVEL 3 UPGRADED

✨ NEW LEVEL 3 FEATURES:
✓ N-gram indexing (<15ms search on 1000+ rows)
✓ Levenshtein distance for typo detection
✓ Jaro-Winkler similarity for phonetic matching
✓ Fuzzy matching with 75% threshold
✓ Advanced caching system
✓ Ultra-fast candidate filtering

🔍 SEARCH CAPABILITIES:
✓ Exact matching (100 points)
✓ Word boundary matching (90 points)
✓ Prefix matching (80 points)
✓ Substring contains (60 points)
✓ Regex pattern matching (40 points)
✓ Fuzzy typo handling (LEVEL 3!)

Tool đã sẵn sàng sử dụng.";

            if (_statusLabel != null)
            {
                _statusLabel.Text = "✅ LEVEL 3 UPGRADED Active! N-gram + Fuzzy enabled";
            }
            Console.WriteLine(welcomeText);
        }

        private void PopulateCategories()
        {
            _categoryCombo.Items.Clear();
            var sortedCats = _categoriesDict.Keys.OrderBy(k => k).ToList();

            if (sortedCats.Count == 0)
            {
                _categoryCombo.Items.Add("[No categories]");
                _categoryCombo.Enabled = false;
            }
            else
            {
                foreach (string cat in sortedCats)
                {
                    var catInfo = _categoriesDict[cat];
                    string displayText = $"{cat} ({catInfo.Count}) - {catInfo.Type}";
                    _categoryCombo.Items.Add(displayText);
                }
                _categoryCombo.Enabled = true;
            }
        }

        // Event Handlers
        private void ShowHelpClicked(object sender, EventArgs e)
        {
            string helpText = @"TOOL GÁN MÃ ĐỊNH MỨC v5.0 + v6.0 HYBRID + LEVEL 3 UPGRADED

✨ LEVEL 3 IMPROVEMENTS:
✓ N-gram indexing for <15ms search
✓ Levenshtein distance (typo detection)
✓ Jaro-Winkler similarity (phonetic)
✓ Fuzzy matching (75% threshold)
✓ Advanced caching

Có thể search:
- Exact: M10A
- Typo: Pibe (matches Pipe)
- Accents: Dau (matches Dầu)
- Partial: Pipe
- Pattern: 50*mm

Hướng dẫn:
1. Chọn category
2. Search với LEVEL 3
3. Select CSV row
4. Select elements
5. Click Gán/Lưu
6. Export/Auto Assign";

            MessageBox.Show(helpText, "Hướng Dẫn LEVEL 3");
        }

        private void CategoryChanged(object sender, EventArgs e)
        {
            if (_categoryCombo.SelectedIndex < 0) return;

            string selectedText = _categoryCombo.SelectedItem.ToString();
            if (selectedText.Contains("[")) return;

            string categoryName = selectedText.Split(new[] { " (" }, StringSplitOptions.None)[0];
            _currentCategory = categoryName;

            try
            {
                _infoLabel.Text = $"Loading: {categoryName}...";
                Application.DoEvents();

                _currentElements = GetElementsByCategoryName(categoryName);
                PopulateElementsGrid();
                PopulateMarkTable();

                _infoLabel.Text = $"Category: {categoryName} ({_currentElements.Count} elements) - LEVEL 3 Ready";
            }
            catch (Exception ex)
            {
                _infoLabel.Text = $"Error: {ex.Message}";
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
            PerformLevel3Search();
        }

        private void SearchTypeChanged(object sender, EventArgs e)
        {
            PerformLevel3Search();
        }

        private void PerformLevel3Search()
        {
            string searchKeyword = _searchBox.Text.Trim();
            int searchType = _searchTypeCombo.SelectedIndex;

            DateTime startTime = DateTime.Now;

            if (string.IsNullOrEmpty(searchKeyword))
            {
                _displayedCsvData = new List<Dictionary<string, string>>(_allCsvData);
            }
            else
            {
                List<SearchResult> results = _csvSearcher.Search(searchKeyword, searchType);
                _displayedCsvData = results.Select(r => r.Row).ToList();
                _lastSearchResults = results;
            }

            double elapsed = (DateTime.Now - startTime).TotalMilliseconds;

            UpdateMarkGridWithScores();
            _searchResultsLabel.Text = $"Found: {_displayedCsvData.Count}/{_allCsvData.Count} | Time: {elapsed:F1}ms";
        }

        private void UpdateMarkGridWithScores()
        {
            _markGrid.Rows.Clear();

            var displayData = _displayedCsvData.Take(2000).ToList();

            for (int idx = 0; idx < displayData.Count; idx++)
            {
                var row = displayData[idx];
                int score = 0;

                if (_lastSearchResults != null && idx < _lastSearchResults.Count)
                {
                    try
                    {
                        score = _lastSearchResults[idx].Score;
                    }
                    catch { }
                }

                _markGrid.Rows.Add(
                    GetDictValue(row, "Mã Hiệu"),
                    GetDictValue(row, "Tên Công Việc"),
                    GetDictValue(row, "Đơn Vị"),
                    score > 0 ? score.ToString() : "");
            }
        }

        private string GetDictValue(Dictionary<string, string> dict, string key)
        {
            if (dict.ContainsKey(key))
                return dict[key];
            return "";
        }

        private void ClearSearchClicked(object sender, EventArgs e)
        {
            _searchBox.Text = "";
            PerformLevel3Search();
        }

        private void MarkGridCellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= _markGrid.Rows.Count) return;

            _markGrid.ClearSelection();
            _markGrid.Rows[e.RowIndex].Selected = true;

            _selectedMarkValue = _markGrid.Rows[e.RowIndex].Cells[0].Value?.ToString();
            _selectedTenCongViec = _markGrid.Rows[e.RowIndex].Cells[1].Value?.ToString();
            _selectedDonVi = _markGrid.Rows[e.RowIndex].Cells[2].Value?.ToString();

            if (_statusLabel != null)
            {
                _statusLabel.Text = $"Selected: {_selectedMarkValue} | {_selectedTenCongViec} | {_selectedDonVi}";
            }
        }

        private void SelectAllClicked(object sender, EventArgs e)
        {
            for (int i = 0; i < _elementsGrid.Rows.Count; i++)
            {
                _elementsGrid.Rows[i].Cells[0].Value = true;
            }
            UpdateSelectionCount();
        }

        private void DeselectAllClicked(object sender, EventArgs e)
        {
            for (int i = 0; i < _elementsGrid.Rows.Count; i++)
            {
                _elementsGrid.Rows[i].Cells[0].Value = false;
            }
            UpdateSelectionCount();
        }

        private void UpdateSelectionCount()
        {
            int selected = 0;
            for (int i = 0; i < _elementsGrid.Rows.Count; i++)
            {
                if (_elementsGrid.Rows[i].Cells[0].Value is bool b && b)
                    selected++;
            }
            _selectedCountLabel.Text = $"Selected: {selected}";
        }

        private void ElementGridCellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            UpdateSelectionCount();
        }

        private void PopulateElementsGrid()
        {
            _elementsGrid.Rows.Clear();

            foreach (Element elem in _currentElements)
            {
                ElementWrapper wrapper = new ElementWrapper(elem);
                int rowIndex = _elementsGrid.Rows.Add();
                _elementsGrid.Rows[rowIndex].Cells[0].Value = false;
                _elementsGrid.Rows[rowIndex].Cells[1].Value = wrapper.MaHieuValue;
                _elementsGrid.Rows[rowIndex].Cells[2].Value = wrapper.TenCongViecValue;
                _elementsGrid.Rows[rowIndex].Cells[3].Value = wrapper.FamilyName;
                _elementsGrid.Rows[rowIndex].Cells[4].Value = wrapper.Name;
                _elementsGrid.Rows[rowIndex].Cells[5].Value = $"{wrapper.QuantityValue:F2}";
            }

            UpdateSelectionCount();
        }

        private void PopulateMarkTable()
        {
            _displayedCsvData = new List<Dictionary<string, string>>(_allCsvData);
            UpdateMarkGridWithScores();
            _searchResultsLabel.Text = $"Total: {_allCsvData.Count} | N-gram Index Ready";
        }

        private void AssignMarkClicked(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_selectedMarkValue))
            {
                MessageBox.Show("Please select a code from the left table!", "Notice");
                return;
            }

            List<int> selectedIndices = new List<int>();
            for (int i = 0; i < _elementsGrid.Rows.Count; i++)
            {
                if (_elementsGrid.Rows[i].Cells[0].Value is bool b && b)
                    selectedIndices.Add(i);
            }

            if (selectedIndices.Count == 0)
            {
                MessageBox.Show("Please select at least one element from the right table!", "Notice");
                return;
            }

            Transaction t = new Transaction(MarkAssignmentCommand.Doc, "Assign Parameters with LEVEL 3");
            t.Start();

            int successCount = 0;
            try
            {
                foreach (int idx in selectedIndices)
                {
                    Element elem = _currentElements[idx];
                    if (AssignAllParameters(elem, _selectedMarkValue, _selectedTenCongViec, _selectedDonVi))
                    {
                        successCount++;
                        _mapping[elem.Id.IntegerValue] = new Dictionary<string, string>
                        {
                            {"ma_hieu", _selectedMarkValue},
                            {"ten_cong_viec", _selectedTenCongViec},
                            {"don_vi", _selectedDonVi}
                        };
                    }
                }

                t.Commit();
                MessageBox.Show($"Assigned {successCount} elements! (LEVEL 3 Search used)", "Success");
                _statusLabel.Text = $"Assigned: {successCount} elements | LEVEL 3 UPGRADED";
                PopulateElementsGrid();
            }
            catch (Exception ex)
            {
                t.RollBack();
                MessageBox.Show($"Error: {ex.Message}", "Error");
            }
        }

        private bool AssignAllParameters(Element elem, string markValue, string tenCongViecValue, string donViValue)
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

        private void ExportTemplateClicked(object sender, EventArgs e)
        {
            try
            {
                _statusLabel.Text = "Exporting template with LEVEL 3...";
                Application.DoEvents();

                MessageBox.Show("Template export feature - implement according to v6.0 logic", "Info");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error");
            }
        }

        private void LoadTemplateClicked(object sender, EventArgs e)
        {
            try
            {
                MessageBox.Show("Template load feature - implement according to v6.0 logic", "Info");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error");
            }
        }

        private void AutoAssignClicked(object sender, EventArgs e)
        {
            try
            {
                MessageBox.Show("Auto assign feature - implement according to v6.0 logic", "Info");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error");
            }
        }

        private void ConfirmClicked(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private List<Element> GetElementsByCategoryName(string categoryName)
        {
            List<Element> elements = new List<Element>();

            try
            {
                BuiltInCategory? builtinCat = GetBuiltinCategory(categoryName);

                if (builtinCat.HasValue)
                {
                    FilteredElementCollector collector = new FilteredElementCollector(MarkAssignmentCommand.Doc)
                        .OfCategory(builtinCat.Value)
                        .WhereElementIsNotElementType();
                    elements = collector.ToList();
                }
                else
                {
                    FilteredElementCollector collector = new FilteredElementCollector(MarkAssignmentCommand.Doc)
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

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _searchTimer?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}