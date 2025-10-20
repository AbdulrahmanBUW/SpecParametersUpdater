using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace SpecParametersUpdater
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class UpdateSpecParametersCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            var stopwatch = Stopwatch.StartNew();

            try
            {
                // Check if we're in a family document
                bool isFamilyDoc = doc.IsFamilyDocument;

                var categories = Categories.GetAll();
                var selection = uidoc.Selection?.GetElementIds();
                bool useSelection = selection != null && selection.Count > 0;

                var elementsToProcess = isFamilyDoc
                    ? CollectFamilyElements(doc, selection, useSelection)
                    : CollectElements(doc, selection, useSelection, categories);

                var stats = new UpdateStats();

                using (var trans = new Transaction(doc, "Update SPEC_SYSTEM, SPEC_SIZE, SPEC_QUANTITY"))
                {
                    trans.Start();

                    // Create missing parameters in family documents (must be inside transaction)
                    if (isFamilyDoc)
                    {
                        CreateMissingFamilyParameters(doc, stats);
                    }

                    // Create cache after parameters are created
                    var cache = new ParameterCache(doc, isFamilyDoc);

                    // Validate parameters after creation
                    if (Config.ValidateParameters)
                    {
                        ValidateParameters(doc, stats.Warnings, isFamilyDoc);
                    }

                    UpdateAllElements(elementsToProcess, cache, stats, categories);
                    trans.Commit();
                }

                stopwatch.Stop();
                ShowResults(doc, stats, stopwatch, useSelection, isFamilyDoc);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error: {ex.Message}";
                return Result.Failed;
            }
        }

        private Dictionary<ElementId, Element> CollectFamilyElements(Document doc, ICollection<ElementId> selection, bool useSelection)
        {
            var result = new Dictionary<ElementId, Element>();

            if (useSelection)
            {
                // Use selected elements
                foreach (var id in selection)
                {
                    var el = doc.GetElement(id);
                    if (el != null)
                        result[id] = el;
                }
            }
            else
            {
                // Collect all elements in the family (not types)
                var collector = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType();

                foreach (Element el in collector)
                {
                    // Skip reference planes, levels, and other non-geometry elements
                    if (el.Category == null)
                        continue;

                    var catId = el.Category.Id.IntegerValue;

                    // Skip invalid categories
                    if (catId < 0)
                        continue;

                    // Include the element
                    result[el.Id] = el;
                }
            }

            return result;
        }

        private Dictionary<ElementId, Element> CollectElements(Document doc, ICollection<ElementId> selection,
            bool useSelection, BuiltInCategory[] categories)
        {
            var elementIds = new HashSet<ElementId>();

            if (useSelection)
            {
                foreach (var id in selection) elementIds.Add(id);
            }
            else
            {
                foreach (var cat in categories)
                {
                    try
                    {
                        var ids = new FilteredElementCollector(doc)
                            .OfCategory(cat)
                            .WhereElementIsNotElementType()
                            .ToElementIds();
                        foreach (var id in ids) elementIds.Add(id);
                    }
                    catch { }
                }
            }

            var result = new Dictionary<ElementId, Element>();
            foreach (var id in elementIds)
            {
                var el = doc.GetElement(id);
                if (el != null) result[id] = el;
            }
            return result;
        }

        private void UpdateAllElements(Dictionary<ElementId, Element> elements, ParameterCache cache,
            UpdateStats stats, BuiltInCategory[] categories)
        {
            var sizeCategories = Categories.GetSizeQuantityCategories();

            foreach (var kv in elements)
            {
                var element = kv.Value;
                if (element == null) continue;

                UpdateSystem(element, cache, stats);
                UpdateSize(element, cache, stats, sizeCategories);
                UpdateQuantity(element, cache, stats, sizeCategories);
            }
        }

        private void UpdateSystem(Element element, ParameterCache cache, UpdateStats stats)
        {
            try
            {
                var param = cache.GetInstanceOrType(element, SpecParams.SYSTEM);
                if (param == null || param.IsReadOnly) return;

                string value = GetSystemValue(element, cache);
                if (string.IsNullOrWhiteSpace(value)) return;

                if (ParameterHelper.UpdateString(param, element, value, cache))
                    stats.SystemUpdates++;
            }
            catch (Exception ex)
            {
                stats.Errors.Add($"[{DateTime.Now:HH:mm:ss}] SPEC_SYSTEM failed for {element.Id}: {ex.Message}");
            }
        }

        private string GetSystemValue(Element element, ParameterCache cache)
        {
            var abbrev = cache.GetInstanceOrType(element, "System Abbreviation");
            if (ParameterHelper.TryGetString(abbrev, out var val)) return val.Trim();

            // In families, RBS_SYSTEM_NAME_PARAM might not exist
            try
            {
                var name = element.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM);
                if (ParameterHelper.TryGetString(name, out val)) return val.Trim();
            }
            catch { }

            try
            {
                var classif = element.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM);
                if (ParameterHelper.TryGetString(classif, out val)) return val.Trim();
            }
            catch { }

            return null;
        }

        private void UpdateSize(Element element, ParameterCache cache, UpdateStats stats,
            HashSet<BuiltInCategory> sizeCategories)
        {
            try
            {
                if (element.Category == null) return;

                var catId = (BuiltInCategory)element.Category.Id.IntegerValue;
                if (!sizeCategories.Contains(catId)) return;

                var sizeParam = FindSizeParameter(element, cache);
                var specSize = cache.GetInstanceOrType(element, SpecParams.SIZE);

                if (specSize == null || specSize.IsReadOnly || sizeParam == null || !sizeParam.HasValue)
                    return;

                string rawSize = ParameterHelper.GetString(sizeParam);
                if (string.IsNullOrWhiteSpace(rawSize)) return;

                string formatted = SizeFormatter.Format(rawSize, element);
                if (ParameterHelper.UpdateString(specSize, element, formatted, cache))
                    stats.SizeUpdates++;
            }
            catch (Exception ex)
            {
                stats.Errors.Add($"[{DateTime.Now:HH:mm:ss}] SPEC_SIZE failed for {element.Id}: {ex.Message}");
            }
        }

        private void UpdateQuantity(Element element, ParameterCache cache, UpdateStats stats,
            HashSet<BuiltInCategory> sizeCategories)
        {
            try
            {
                if (element.Category == null) return;
                var catId = (BuiltInCategory)element.Category.Id.IntegerValue;
                if (!sizeCategories.Contains(catId)) return;

                Parameter lengthParam = null;

                try
                {
                    lengthParam = element.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                }
                catch { }

                if (lengthParam == null)
                    lengthParam = cache.GetInstanceOrType(element, "Length");

                var specQty = cache.GetInstanceOrType(element, SpecParams.QUANTITY);

                if (specQty == null || specQty.IsReadOnly || lengthParam == null || !lengthParam.HasValue)
                    return;

                double lengthFeet = ParameterHelper.GetDouble(lengthParam);
                double lengthMeters = lengthFeet * 0.3048;
                double finalQty = lengthMeters > 1.0 ? Math.Round(lengthMeters, 2) : 1.0;

                if (ParameterHelper.UpdateNumeric(specQty, element, finalQty, cache))
                    stats.QuantityUpdates++;
            }
            catch (Exception ex)
            {
                stats.Errors.Add($"[{DateTime.Now:HH:mm:ss}] SPEC_QUANTITY failed for {element.Id}: {ex.Message}");
            }
        }

        private Parameter FindSizeParameter(Element element, ParameterCache cache)
        {
            try
            {
                var p = element.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE);
                if (p != null && p.HasValue) return p;
            }
            catch { }

            string[] names = { "Size", "Diameter", "Width", "Height", "Tray Width", "Outside Diameter", "NW" };

            foreach (var name in names)
            {
                var p = element.LookupParameter(name);
                if (p != null && p.HasValue) return p;
            }

            foreach (var name in names)
            {
                var p = cache.GetTypeParameter(element, name);
                if (p != null && p.HasValue) return p;
            }

            return null;
        }

        private void ValidateParameters(Document doc, List<string> warnings, bool isFamilyDoc)
        {
            Element testElement = null;

            if (isFamilyDoc)
            {
                // In family documents, check family manager parameters
                var familyManager = doc.FamilyManager;
                if (familyManager != null)
                {
                    var missing = new List<string>();
                    foreach (var paramName in SpecParams.GetAll())
                    {
                        var param = familyManager.get_Parameter(paramName);
                        if (param == null)
                            missing.Add(paramName);
                    }

                    if (missing.Count > 0)
                    {
                        string msg = $"Missing {missing.Count} family parameters (will be created): " + string.Join(", ", missing.Take(10));
                        if (missing.Count > 10) msg += $" ... and {missing.Count - 10} more";
                        warnings.Add(msg);
                    }
                }
            }
            else
            {
                // Original project validation
                testElement = new FilteredElementCollector(doc).WhereElementIsNotElementType().FirstOrDefault();
                if (testElement == null) return;

                var missing = new List<string>();
                foreach (var param in SpecParams.GetAll())
                {
                    if (testElement.LookupParameter(param) == null)
                        missing.Add(param);
                }

                if (missing.Count > 0)
                {
                    string msg = $"Missing {missing.Count} parameters: " + string.Join(", ", missing.Take(10));
                    if (missing.Count > 10) msg += $" ... and {missing.Count - 10} more";
                    warnings.Add(msg);
                }
            }
        }

        private void CreateMissingFamilyParameters(Document doc, UpdateStats stats)
        {
            var familyManager = doc.FamilyManager;
            if (familyManager == null)
            {
                stats.Warnings.Add("FamilyManager not available");
                return;
            }

            // Use the shared parameter workaround for Revit 2023
            DefinitionFile defFile = doc.Application.OpenSharedParameterFile();
            if (defFile == null)
            {
                string tempFile = Path.Combine(Path.GetTempPath(), "Temp_SpecParams.txt");
                if (!File.Exists(tempFile))
                    File.WriteAllText(tempFile, "# Temp Shared Parameter File");
                doc.Application.SharedParametersFilename = tempFile;
                defFile = doc.Application.OpenSharedParameterFile();
            }

            var group = defFile.Groups.get_Item("SpecParameters") ?? defFile.Groups.Create("SpecParameters");

            var parameterDefinitions = new List<(string name, ForgeTypeId type, bool isInstance)>
            {
                (SpecParams.SYSTEM, SpecTypeId.String.Text, true),
                (SpecParams.POSITION, SpecTypeId.String.Text, true),
                (SpecParams.SYSTEM, SpecTypeId.String.Text, true),
                (SpecParams.SIZE, SpecTypeId.String.Text, true),
                (SpecParams.QUANTITY, SpecTypeId.Number, true),
                (SpecParams.POSITION, SpecTypeId.String.Text, true),
                (SpecParams.FILTER, SpecTypeId.String.Text, true),
                (SpecParams.SUPPLIER, SpecTypeId.String.Text, true),
                (SpecParams.WP, SpecTypeId.String.Text, true),
                (SpecParams.MATERIAL, SpecTypeId.String.Text, true),
                (SpecParams.MATERIAL_TEXT_EN, SpecTypeId.String.Text, true),
                (SpecParams.MATERIAL_TEXT_DE, SpecTypeId.String.Text, true),
                (SpecParams.UNITS_EN, SpecTypeId.String.Text, true),
                (SpecParams.UNITS_DE, SpecTypeId.String.Text, true),
                (SpecParams.NAME_EN, SpecTypeId.String.Text, true),
                (SpecParams.NAME_DE, SpecTypeId.String.Text, true),
                (SpecParams.CATEGORY_EN, SpecTypeId.String.Text, true),
                (SpecParams.CATEGORY_DE, SpecTypeId.String.Text, true),
                (SpecParams.NAME_SHORT_EN, SpecTypeId.String.Text, true),
                (SpecParams.NAME_SHORT_DE, SpecTypeId.String.Text, true),
                (SpecParams.COMMENTS_EN, SpecTypeId.String.Text, true),
                (SpecParams.COMMENTS_DE, SpecTypeId.String.Text, true),
                (SpecParams.MANUFACTURER_EN, SpecTypeId.String.Text, true),
                (SpecParams.MANUFACTURER_DE, SpecTypeId.String.Text, true),
                (SpecParams.TYPE_EN, SpecTypeId.String.Text, true),
                (SpecParams.TYPE_DE, SpecTypeId.String.Text, true),
                (SpecParams.ARTICLE_EN, SpecTypeId.String.Text, true),
                (SpecParams.ARTICLE_DE, SpecTypeId.String.Text, true),
                (SpecParams.STATUS_CODE, SpecTypeId.String.Text, true),
                (SpecParams.TOOL_ID, SpecTypeId.String.Text, true)
            };

            foreach (var (paramName, typeId, isInstance) in parameterDefinitions)
            {
                try
                {
                    var existing = familyManager.get_Parameter(paramName);
                    if (existing != null)
                        continue;

                    // Look for existing shared definition
                    var def = group.Definitions
                        .OfType<ExternalDefinition>()
                        .FirstOrDefault(d => d.Name == paramName);

                    if (def == null)
                    {
                        var options = new ExternalDefinitionCreationOptions(paramName, typeId)
                        {
                            Visible = true
                        };
                        def = group.Definitions.Create(options) as ExternalDefinition;
                    }

                    if (def != null)
                    {
                        familyManager.AddParameter(def, BuiltInParameterGroup.PG_TEXT, isInstance);
                        stats.ParametersCreated++;
                        stats.CreatedParameters.Add(paramName);
                    }
                    else
                    {
                        stats.Warnings.Add($"Parameter {paramName} could not be added — not an ExternalDefinition.");
                    }
                }
                catch (Exception ex)
                {
                    stats.Errors.Add($"Failed to create parameter {paramName}: {ex.Message}");
                }
            }
        }



        private void ShowResults(Document doc, UpdateStats stats, Stopwatch stopwatch, bool useSelection, bool isFamilyDoc)
        {
            var summary = BuildSummary(doc, stats, stopwatch, useSelection, isFamilyDoc);

            string logPath = null;
            if (Config.WriteLogFile)
            {
                try
                {
                    logPath = Path.Combine(Path.GetTempPath(),
                        $"SpecParametersUpdater_{DateTime.Now:yyyyMMdd_HHmmss}.log");
                    File.WriteAllText(logPath, summary, Encoding.UTF8);
                }
                catch { }
            }

            var dialog = new TaskDialog("SPEC Parameters Update Complete");
            dialog.MainInstruction = "Update Completed";

            string docType = isFamilyDoc ? "FAMILY" : "PROJECT";
            string mode = useSelection ? "SELECTION" : (isFamilyDoc ? "ALL FAMILY ELEMENTS" : "ALL CATEGORIES");

            var contentBuilder = new StringBuilder();
            contentBuilder.AppendLine($"Document Type: {docType}");
            contentBuilder.AppendLine($"Mode: {mode}");
            contentBuilder.AppendLine();

            // Show parameters created (for families)
            if (isFamilyDoc && stats.ParametersCreated > 0)
            {
                contentBuilder.AppendLine($"Parameters Created: {stats.ParametersCreated}");
                contentBuilder.AppendLine();
            }

            contentBuilder.AppendLine($"SPEC_SYSTEM: {stats.SystemUpdates}");
            contentBuilder.AppendLine($"SPEC_SIZE: {stats.SizeUpdates}");
            contentBuilder.AppendLine($"SPEC_QUANTITY: {stats.QuantityUpdates}");
            contentBuilder.AppendLine();
            contentBuilder.AppendLine($"Total Updates: {stats.TotalUpdates}");
            contentBuilder.AppendLine($"Errors: {stats.Errors.Count}");
            contentBuilder.AppendLine($"Time: {stopwatch.Elapsed.TotalSeconds:F2}s");

            dialog.MainContent = contentBuilder.ToString();

            // Build expanded content
            var expandedBuilder = new StringBuilder();
            if (!string.IsNullOrEmpty(logPath))
            {
                expandedBuilder.AppendLine($"Log: {logPath}");
                expandedBuilder.AppendLine();
            }

            if (stats.ParametersCreated > 0)
            {
                expandedBuilder.AppendLine("Created Parameters:");
                foreach (var param in stats.CreatedParameters)
                {
                    expandedBuilder.AppendLine($"  • {param}");
                }
                expandedBuilder.AppendLine();
            }

            if (stats.Warnings.Count > 0)
            {
                expandedBuilder.AppendLine("Warnings:");
                foreach (var w in stats.Warnings.Take(10))
                {
                    expandedBuilder.AppendLine($"  • {w}");
                }
                if (stats.Warnings.Count > 10)
                    expandedBuilder.AppendLine($"  ... and {stats.Warnings.Count - 10} more (see log)");
            }

            if (expandedBuilder.Length > 0)
            {
                dialog.ExpandedContent = expandedBuilder.ToString();
            }

            dialog.Show();
        }

        private string BuildSummary(Document doc, UpdateStats stats, Stopwatch stopwatch, bool useSelection, bool isFamilyDoc)
        {
            var sb = new StringBuilder();
            sb.AppendLine("SPEC PARAMETERS UPDATE");
            sb.AppendLine($"Document: {doc.Title}");
            sb.AppendLine($"Document Type: {(isFamilyDoc ? "FAMILY" : "PROJECT")}");
            sb.AppendLine($"Mode: {(useSelection ? "SELECTION" : "ALL")}");
            sb.AppendLine($"Time: {stopwatch.Elapsed.TotalSeconds:F2}s");
            sb.AppendLine();

            // Add diagnostic info for families
            if (isFamilyDoc)
            {
                var collector = new FilteredElementCollector(doc).WhereElementIsNotElementType();
                var allElements = collector.ToList();
                sb.AppendLine($"Total elements found in family: {allElements.Count}");

                var categorized = allElements.GroupBy(e => e.Category?.Name ?? "No Category");
                foreach (var group in categorized)
                {
                    sb.AppendLine($"  {group.Key}: {group.Count()}");
                }
                sb.AppendLine();
            }

            // Show parameters created
            if (stats.ParametersCreated > 0)
            {
                sb.AppendLine($"Parameters Created: {stats.ParametersCreated}");
                foreach (var param in stats.CreatedParameters)
                {
                    sb.AppendLine($"  {param}");
                }
                sb.AppendLine();
            }

            sb.AppendLine($"SPEC_SYSTEM: {stats.SystemUpdates}");
            sb.AppendLine($"SPEC_SIZE: {stats.SizeUpdates}");
            sb.AppendLine($"SPEC_QUANTITY: {stats.QuantityUpdates}");
            sb.AppendLine($"Total Updates: {stats.TotalUpdates}");
            sb.AppendLine($"Errors: {stats.Errors.Count}");

            if (stats.Warnings.Count > 0)
            {
                sb.AppendLine("\nWARNINGS:");
                foreach (var w in stats.Warnings) sb.AppendLine($"  {w}");
            }

            if (stats.Errors.Count > 0)
            {
                sb.AppendLine("\nERRORS:");
                foreach (var e in stats.Errors.Take(50)) sb.AppendLine($"  {e}");
            }

            return sb.ToString();
        }
    }
}