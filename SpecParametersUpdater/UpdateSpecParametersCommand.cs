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

            // Store original shared parameter file path - CRITICAL TO RESTORE IT LATER
            string originalSharedParamFile = doc.Application.SharedParametersFilename;
            string tempSharedParamFile = null;

            try
            {
                // Create a temporary shared parameter file with exact GUIDs
                tempSharedParamFile = Path.Combine(Path.GetTempPath(), "SpecParams_" + Guid.NewGuid().ToString() + ".txt");

                // Build the shared parameter file content with exact GUIDs
                var sharedParamContent = new StringBuilder();
                sharedParamContent.AppendLine("# This is a Revit shared parameter file.");
                sharedParamContent.AppendLine("# Do not edit manually");
                sharedParamContent.AppendLine("*META\tVERSION\tMINVERSION");
                sharedParamContent.AppendLine("META\t2\t1");
                sharedParamContent.AppendLine("*GROUP\tID\tNAME");
                sharedParamContent.AppendLine("GROUP\t1\tSPEC_Parameters");
                sharedParamContent.AppendLine("*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\tVISIBLE\tDESCRIPTION\tUSERMODIFIABLE");

                // Add all parameters with their specific GUIDs - THESE GUIDS MUST NEVER CHANGE
                sharedParamContent.AppendLine("PARAM\t9f2f7b1a-517f-48a1-8c2b-bfec1948dc74\tSPEC_SIZE\tTEXT\t\t1\t1\t\t1");
                sharedParamContent.AppendLine("PARAM\t25d801a6-2723-4ed1-8e69-bab44cab7e2f\tSPEC_MATERIAL\tMATERIAL\t\t1\t1\t\t1");
                sharedParamContent.AppendLine("PARAM\tec43941f-42ec-4fd6-9826-033c28505638\tSPEC_MATERIAL_TEXT_EN\tTEXT\t\t1\t1\t\t1");
                sharedParamContent.AppendLine("PARAM\tb97c365e-4ad7-4ae2-903b-e6d7d1303d31\tSPEC_MATERIAL_TEXT_DE\tTEXT\t\t1\t1\t\t1");
                sharedParamContent.AppendLine("PARAM\t5bddf4a9-95df-4462-9605-660e73c98678\tSPEC_UNITS_EN\tTEXT\t\t1\t1\t\t1");
                sharedParamContent.AppendLine("PARAM\ta33c632c-37d6-4c7e-92e7-23e0759bd769\tSPEC_UNITS_DE\tTEXT\t\t1\t1\t\t1");
                sharedParamContent.AppendLine("PARAM\t86634033-4673-4cfa-8547-f91fafeecae3\tSPEC_NAME_EN\tTEXT\t\t1\t1\t\t1");
                sharedParamContent.AppendLine("PARAM\t5931b45e-9d32-4af5-b873-584bb185318b\tSPEC_NAME_DE\tTEXT\t\t1\t1\t\t1");
                sharedParamContent.AppendLine("PARAM\tb97019fc-51bb-4ea9-b393-a25e2873aca4\tSPEC_CATEGORY_EN\tTEXT\t\t1\t1\t\t1");
                sharedParamContent.AppendLine("PARAM\t32c12082-d033-473b-a1a0-5ee242fb0662\tSPEC_CATEGORY_DE\tTEXT\t\t1\t1\t\t1");
                sharedParamContent.AppendLine("PARAM\tc2726284-4b65-485a-bb03-29d283d1f2ff\tSPEC_NAME_SHORT_EN\tTEXT\t\t1\t1\t\t1");
                sharedParamContent.AppendLine("PARAM\t87b9f4b9-99b3-4bca-8766-4fb92690a3dd\tSPEC_NAME_SHORT_DE\tTEXT\t\t1\t1\t\t1");
                sharedParamContent.AppendLine("PARAM\t9c5d5890-143e-4904-837b-8ebc8cd7002d\tSPEC_COMMENTS_DE\tTEXT\t\t1\t1\t\t1");
                sharedParamContent.AppendLine("PARAM\tc2b4331c-0bd2-47d2-a67f-bd30de89b52b\tSPEC_COMMENTS_EN\tTEXT\t\t1\t1\t\t1");
                sharedParamContent.AppendLine("PARAM\t146d63a1-b982-4fc9-987c-c94c58fd3251\tSPEC_MANUFACTURER_EN\tTEXT\t\t1\t1\t\t1");
                sharedParamContent.AppendLine("PARAM\tcf30fba4-9841-4298-acfc-1e693fb2b1bf\tSPEC_MANUFACTURER_DE\tTEXT\t\t1\t1\t\t1");
                sharedParamContent.AppendLine("PARAM\t683ef5a1-3265-449e-9f96-4090af1a59af\tSPEC_TYPE_EN\tTEXT\t\t1\t1\t\t1");
                sharedParamContent.AppendLine("PARAM\t110d5cb6-52e0-4c5b-874c-2a8542bc7735\tSPEC_TYPE_DE\tTEXT\t\t1\t1\t\t1");
                sharedParamContent.AppendLine("PARAM\tbae8ecc0-ea26-4bab-9f5a-7ff802f31845\tSPEC_ARTICLE_EN\tTEXT\t\t1\t1\t\t1");
                sharedParamContent.AppendLine("PARAM\t618dd8e2-f854-4001-a929-d3a94f77b168\tSPEC_ARTICLE_DE\tTEXT\t\t1\t1\t\t1");
                sharedParamContent.AppendLine("PARAM\t60a336d6-22e3-4fb4-8dc2-08063d6703a9\tSPEC_QUANTITY\tNUMBER\t\t1\t1\t\t1");
                sharedParamContent.AppendLine("PARAM\t40a95cfa-3f30-40c6-846f-3a51645215e4\tSPEC_SYSTEM\tTEXT\t\t1\t1\t\t1");
                sharedParamContent.AppendLine("PARAM\t7073c9b9-59e8-44a9-8c1d-3d4eee26e683\tSPEC_POSITION\tTEXT\t\t1\t1\t\t1");
                sharedParamContent.AppendLine("PARAM\t63379d40-c680-4c2e-9cad-cbbccbabcf22\tCMMN_STATUS_CODE\tTEXT\t\t1\t1\t\t1");
                sharedParamContent.AppendLine("PARAM\t6a9affe3-aee8-4ccb-9a7c-dbe3d400e8f7\tCMMN_TOOL_ID\tTEXT\t\t1\t1\t\t1");
                sharedParamContent.AppendLine("PARAM\tc7e4d13f-c08f-4713-98e6-2878eb1955a7\tSPEC_SUPPLIER\tTEXT\t\t1\t1\t\t1");
                sharedParamContent.AppendLine("PARAM\t6963799c-6e38-46e6-8529-da230c86ae38\tSPEC_WP\tTEXT\t\t1\t1\t\t1");
                sharedParamContent.AppendLine("PARAM\td2ad2046-40cc-43fc-b1ba-ff6406ee781c\tSPEC_FILTER\tTEXT\t\t1\t1\t\t1");

                // Write the temporary shared parameter file
                File.WriteAllText(tempSharedParamFile, sharedParamContent.ToString());

                // Temporarily set the shared parameter file path
                doc.Application.SharedParametersFilename = tempSharedParamFile;

                // Open the temporary shared parameter file
                DefinitionFile defFile = doc.Application.OpenSharedParameterFile();
                if (defFile == null)
                {
                    stats.Errors.Add("Failed to open temporary shared parameter file");
                    return;
                }

                var group = defFile.Groups.get_Item("SPEC_Parameters");
                if (group == null)
                {
                    stats.Errors.Add("SPEC_Parameters group not found in shared parameter file");
                    return;
                }

                // Define parameter mapping with type information
                var parameterDefinitions = new Dictionary<string, (ForgeTypeId type, bool isInstance, BuiltInParameterGroup group)>
                {
                    [SpecParams.SYSTEM] = (SpecTypeId.String.Text, true, BuiltInParameterGroup.PG_TEXT),
                    [SpecParams.SIZE] = (SpecTypeId.String.Text, true, BuiltInParameterGroup.PG_TEXT),
                    [SpecParams.QUANTITY] = (SpecTypeId.Number, true, BuiltInParameterGroup.PG_TEXT),
                    [SpecParams.POSITION] = (SpecTypeId.String.Text, true, BuiltInParameterGroup.PG_TEXT),
                    [SpecParams.FILTER] = (SpecTypeId.String.Text, true, BuiltInParameterGroup.PG_TEXT),
                    [SpecParams.SUPPLIER] = (SpecTypeId.String.Text, true, BuiltInParameterGroup.PG_TEXT),
                    [SpecParams.WP] = (SpecTypeId.String.Text, true, BuiltInParameterGroup.PG_TEXT),
                    [SpecParams.MATERIAL] = (SpecTypeId.Reference.Material, true, BuiltInParameterGroup.PG_TEXT),
                    [SpecParams.MATERIAL_TEXT_EN] = (SpecTypeId.String.Text, true, BuiltInParameterGroup.PG_TEXT),
                    [SpecParams.MATERIAL_TEXT_DE] = (SpecTypeId.String.Text, true, BuiltInParameterGroup.PG_TEXT),
                    [SpecParams.UNITS_EN] = (SpecTypeId.String.Text, true, BuiltInParameterGroup.PG_TEXT),
                    [SpecParams.UNITS_DE] = (SpecTypeId.String.Text, true, BuiltInParameterGroup.PG_TEXT),
                    [SpecParams.NAME_EN] = (SpecTypeId.String.Text, true, BuiltInParameterGroup.PG_TEXT),
                    [SpecParams.NAME_DE] = (SpecTypeId.String.Text, true, BuiltInParameterGroup.PG_TEXT),
                    [SpecParams.CATEGORY_EN] = (SpecTypeId.String.Text, true, BuiltInParameterGroup.PG_TEXT),
                    [SpecParams.CATEGORY_DE] = (SpecTypeId.String.Text, true, BuiltInParameterGroup.PG_TEXT),
                    [SpecParams.NAME_SHORT_EN] = (SpecTypeId.String.Text, true, BuiltInParameterGroup.PG_TEXT),
                    [SpecParams.NAME_SHORT_DE] = (SpecTypeId.String.Text, true, BuiltInParameterGroup.PG_TEXT),
                    [SpecParams.COMMENTS_EN] = (SpecTypeId.String.Text, true, BuiltInParameterGroup.PG_TEXT),
                    [SpecParams.COMMENTS_DE] = (SpecTypeId.String.Text, true, BuiltInParameterGroup.PG_TEXT),
                    [SpecParams.MANUFACTURER_EN] = (SpecTypeId.String.Text, true, BuiltInParameterGroup.PG_TEXT),
                    [SpecParams.MANUFACTURER_DE] = (SpecTypeId.String.Text, true, BuiltInParameterGroup.PG_TEXT),
                    [SpecParams.TYPE_EN] = (SpecTypeId.String.Text, true, BuiltInParameterGroup.PG_TEXT),
                    [SpecParams.TYPE_DE] = (SpecTypeId.String.Text, true, BuiltInParameterGroup.PG_TEXT),
                    [SpecParams.ARTICLE_EN] = (SpecTypeId.String.Text, true, BuiltInParameterGroup.PG_TEXT),
                    [SpecParams.ARTICLE_DE] = (SpecTypeId.String.Text, true, BuiltInParameterGroup.PG_TEXT),
                    [SpecParams.STATUS_CODE] = (SpecTypeId.String.Text, true, BuiltInParameterGroup.PG_CONSTRAINTS),
                    [SpecParams.TOOL_ID] = (SpecTypeId.String.Text, true, BuiltInParameterGroup.PG_CONSTRAINTS)
                };

                foreach (var kvp in parameterDefinitions)
                {
                    string paramName = kvp.Key;
                    var (typeId, isInstance, paramGroup) = kvp.Value;

                    try
                    {
                        var existing = familyManager.get_Parameter(paramName);

                        // SPEC_SYSTEM and SPEC_POSITION must always be instance
                        bool forceInstance = paramName == SpecParams.SYSTEM || paramName == SpecParams.POSITION;
                        if (existing != null)
                        {
                            if (forceInstance && !existing.IsInstance)
                            {
                                // Remove type parameter and re-add as instance
                                familyManager.RemoveParameter(existing);
                                existing = null; // will create below
                            }
                            else
                            {
                                continue; // already exists correctly
                            }
                        }

                        // Get the definition from the shared parameter file (with the specific GUID)
                        var def = group.Definitions.get_Item(paramName) as ExternalDefinition;

                        if (def != null)
                        {
                            familyManager.AddParameter(def, paramGroup, forceInstance ? true : isInstance);
                            stats.ParametersCreated++;
                            stats.CreatedParameters.Add(paramName);
                        }
                        else
                        {
                            stats.Warnings.Add($"Parameter {paramName} definition not found in shared parameter file.");
                        }
                    }
                    catch (Exception ex)
                    {
                        stats.Errors.Add($"Failed to create parameter {paramName}: {ex.Message}");
                    }
                }
            }
            finally
            {
                // CRITICAL: Restore original shared parameter file
                if (!string.IsNullOrEmpty(originalSharedParamFile))
                {
                    doc.Application.SharedParametersFilename = originalSharedParamFile;
                }
                else
                {
                    doc.Application.SharedParametersFilename = "";
                }

                // Clean up temporary file
                if (!string.IsNullOrEmpty(tempSharedParamFile))
                {
                    try
                    {
                        if (File.Exists(tempSharedParamFile))
                        {
                            File.Delete(tempSharedParamFile);
                        }
                    }
                    catch
                    {
                        // Ignore cleanup errors
                    }
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