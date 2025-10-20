using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Data;
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
                var categories = Categories.GetAll();
                var selection = uidoc.Selection?.GetElementIds();
                bool useSelection = selection != null && selection.Count > 0;

                var elementsToProcess = CollectElements(doc, selection, useSelection, categories);
                var cache = new ParameterCache(doc);
                var stats = new UpdateStats();

                if (Config.ValidateParameters)
                {
                    ValidateParameters(doc, stats.Warnings);
                }

                using (var trans = new Transaction(doc, "Update SPEC_SYSTEM, SPEC_SIZE, SPEC_QUANTITY"))
                {
                    trans.Start();
                    UpdateAllElements(elementsToProcess, cache, stats, categories);
                    trans.Commit();
                }

                stopwatch.Stop();
                ShowResults(doc, stats, stopwatch, useSelection);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error: {ex.Message}";
                return Result.Failed;
            }
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

            var name = element.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM);
            if (ParameterHelper.TryGetString(name, out val)) return val.Trim();

            var classif = element.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM);
            if (ParameterHelper.TryGetString(classif, out val)) return val.Trim();

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

                var lengthParam = element.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)
                    ?? cache.GetInstanceOrType(element, "Length");
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

        private void ValidateParameters(Document doc, List<string> warnings)
        {
            var testElement = new FilteredElementCollector(doc).WhereElementIsNotElementType().FirstOrDefault();
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

        private void ShowResults(Document doc, UpdateStats stats, Stopwatch stopwatch, bool useSelection)
        {
            var summary = BuildSummary(doc, stats, stopwatch, useSelection);

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
            dialog.MainContent =
                $"Mode: {(useSelection ? "SELECTION" : "ALL CATEGORIES")}\n\n" +
                $"SPEC_SYSTEM: {stats.SystemUpdates}\n" +
                $"SPEC_SIZE: {stats.SizeUpdates}\n" +
                $"SPEC_QUANTITY: {stats.QuantityUpdates}\n\n" +
                $"Total: {stats.TotalUpdates}\n" +
                $"Errors: {stats.Errors.Count}\n" +
                $"Time: {stopwatch.Elapsed.TotalSeconds:F2}s";

            if (!string.IsNullOrEmpty(logPath))
                dialog.ExpandedContent = $"Log: {logPath}";

            dialog.Show();
        }

        private string BuildSummary(Document doc, UpdateStats stats, Stopwatch stopwatch, bool useSelection)
        {
            var sb = new StringBuilder();
            sb.AppendLine("SPEC PARAMETERS UPDATE");
            sb.AppendLine($"Document: {doc.Title}");
            sb.AppendLine($"Mode: {(useSelection ? "SELECTION" : "ALL")}");
            sb.AppendLine($"Time: {stopwatch.Elapsed.TotalSeconds:F2}s");
            sb.AppendLine();
            sb.AppendLine($"SPEC_SYSTEM: {stats.SystemUpdates}");
            sb.AppendLine($"SPEC_SIZE: {stats.SizeUpdates}");
            sb.AppendLine($"SPEC_QUANTITY: {stats.QuantityUpdates}");
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