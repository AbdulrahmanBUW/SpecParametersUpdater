using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
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

            if (!doc.IsFamilyDocument)
            {
                TaskDialog.Show("Error", "This tool must be run in a Family document");
                return Result.Cancelled;
            }

            try
            {
                using (var trans = new Transaction(doc, "Add/Update SPEC Type and Instance Parameters"))
                {
                    trans.Start();
                    AddSpecParameters(doc);
                    trans.Commit();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error: {ex.Message}";
                return Result.Failed;
            }
        }

        private void AddSpecParameters(Document doc)
        {
            var familyManager = doc.FamilyManager;
            if (familyManager == null)
            {
                TaskDialog.Show("Error", "FamilyManager not available");
                return;
            }

            // Store original shared parameter file
            string originalSharedParamFile = doc.Application.SharedParametersFilename;
            string tempFilePath = null;

            try
            {
                // Create temporary shared parameter file
                tempFilePath = CreateSharedParameterFile();

                // Set the temporary file as shared parameter file
                doc.Application.SharedParametersFilename = tempFilePath;

                // Open the shared parameter file
                DefinitionFile sharedParamFile = doc.Application.OpenSharedParameterFile();

                if (sharedParamFile == null)
                {
                    TaskDialog.Show("Error", "Failed to open shared parameter file");
                    return;
                }

                int addedTypeCount = 0;
                int addedInstanceCount = 0;
                int convertedCount = 0;
                int failedCount = 0;
                int skippedCount = 0;

                var logMessages = new List<string>();
                logMessages.Add("Adding SPEC parameters...");
                logMessages.Add(new string('=', 60));

                // Get the parameter group
                DefinitionGroup paramGroup = sharedParamFile.Groups.get_Item("SPEC_Parameters");

                if (paramGroup == null)
                {
                    TaskDialog.Show("Error", "SPEC_Parameters group not found in shared parameter file");
                    return;
                }

                // Define which parameters should be instance parameters
                var instanceParameters = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "SPEC_FILTER",
                    "CMMN_TOOL_ID",
                    "CMMN_STATUS_CODE",
                    "SPEC_SUPPLIER",
                    "SPEC_WP",
                    "SPEC_SYSTEM",
                    "SPEC_POSITION"
                };

                // Get existing parameters
                var existingParams = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (FamilyParameter fp in familyManager.Parameters)
                {
                    if (fp?.Definition?.Name != null)
                    {
                        existingParams.Add(fp.Definition.Name);
                    }
                }

                // All SPEC parameters in order
                var specParameters = new[]
                {
                    "SPEC_SIZE",
                    "SPEC_MATERIAL",
                    "SPEC_MATERIAL_TEXT_EN",
                    "SPEC_MATERIAL_TEXT_DE",
                    "SPEC_UNITS_EN",
                    "SPEC_UNITS_DE",
                    "SPEC_NAME_EN",
                    "SPEC_NAME_DE",
                    "SPEC_CATEGORY_EN",
                    "SPEC_CATEGORY_DE",
                    "SPEC_NAME_SHORT_EN",
                    "SPEC_NAME_SHORT_DE",
                    "SPEC_COMMENTS_DE",
                    "SPEC_COMMENTS_EN",
                    "SPEC_MANUFACTURER_EN",
                    "SPEC_MANUFACTURER_DE",
                    "SPEC_TYPE_EN",
                    "SPEC_TYPE_DE",
                    "SPEC_ARTICLE_EN",
                    "SPEC_ARTICLE_DE",
                    "SPEC_QUANTITY",
                    "SPEC_SYSTEM",
                    "SPEC_POSITION",
                    "CMMN_STATUS_CODE",
                    "CMMN_TOOL_ID",
                    "SPEC_SUPPLIER",
                    "SPEC_WP",
                    "SPEC_FILTER"
                };

                // Process each SPEC parameter
                foreach (var paramName in specParameters)
                {
                    // Determine if this should be an instance parameter
                    bool isInstance = instanceParameters.Contains(paramName);
                    string paramType = isInstance ? "Instance" : "Type";

                    logMessages.Add($"Processing: {paramName} ({paramType} parameter)");

                    // Check if parameter already exists
                    if (existingParams.Contains(paramName))
                    {
                        // Special handling for parameters that need conversion
                        if (paramName == "SPEC_SYSTEM" || paramName == "SPEC_POSITION")
                        {
                            FamilyParameter existingParam = familyManager.get_Parameter(paramName);

                            if (existingParam != null && !existingParam.IsInstance)
                            {
                                logMessages.Add($"  ⚠ {paramName} exists as Type parameter - converting to Instance");
                                try
                                {
                                    // Remove the existing type parameter
                                    familyManager.RemoveParameter(existingParam);

                                    // Find the parameter definition in the shared parameter file
                                    ExternalDefinition paramDef = paramGroup.Definitions.get_Item(paramName) as ExternalDefinition;

                                    if (paramDef != null)
                                    {
                                        // Re-add as instance parameter
                                        FamilyParameter familyParam = familyManager.AddParameter(
                                            paramDef,
                                            BuiltInParameterGroup.PG_TEXT,
                                            true);

                                        if (familyParam != null)
                                        {
                                            logMessages.Add("  ✓ Successfully converted to Instance");
                                            convertedCount++;
                                        }
                                        else
                                        {
                                            logMessages.Add("  ✗ Failed to convert to Instance");
                                            failedCount++;
                                        }
                                    }
                                    else
                                    {
                                        logMessages.Add("  ✗ Parameter definition not found for conversion");
                                        failedCount++;
                                    }
                                }
                                catch (Exception e)
                                {
                                    logMessages.Add($"  ✗ Error converting: {e.Message}");
                                    failedCount++;
                                }
                            }
                            else
                            {
                                logMessages.Add("  ⚠ Already exists - skipping");
                                skippedCount++;
                            }
                        }
                        else
                        {
                            logMessages.Add("  ⚠ Already exists - skipping");
                            skippedCount++;
                        }
                        continue;
                    }

                    try
                    {
                        // Find the parameter definition in the shared parameter file
                        ExternalDefinition paramDef = paramGroup.Definitions.get_Item(paramName) as ExternalDefinition;

                        if (paramDef != null)
                        {
                            // Add parameter (true for instance, false for type)
                            FamilyParameter familyParam = familyManager.AddParameter(
                                paramDef,
                                BuiltInParameterGroup.PG_TEXT,
                                isInstance);

                            if (familyParam != null)
                            {
                                logMessages.Add($"  ✓ Successfully added as {paramType}");
                                if (isInstance)
                                    addedInstanceCount++;
                                else
                                    addedTypeCount++;
                            }
                            else
                            {
                                logMessages.Add("  ✗ Failed to add to family");
                                failedCount++;
                            }
                        }
                        else
                        {
                            logMessages.Add("  ✗ Parameter definition not found");
                            failedCount++;
                        }
                    }
                    catch (Exception e)
                    {
                        logMessages.Add($"  ✗ Error: {e.Message}");
                        failedCount++;
                    }
                }

                // Show results
                logMessages.Add(new string('=', 60));
                logMessages.Add("OPERATION COMPLETE");
                logMessages.Add(new string('=', 60));

                var resultMsg = new StringBuilder();
                resultMsg.AppendLine("SPEC Parameters Results:\n");
                resultMsg.AppendLine($"✓ Added Type parameters: {addedTypeCount}");
                resultMsg.AppendLine($"✓ Added Instance parameters: {addedInstanceCount}");
                if (convertedCount > 0)
                    resultMsg.AppendLine($"✓ Converted to Instance: {convertedCount}");
                resultMsg.AppendLine($"⚠ Skipped (already exist): {skippedCount}");
                resultMsg.AppendLine($"✗ Failed: {failedCount}\n");
                resultMsg.AppendLine("Instance Parameters:");
                foreach (var param in instanceParameters)
                {
                    resultMsg.AppendLine($"  • {param}");
                }
                resultMsg.AppendLine("\n* SPEC_SYSTEM and SPEC_POSITION will be converted to Instance if they exist as Type");
                resultMsg.AppendLine("\nAll other parameters added as Type parameters");

                // Show dialog
                var dialog = new TaskDialog("SPEC Parameters Complete");
                dialog.MainInstruction = "SPEC Parameters Complete";
                dialog.MainContent = resultMsg.ToString();
                dialog.ExpandedContent = string.Join("\n", logMessages);
                dialog.Show();
            }
            finally
            {
                // Restore original shared parameter file
                if (!string.IsNullOrEmpty(originalSharedParamFile))
                {
                    doc.Application.SharedParametersFilename = originalSharedParamFile;
                }
                else
                {
                    doc.Application.SharedParametersFilename = "";
                }

                // Clean up temporary file
                if (!string.IsNullOrEmpty(tempFilePath))
                {
                    try
                    {
                        if (File.Exists(tempFilePath))
                            File.Delete(tempFilePath);
                    }
                    catch { }
                }
            }
        }

        private string CreateSharedParameterFile()
        {
            string tempFile = Path.Combine(Path.GetTempPath(), $"SpecParams_{Guid.NewGuid()}.txt");

            var content = new StringBuilder();
            content.AppendLine("# This is a Revit shared parameter file.");
            content.AppendLine("# Do not edit manually.");
            content.AppendLine("*META\tVERSION\tMINVERSION");
            content.AppendLine("META\t2\t1");
            content.AppendLine("*GROUP\tID\tNAME");
            content.AppendLine("GROUP\t1\tSPEC_Parameters");
            content.AppendLine("*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\tVISIBLE\tDESCRIPTION\tUSERMODIFIABLE");

            // Add parameters with exact GUIDs
            content.AppendLine("PARAM\t9f2f7b1a-517f-48a1-8c2b-bfec1948dc74\tSPEC_SIZE\tTEXT\tTEXT\t1\t1\tSPEC_SIZE\t1");
            content.AppendLine("PARAM\t25d801a6-2723-4ed1-8e69-bab44cab7e2f\tSPEC_MATERIAL\tMATERIAL\tMATERIAL\t1\t1\tSPEC_MATERIAL\t1");
            content.AppendLine("PARAM\tec43941f-42ec-4fd6-9826-033c28505638\tSPEC_MATERIAL_TEXT_EN\tTEXT\tTEXT\t1\t1\tSPEC_MATERIAL_TEXT_EN\t1");
            content.AppendLine("PARAM\tb97c365e-4ad7-4ae2-903b-e6d7d1303d31\tSPEC_MATERIAL_TEXT_DE\tTEXT\tTEXT\t1\t1\tSPEC_MATERIAL_TEXT_DE\t1");
            content.AppendLine("PARAM\t5bddf4a9-95df-4462-9605-660e73c98678\tSPEC_UNITS_EN\tTEXT\tTEXT\t1\t1\tSPEC_UNITS_EN\t1");
            content.AppendLine("PARAM\ta33c632c-37d6-4c7e-92e7-23e0759bd769\tSPEC_UNITS_DE\tTEXT\tTEXT\t1\t1\tSPEC_UNITS_DE\t1");
            content.AppendLine("PARAM\t86634033-4673-4cfa-8547-f91fafeecae3\tSPEC_NAME_EN\tTEXT\tTEXT\t1\t1\tSPEC_NAME_EN\t1");
            content.AppendLine("PARAM\t5931b45e-9d32-4af5-b873-584bb185318b\tSPEC_NAME_DE\tTEXT\tTEXT\t1\t1\tSPEC_NAME_DE\t1");
            content.AppendLine("PARAM\tb97019fc-51bb-4ea9-b393-a25e2873aca4\tSPEC_CATEGORY_EN\tTEXT\tTEXT\t1\t1\tSPEC_CATEGORY_EN\t1");
            content.AppendLine("PARAM\t32c12082-d033-473b-a1a0-5ee242fb0662\tSPEC_CATEGORY_DE\tTEXT\tTEXT\t1\t1\tSPEC_CATEGORY_DE\t1");
            content.AppendLine("PARAM\tc2726284-4b65-485a-bb03-29d283d1f2ff\tSPEC_NAME_SHORT_EN\tTEXT\tTEXT\t1\t1\tSPEC_NAME_SHORT_EN\t1");
            content.AppendLine("PARAM\t87b9f4b9-99b3-4bca-8766-4fb92690a3dd\tSPEC_NAME_SHORT_DE\tTEXT\tTEXT\t1\t1\tSPEC_NAME_SHORT_DE\t1");
            content.AppendLine("PARAM\t9c5d5890-143e-4904-837b-8ebc8cd7002d\tSPEC_COMMENTS_DE\tTEXT\tTEXT\t1\t1\tSPEC_COMMENTS_DE\t1");
            content.AppendLine("PARAM\tc2b4331c-0bd2-47d2-a67f-bd30de89b52b\tSPEC_COMMENTS_EN\tTEXT\tTEXT\t1\t1\tSPEC_COMMENTS_EN\t1");
            content.AppendLine("PARAM\t146d63a1-b982-4fc9-987c-c94c58fd3251\tSPEC_MANUFACTURER_EN\tTEXT\tTEXT\t1\t1\tSPEC_MANUFACTURER_EN\t1");
            content.AppendLine("PARAM\tcf30fba4-9841-4298-acfc-1e693fb2b1bf\tSPEC_MANUFACTURER_DE\tTEXT\tTEXT\t1\t1\tSPEC_MANUFACTURER_DE\t1");
            content.AppendLine("PARAM\t683ef5a1-3265-449e-9f96-4090af1a59af\tSPEC_TYPE_EN\tTEXT\tTEXT\t1\t1\tSPEC_TYPE_EN\t1");
            content.AppendLine("PARAM\t110d5cb6-52e0-4c5b-874c-2a8542bc7735\tSPEC_TYPE_DE\tTEXT\tTEXT\t1\t1\tSPEC_TYPE_DE\t1");
            content.AppendLine("PARAM\tbae8ecc0-ea26-4bab-9f5a-7ff802f31845\tSPEC_ARTICLE_EN\tTEXT\tTEXT\t1\t1\tSPEC_ARTICLE_EN\t1");
            content.AppendLine("PARAM\t618dd8e2-f854-4001-a929-d3a94f77b168\tSPEC_ARTICLE_DE\tTEXT\tTEXT\t1\t1\tSPEC_ARTICLE_DE\t1");
            content.AppendLine("PARAM\t60a336d6-22e3-4fb4-8dc2-08063d6703a9\tSPEC_QUANTITY\tNUMBER\tNUMBER\t1\t1\tSPEC_QUANTITY\t1");
            content.AppendLine("PARAM\t40a95cfa-3f30-40c6-846f-3a51645215e4\tSPEC_SYSTEM\tTEXT\tTEXT\t1\t1\tSPEC_SYSTEM\t1");
            content.AppendLine("PARAM\t7073c9b9-59e8-44a9-8c1d-3d4eee26e683\tSPEC_POSITION\tTEXT\tTEXT\t1\t1\tSPEC_POSITION\t1");
            content.AppendLine("PARAM\t63379d40-c680-4c2e-9cad-cbbccbabcf22\tCMMN_STATUS_CODE\tTEXT\tTEXT\t1\t1\tCMMN_STATUS_CODE\t1");
            content.AppendLine("PARAM\t6a9affe3-aee8-4ccb-9a7c-dbe3d400e8f7\tCMMN_TOOL_ID\tTEXT\tTEXT\t1\t1\tCMMN_TOOL_ID\t1");
            content.AppendLine("PARAM\tc7e4d13f-c08f-4713-98e6-2878eb1955a7\tSPEC_SUPPLIER\tTEXT\tTEXT\t1\t1\tSPEC_SUPPLIER\t1");
            content.AppendLine("PARAM\t6963799c-6e38-46e6-8529-da230c86ae38\tSPEC_WP\tTEXT\tTEXT\t1\t1\tSPEC_WP\t1");
            content.AppendLine("PARAM\td2ad2046-40cc-43fc-b1ba-ff6406ee781c\tSPEC_FILTER\tTEXT\tTEXT\t1\t1\tSPEC_FILTER\t1");

            File.WriteAllText(tempFile, content.ToString());
            return tempFile;
        }
    }
}