using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace SpecParametersUpdater
{
    public static class Config
    {
        public const bool WriteLogFile = true;
        public const bool ValidateParameters = true;
        public const double DoubleTolerance = 1e-6;
        public const double InchTolerance = 0.6;
    }

    public static class SpecParams
    {
        public const string SYSTEM = "SPEC_SYSTEM";
        public const string SIZE = "SPEC_SIZE";
        public const string QUANTITY = "SPEC_QUANTITY";
        public const string POSITION = "SPEC_POSITION";
        public const string FILTER = "SPEC_FILTER";
        public const string SUPPLIER = "SPEC_SUPPLIER";
        public const string WP = "SPEC_WP";
        public const string MATERIAL = "SPEC_MATERIAL";
        public const string MATERIAL_TEXT_EN = "SPEC_MATERIAL_TEXT_EN";
        public const string MATERIAL_TEXT_DE = "SPEC_MATERIAL_TEXT_DE";
        public const string UNITS_EN = "SPEC_UNITS_EN";
        public const string UNITS_DE = "SPEC_UNITS_DE";
        public const string NAME_EN = "SPEC_NAME_EN";
        public const string NAME_DE = "SPEC_NAME_DE";
        public const string CATEGORY_EN = "SPEC_CATEGORY_EN";
        public const string CATEGORY_DE = "SPEC_CATEGORY_DE";
        public const string NAME_SHORT_EN = "SPEC_NAME_SHORT_EN";
        public const string NAME_SHORT_DE = "SPEC_NAME_SHORT_DE";
        public const string COMMENTS_EN = "SPEC_COMMENTS_EN";
        public const string COMMENTS_DE = "SPEC_COMMENTS_DE";
        public const string MANUFACTURER_EN = "SPEC_MANUFACTURER_EN";
        public const string MANUFACTURER_DE = "SPEC_MANUFACTURER_DE";
        public const string TYPE_EN = "SPEC_TYPE_EN";
        public const string TYPE_DE = "SPEC_TYPE_DE";
        public const string ARTICLE_EN = "SPEC_ARTICLE_EN";
        public const string ARTICLE_DE = "SPEC_ARTICLE_DE";
        public const string STATUS_CODE = "CMMN_STATUS_CODE";
        public const string TOOL_ID = "CMMN_TOOL_ID";

        public static string[] GetAll()
        {
            return new[]
            {
                SYSTEM, SIZE, QUANTITY, POSITION, FILTER, SUPPLIER, WP, MATERIAL,
                MATERIAL_TEXT_EN, MATERIAL_TEXT_DE, UNITS_EN, UNITS_DE,
                NAME_EN, NAME_DE, CATEGORY_EN, CATEGORY_DE,
                NAME_SHORT_EN, NAME_SHORT_DE, COMMENTS_EN, COMMENTS_DE,
                MANUFACTURER_EN, MANUFACTURER_DE, TYPE_EN, TYPE_DE,
                ARTICLE_EN, ARTICLE_DE, STATUS_CODE, TOOL_ID
            };
        }
    }

    public static class Categories
    {
        public static BuiltInCategory[] GetAll()
        {
            return new[]
            {
                BuiltInCategory.OST_PipeCurves, BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_PipeAccessory, BuiltInCategory.OST_PipeInsulations,
                BuiltInCategory.OST_FlexPipeCurves, BuiltInCategory.OST_PlumbingFixtures,
                BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_DuctFitting,
                BuiltInCategory.OST_DuctAccessory, BuiltInCategory.OST_DuctInsulations,
                BuiltInCategory.OST_FlexDuctCurves, BuiltInCategory.OST_MechanicalEquipment,
                BuiltInCategory.OST_CableTray, BuiltInCategory.OST_CableTrayFitting,
                BuiltInCategory.OST_Conduit, BuiltInCategory.OST_ConduitFitting,
                BuiltInCategory.OST_ElectricalEquipment, BuiltInCategory.OST_ElectricalFixtures,
                BuiltInCategory.OST_LightingFixtures
            };
        }

        public static HashSet<BuiltInCategory> GetSizeQuantityCategories()
        {
            return new HashSet<BuiltInCategory>(new[]
            {
                BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_DuctInsulations,
                BuiltInCategory.OST_DuctFitting, BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_CableTrayFitting, BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_ConduitFitting, BuiltInCategory.OST_FlexDuctCurves,
                BuiltInCategory.OST_PipeCurves, BuiltInCategory.OST_PipeInsulations,
                BuiltInCategory.OST_FlexPipeCurves, BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_ElectricalEquipment
            });
        }
    }

    public class UpdateStats
    {
        public int SystemUpdates { get; set; }
        public int SizeUpdates { get; set; }
        public int QuantityUpdates { get; set; }
        public List<string> Errors { get; } = new List<string>();
        public List<string> Warnings { get; } = new List<string>();
        public int TotalUpdates => SystemUpdates + SizeUpdates + QuantityUpdates;
    }
}