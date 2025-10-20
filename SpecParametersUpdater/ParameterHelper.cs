using System;
using System.Globalization;
using Autodesk.Revit.DB;

namespace SpecParametersUpdater
{
    public static class ParameterHelper
    {
        public static bool TryGetString(Parameter param, out string value)
        {
            value = null;
            if (param == null || !param.HasValue) return false;

            try
            {
                if (param.StorageType == StorageType.String)
                {
                    value = param.AsString();
                    return !string.IsNullOrWhiteSpace(value);
                }

                var vs = param.AsValueString();
                if (!string.IsNullOrWhiteSpace(vs))
                {
                    value = vs;
                    return true;
                }

                if (param.StorageType == StorageType.Double)
                {
                    value = param.AsDouble().ToString(CultureInfo.InvariantCulture);
                    return true;
                }

                if (param.StorageType == StorageType.Integer)
                {
                    value = param.AsInteger().ToString(CultureInfo.InvariantCulture);
                    return true;
                }
            }
            catch { }

            return false;
        }

        public static string GetString(Parameter param)
        {
            if (param == null || !param.HasValue) return null;

            try
            {
                if (param.StorageType == StorageType.String) return param.AsString();
                var vs = param.AsValueString();
                if (!string.IsNullOrWhiteSpace(vs)) return vs;
                if (param.StorageType == StorageType.Double)
                    return param.AsDouble().ToString(CultureInfo.InvariantCulture);
                if (param.StorageType == StorageType.Integer)
                    return param.AsInteger().ToString(CultureInfo.InvariantCulture);
            }
            catch { }

            return null;
        }

        public static double GetDouble(Parameter param)
        {
            if (param == null || !param.HasValue) return 0.0;

            try
            {
                if (param.StorageType == StorageType.Double) return param.AsDouble();
                if (param.StorageType == StorageType.Integer) return param.AsInteger();
                if (param.StorageType == StorageType.String)
                {
                    if (double.TryParse(param.AsString(), NumberStyles.Any,
                        CultureInfo.InvariantCulture, out double result))
                        return result;
                }
            }
            catch { }

            return 0.0;
        }

        public static bool UpdateString(Parameter param, Element element, string newValue, ParameterCache cache)
        {
            if (param == null || string.IsNullOrWhiteSpace(newValue)) return false;

            string current = GetString(param) ?? string.Empty;
            if (current.Trim().Equals(newValue.Trim(), StringComparison.OrdinalIgnoreCase)) return false;

            try
            {
                param.Set(newValue);
                return true;
            }
            catch
            {
                try
                {
                    var typeParam = cache.GetTypeParameter(element, param.Definition.Name);
                    if (typeParam != null && !typeParam.IsReadOnly)
                    {
                        typeParam.Set(newValue);
                        return true;
                    }
                }
                catch { }
            }

            return false;
        }

        public static bool UpdateNumeric(Parameter param, Element element, double newValue, ParameterCache cache)
        {
            if (param == null) return false;

            try
            {
                if (param.StorageType == StorageType.Double)
                {
                    double current = param.AsDouble();
                    if (Math.Abs(current - newValue) > Config.DoubleTolerance)
                    {
                        param.Set(newValue);
                        return true;
                    }
                    return false;
                }
                else if (param.StorageType == StorageType.Integer)
                {
                    int current = param.AsInteger();
                    int newInt = Convert.ToInt32(Math.Round(newValue));
                    if (current != newInt)
                    {
                        param.Set(newInt);
                        return true;
                    }
                    return false;
                }
                else if (param.StorageType == StorageType.String)
                {
                    string current = param.AsString() ?? string.Empty;
                    string newStr = newValue.ToString("F2", CultureInfo.InvariantCulture);
                    if (!current.Trim().Equals(newStr.Trim()))
                    {
                        param.Set(newStr);
                        return true;
                    }
                    return false;
                }
            }
            catch { }

            try
            {
                var typeParam = cache.GetTypeParameter(element, param.Definition.Name);
                if (typeParam != null && !typeParam.IsReadOnly)
                {
                    if (typeParam.StorageType == StorageType.Double)
                    {
                        double current = typeParam.AsDouble();
                        if (Math.Abs(current - newValue) > Config.DoubleTolerance)
                        {
                            typeParam.Set(newValue);
                            return true;
                        }
                    }
                    else if (typeParam.StorageType == StorageType.Integer)
                    {
                        int current = typeParam.AsInteger();
                        int newInt = Convert.ToInt32(Math.Round(newValue));
                        if (current != newInt)
                        {
                            typeParam.Set(newInt);
                            return true;
                        }
                    }
                    else if (typeParam.StorageType == StorageType.String)
                    {
                        string current = typeParam.AsString() ?? string.Empty;
                        string newStr = newValue.ToString("F2", CultureInfo.InvariantCulture);
                        if (!current.Trim().Equals(newStr.Trim()))
                        {
                            typeParam.Set(newStr);
                            return true;
                        }
                    }
                }
            }
            catch { }

            return false;
        }
    }
}