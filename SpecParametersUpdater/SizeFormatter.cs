using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;

namespace SpecParametersUpdater
{
    public static class SizeFormatter
    {
        private static readonly Dictionary<double, string> InchSizes = new Dictionary<double, string>
        {
            { 6.35, "1/4\"" }, { 9.53, "3/8\"" }, { 12.7, "1/2\"" }, { 19.05, "3/4\"" },
            { 25.4, "1\"" }, { 31.75, "1-1/4\"" }, { 38.1, "1-1/2\"" }, { 50.8, "2\"" },
            { 63.5, "2-1/2\"" }, { 76.2, "3\"" }, { 88.9, "3-1/2\"" }, { 101.6, "4\"" },
            { 127, "5\"" }, { 152.4, "6\"" }, { 203.2, "8\"" }, { 254, "10\"" }, { 304.8, "12\"" }
        };

        private static readonly char[] Separators = new[] { 'x', '×', '-', '*' };

        public static string Format(string rawSize, Element element)
        {
            if (string.IsNullOrWhiteSpace(rawSize)) return rawSize?.Trim() ?? string.Empty;

            string clean = rawSize.Trim();
            clean = Regex.Replace(clean, @"\s+", " ");
            clean = Regex.Replace(clean, @"DN|NW", "", RegexOptions.IgnoreCase).Trim();

            bool isPipe = IsPipe(element, clean);
            return isPipe ? FormatPipe(clean) : FormatDuct(clean);
        }

        private static bool IsPipe(Element element, string cleaned)
        {
            if (element.Category != null)
            {
                string catName = element.Category.Name;
                if (catName.IndexOf("Pipe", StringComparison.InvariantCultureIgnoreCase) >= 0 ||
                    catName.IndexOf("Rohr", StringComparison.InvariantCultureIgnoreCase) >= 0)
                    return true;
            }

            return cleaned.Contains("\"") || cleaned.Contains("mm") ||
                   Regex.IsMatch(cleaned, @"\b\d+\s*/\s*\d+\b");
        }

        private static string FormatPipe(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return value;

            string clean = Regex.Replace(value, @"mm|in", "", RegexOptions.IgnoreCase);
            clean = clean.Replace("\"", "").Trim();

            var fraction = Regex.Match(clean, @"\b(\d+)\s*/\s*(\d+)\b");
            if (fraction.Success)
                return $"{fraction.Groups[1].Value}/{fraction.Groups[2].Value}\"";

            if (clean.IndexOfAny(Separators) >= 0)
            {
                var parts = Regex.Split(clean, @"[×x\-*]")
                    .Select(p => p.Trim())
                    .Where(p => p.Length > 0)
                    .ToArray();

                var dims = new List<string>();
                foreach (var part in parts)
                {
                    if (TryParse(part, out double val))
                        dims.Add(FormatPipeDimension(val));
                    else
                        dims.Add(part);
                }
                return string.Join("x", dims);
            }

            if (TryParse(clean, out double single))
                return FormatPipeDimension(single);

            return value;
        }

        private static string FormatPipeDimension(double mm)
        {
            foreach (var kvp in InchSizes)
            {
                if (Math.Abs(mm - kvp.Key) < Config.InchTolerance)
                    return kvp.Value;
            }

            if (mm < 6)
                return mm.ToString("F1", CultureInfo.InvariantCulture).TrimEnd('0').TrimEnd('.');

            return "DN" + Math.Round(mm).ToString(CultureInfo.InvariantCulture);
        }

        private static string FormatDuct(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return value;

            string clean = Regex.Replace(value, @"DN|NW", "", RegexOptions.IgnoreCase).Trim();

            if (clean.IndexOfAny(Separators) >= 0)
            {
                var parts = Regex.Split(clean, @"[×x\-*]")
                    .Select(p => p.Trim())
                    .Where(p => p.Length > 0)
                    .ToArray();

                var dims = new List<string>();
                foreach (var part in parts)
                {
                    if (TryParse(part, out double val))
                        dims.Add(((int)Math.Round(val)).ToString(CultureInfo.InvariantCulture));
                    else
                        dims.Add(part);
                }

                if (dims.Count == 2)
                {
                    if (TryParse(parts[0], out double d1) && TryParse(parts[1], out double d2))
                    {
                        if (d1 < d2) dims.Reverse();
                    }
                }
                else if (dims.Count == 3)
                {
                    if (TryParse(parts[0], out double d1) && TryParse(parts[1], out double d2))
                    {
                        if (Math.Abs(d1 - d2) < 1.0) dims.RemoveAt(1);
                    }
                }

                return "DN" + string.Join("x", dims);
            }

            if (TryParse(clean, out double num))
                return "DN" + ((int)Math.Round(num)).ToString(CultureInfo.InvariantCulture);

            return "DN" + clean;
        }

        private static bool TryParse(string input, out double value)
        {
            value = 0;
            if (string.IsNullOrWhiteSpace(input)) return false;

            string clean = input.Trim();
            clean = Regex.Replace(clean, @"[^\d\-,\./]", "");

            if (clean.Contains("/"))
            {
                var parts = clean.Split('/');
                if (parts.Length == 2 &&
                    double.TryParse(parts[0], NumberStyles.Any, CultureInfo.InvariantCulture, out double num) &&
                    double.TryParse(parts[1], NumberStyles.Any, CultureInfo.InvariantCulture, out double den) &&
                    Math.Abs(den) > double.Epsilon)
                {
                    value = num / den;
                    return true;
                }
                return false;
            }

            clean = clean.Replace(",", ".");

            if (double.TryParse(clean, NumberStyles.Any, CultureInfo.InvariantCulture, out double result))
            {
                value = result;
                return true;
            }

            return false;
        }
    }
}