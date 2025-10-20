using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace SpecParametersUpdater
{
    public class ParameterCache
    {
        private readonly Document _doc;
        private readonly Dictionary<ElementId, Dictionary<string, Parameter>> _typeCache;

        public ParameterCache(Document doc)
        {
            _doc = doc;
            _typeCache = new Dictionary<ElementId, Dictionary<string, Parameter>>();
        }

        public Parameter GetTypeParameter(Element element, string paramName)
        {
            try
            {
                var typeId = element.GetTypeId();
                if (typeId == ElementId.InvalidElementId) return null;

                if (!_typeCache.TryGetValue(typeId, out var dict))
                {
                    dict = new Dictionary<string, Parameter>(StringComparer.OrdinalIgnoreCase);
                    var typeElem = _doc.GetElement(typeId);
                    if (typeElem != null)
                    {
                        foreach (Parameter p in typeElem.Parameters)
                        {
                            if (p?.Definition?.Name != null && !dict.ContainsKey(p.Definition.Name))
                                dict[p.Definition.Name] = p;
                        }
                    }
                    _typeCache[typeId] = dict;
                }

                dict.TryGetValue(paramName, out var result);
                return result;
            }
            catch { return null; }
        }

        public Parameter GetInstanceOrType(Element element, string paramName)
        {
            try
            {
                var p = element.LookupParameter(paramName);
                return p ?? GetTypeParameter(element, paramName);
            }
            catch { return null; }
        }
    }
}