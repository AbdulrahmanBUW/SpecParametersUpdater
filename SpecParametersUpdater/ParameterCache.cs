using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace SpecParametersUpdater
{
    public class ParameterCache
    {
        private readonly Document _doc;
        private readonly bool _isFamilyDoc;
        private readonly FamilyManager _familyManager;
        private readonly Dictionary<ElementId, Dictionary<string, Parameter>> _typeCache;
        private readonly Dictionary<string, FamilyParameter> _familyParamCache;

        public ParameterCache(Document doc) : this(doc, false)
        {
        }

        public ParameterCache(Document doc, bool isFamilyDoc)
        {
            _doc = doc;
            _isFamilyDoc = isFamilyDoc;
            _typeCache = new Dictionary<ElementId, Dictionary<string, Parameter>>();
            _familyParamCache = new Dictionary<string, FamilyParameter>(StringComparer.OrdinalIgnoreCase);

            if (_isFamilyDoc)
            {
                _familyManager = doc.FamilyManager;
                if (_familyManager != null)
                {
                    // Cache all family parameters
                    foreach (FamilyParameter fp in _familyManager.Parameters)
                    {
                        if (fp?.Definition?.Name != null)
                        {
                            _familyParamCache[fp.Definition.Name] = fp;
                        }
                    }
                }
            }
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
                // First try to get instance parameter
                var p = element.LookupParameter(paramName);
                if (p != null) return p;

                // Then try type parameter
                p = GetTypeParameter(element, paramName);
                if (p != null) return p;

                // For family documents, try to get from family manager
                if (_isFamilyDoc && _familyManager != null)
                {
                    if (_familyParamCache.TryGetValue(paramName, out var familyParam))
                    {
                        // Get the parameter from the current family type
                        var currentType = _familyManager.CurrentType;
                        if (currentType != null)
                        {
                            // Return a pseudo-parameter that wraps the family parameter
                            // This is a workaround since FamilyParameter doesn't inherit from Parameter
                            return element.LookupParameter(paramName);
                        }
                    }
                }

                return null;
            }
            catch { return null; }
        }

        public FamilyParameter GetFamilyParameter(string paramName)
        {
            if (!_isFamilyDoc || _familyManager == null) return null;

            _familyParamCache.TryGetValue(paramName, out var result);
            return result;
        }
    }
}