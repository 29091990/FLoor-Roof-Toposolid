using Autodesk.Revit.DB;
using System.Reflection;

namespace FloorRoofTopo
{
    /// <summary>
    /// Essential helper for Revit 2020-2026 compatibility.
    /// Prevents MissingMethodException in 2026 (.Value only)
    /// and MissingMemberException in pre-2024 (.IntegerValue only).
    /// </summary>
    public static class ElementIdHelper
    {
        private static readonly PropertyInfo _valueProp =
            typeof(ElementId).GetProperty("Value");
        private static readonly PropertyInfo _intValueProp =
            typeof(ElementId).GetProperty("IntegerValue");

        /// <summary>
        /// Gets the numeric value of an ElementId safely across Revit 2020-2026.
        /// </summary>
        public static long GetValue(ElementId id)
        {
            if (id == null) return -1;

            // Try .Value first (Revit 2024+)
            if (_valueProp != null)
            {
                try { return (long)_valueProp.GetValue(id); } catch { }
            }

            // Fallback to .IntegerValue (Revit 2020-2025)
            if (_intValueProp != null)
            {
                try { return (int)_intValueProp.GetValue(id); } catch { }
            }

            return -1;
        }

        public static bool AreEqual(ElementId id1, ElementId id2)
        {
            if (id1 == null && id2 == null) return true;
            if (id1 == null || id2 == null) return false;
            return GetValue(id1) == GetValue(id2);
        }
    }
}
