using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;

namespace FloorRoofTopo
{
    /// <summary>
    /// Target element type for conversion.
    /// </summary>
    public enum ConvertTarget
    {
        Floor,
        Roof,
        TopoSolid
    }

    /// <summary>
    /// Result of a single element conversion.
    /// </summary>
    public class ConversionResult
    {
        public bool Success { get; set; }
        public ElementId NewElementId { get; set; }
        public string ErrorMessage { get; set; }
        public string WarningMessage { get; set; }
        public ConversionDiagnosticInfo Diagnostics { get; set; }
    }

    /// <summary>
    /// Detailed data for a single shape point during conversion.
    /// Records source position, relative offset, and target position.
    /// </summary>
    public class ShapePointData
    {
        public int Index { get; set; }
        public string VertexType { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double SourceZ { get; set; }
        public double RelativeZ { get; set; }
        public double TargetZ { get; set; }
        public bool Applied { get; set; }
    }

    /// <summary>
    /// Diagnostic information collected during an element conversion.
    /// Used by DiagnosticDialog to display detailed analysis.
    /// </summary>
    public class ConversionDiagnosticInfo
    {
        public string SourceType { get; set; }
        public long SourceId { get; set; }
        public string TargetType { get; set; }
        public long TargetId { get; set; }
        public string LevelName { get; set; }
        public double LevelElevation { get; set; }
        public double SourceOffset { get; set; }
        public double SourceBaseZ { get; set; }
        public double TargetBaseZ { get; set; }
        public int BoundaryLoops { get; set; }
        public int TotalCurves { get; set; }
        public List<ShapePointData> ShapePoints { get; set; } = new List<ShapePointData>();
        public int AppliedCount { get; set; }
        public bool Success { get; set; }
        public string Message { get; set; }
    }

    /// <summary>
    /// Core conversion engine: extracts boundary + shape points from source,
    /// creates target element, and re-applies shape editing.
    /// </summary>
    public static class ConversionEngine
    {
        // ================================================================
        //  BOUNDARY EXTRACTION
        // ================================================================

        /// <summary>
        /// Extract boundary CurveLoops from any supported element.
        /// Priority: Sketch → RoofProfiles → Geometry fallback.
        /// </summary>
        public static IList<CurveLoop> ExtractBoundary(Document doc, Element element)
        {
            // 1. Try Sketch (most reliable for Floor, Roof, TopoSolid)
            IList<CurveLoop> loops = GetFromSketch(doc, element);
            if (loops != null && loops.Count > 0) return loops;

            // 2. For FootPrintRoof, try GetProfiles
            FootPrintRoof fpRoof = element as FootPrintRoof;
            if (fpRoof != null)
            {
                loops = GetFromRoofProfiles(fpRoof);
                if (loops != null && loops.Count > 0) return loops;
            }

            // 3. Geometry fallback (bottom face edge loops)
            return GetFromGeometry(element);
        }

        /// <summary>
        /// Extract boundary from the element's dependent Sketch.
        /// </summary>
        private static IList<CurveLoop> GetFromSketch(Document doc, Element element)
        {
            try
            {
                var filter = new ElementClassFilter(typeof(Sketch));
                ICollection<ElementId> depIds = element.GetDependentElements(filter);

                foreach (ElementId id in depIds)
                {
                    Sketch sketch = doc.GetElement(id) as Sketch;
                    if (sketch == null) continue;

                    IList<CurveLoop> result = new List<CurveLoop>();
                    foreach (CurveArray curveArray in sketch.Profile)
                    {
                        CurveLoop loop = new CurveLoop();
                        foreach (Curve curve in curveArray)
                        {
                            loop.Append(curve);
                        }
                        result.Add(loop);
                    }

                    if (result.Count > 0) return result;
                }
            }
            catch { }

            return null;
        }

        /// <summary>
        /// Extract boundary from FootPrintRoof.GetProfiles().
        /// </summary>
        private static IList<CurveLoop> GetFromRoofProfiles(FootPrintRoof roof)
        {
            try
            {
                ModelCurveArrArray profiles = roof.GetProfiles();
                IList<CurveLoop> result = new List<CurveLoop>();

                foreach (ModelCurveArray mca in profiles)
                {
                    CurveLoop loop = new CurveLoop();
                    foreach (ModelCurve mc in mca)
                    {
                        loop.Append(mc.GeometryCurve);
                    }
                    result.Add(loop);
                }

                return result.Count > 0 ? result : null;
            }
            catch { return null; }
        }

        /// <summary>
        /// Extract boundary from geometry by finding the bottom face.
        /// This is the last-resort fallback for elements without accessible sketches.
        /// </summary>
        private static IList<CurveLoop> GetFromGeometry(Element element)
        {
            try
            {
                Options opts = new Options();
                opts.ComputeReferences = true;
                GeometryElement geom = element.get_Geometry(opts);
                if (geom == null) return null;

                foreach (GeometryObject gObj in geom)
                {
                    Solid solid = gObj as Solid;
                    if (solid == null || solid.Volume < 1e-9) continue;

                    // Find the bottom face (normal pointing down)
                    PlanarFace bestFace = null;
                    double bestArea = 0;

                    foreach (Face face in solid.Faces)
                    {
                        PlanarFace pf = face as PlanarFace;
                        if (pf == null) continue;
                        if (pf.FaceNormal.Z < -0.8) // Approximately downward
                        {
                            double area = pf.Area;
                            if (area > bestArea)
                            {
                                bestArea = area;
                                bestFace = pf;
                            }
                        }
                    }

                    if (bestFace != null)
                    {
                        return bestFace.GetEdgesAsCurveLoops();
                    }
                }
            }
            catch { }

            return null;
        }

        /// <summary>
        /// Project all curves in CurveLoops to a specific Z elevation.
        /// Ensures the boundary is planar for Floor.Create / NewFootPrintRoof.
        /// </summary>
        public static IList<CurveLoop> FlattenToZ(IList<CurveLoop> loops, double z)
        {
            var result = new List<CurveLoop>();

            foreach (CurveLoop loop in loops)
            {
                CurveLoop newLoop = new CurveLoop();
                foreach (Curve curve in loop)
                {
                    XYZ p0 = curve.GetEndPoint(0);
                    XYZ p1 = curve.GetEndPoint(1);
                    XYZ np0 = new XYZ(p0.X, p0.Y, z);
                    XYZ np1 = new XYZ(p1.X, p1.Y, z);

                    if (np0.DistanceTo(np1) < 1e-9) continue;

                    if (curve is Line)
                    {
                        newLoop.Append(Line.CreateBound(np0, np1));
                    }
                    else if (curve is Arc)
                    {
                        // Project the midpoint too for Arc reconstruction
                        XYZ mid = curve.Evaluate(0.5, true);
                        XYZ nmid = new XYZ(mid.X, mid.Y, z);
                        newLoop.Append(Arc.Create(np0, np1, nmid));
                    }
                    else
                    {
                        // For other curves (NurbSpline, etc.), approximate with line
                        newLoop.Append(Line.CreateBound(np0, np1));
                    }
                }

                if (newLoop.Count() > 0)
                    result.Add(newLoop);
            }

            return result;
        }

        // ================================================================
        //  SLABSHAPEEDITOR HELPERS (Reflection-based cross-version)
        // ================================================================

        /// <summary>
        /// Get SlabShapeEditor from element using reflection.
        /// Revit 2024+: method GetSlabShapeEditor()
        /// Revit 2019-2023: property SlabShapeEditor
        /// </summary>
        public static SlabShapeEditor GetEditor(Element element)
        {
            if (element == null) return null;

            // Try method GetSlabShapeEditor() first (Revit 2024+)
            try
            {
                var method = element.GetType().GetMethod("GetSlabShapeEditor",
                    BindingFlags.Public | BindingFlags.Instance,
                    null, Type.EmptyTypes, null);
                if (method != null)
                    return method.Invoke(element, null) as SlabShapeEditor;
            }
            catch { }

            // Fallback: property SlabShapeEditor (Revit 2019-2023)
            try
            {
                var prop = element.GetType().GetProperty("SlabShapeEditor",
                    BindingFlags.Public | BindingFlags.Instance);
                if (prop != null)
                    return prop.GetValue(element) as SlabShapeEditor;
            }
            catch { }

            return null;
        }

        /// <summary>
        /// Extract all shape vertices from a SlabShapeEditor-enabled element.
        /// Returns empty list if shape editing is not enabled.
        /// </summary>
        public static List<XYZ> ExtractShapePoints(Element element)
        {
            var points = new List<XYZ>();

            SlabShapeEditor editor = GetEditor(element);
            if (editor == null) return points;

            try
            {
                if (!editor.IsEnabled) return points;

                foreach (SlabShapeVertex vertex in editor.SlabShapeVertices)
                {
                    points.Add(vertex.Position);
                }
            }
            catch { }

            return points;
        }

        /// <summary>
        /// Apply shape points to a target element's SlabShapeEditor.
        /// Enables the editor if not already enabled.
        /// NOTE: For cross-type conversion, use ApplyShapePointsAdjusted instead.
        /// </summary>
        public static int ApplyShapePoints(Element element, List<XYZ> points)
        {
            if (points == null || points.Count == 0) return 0;

            SlabShapeEditor editor = GetEditor(element);
            if (editor == null) return 0;

            if (!editor.IsEnabled)
                editor.Enable();

            int applied = 0;
            foreach (XYZ pt in points)
            {
                try
                {
                    if (AddPointSafe(editor, pt))
                        applied++;
                }
                catch
                {
                    // Individual point failure — continue with others
                }
            }

            return applied;
        }

        /// <summary>
        /// Add a point to SlabShapeEditor with cross-version reflection.
        /// Revit 2026+: AddPoint(XYZ), Revit 2019-2025: DrawPoint(XYZ).
        /// </summary>
        private static bool AddPointSafe(SlabShapeEditor editor, XYZ point)
        {
            if (editor == null || point == null) return false;

            // 1. Try AddPoint first (Revit 2026+)
            try
            {
                var method = editor.GetType().GetMethod("AddPoint",
                    BindingFlags.Public | BindingFlags.Instance,
                    null, new Type[] { typeof(XYZ) }, null);
                if (method != null)
                {
                    method.Invoke(editor, new object[] { point });
                    return true;
                }
            }
            catch { }

            // 2. Fallback: DrawPoint (Revit 2019-2025)
            try
            {
                var method = editor.GetType().GetMethod("DrawPoint",
                    BindingFlags.Public | BindingFlags.Instance,
                    null, new Type[] { typeof(XYZ) }, null);
                if (method != null)
                {
                    method.Invoke(editor, new object[] { point });
                    return true;
                }
            }
            catch { }

            return false;
        }

        // ================================================================
        //  ELEMENT CREATION
        // ================================================================

        /// <summary>
        /// Create a Floor element from boundary CurveLoops.
        /// </summary>
        public static Floor CreateFloor(Document doc, IList<CurveLoop> boundary, ElementId levelId)
        {
            ElementId typeId = GetDefaultTypeId(doc, typeof(FloorType));
            if (typeId == null || typeId == ElementId.InvalidElementId)
                throw new InvalidOperationException("Không tìm thấy Floor Type trong project.");

            Level level = doc.GetElement(levelId) as Level;
            double levelZ = level?.Elevation ?? 0;
            IList<CurveLoop> flatBoundary = FlattenToZ(boundary, levelZ);

#if REVIT_LEGACY
            // Revit 2020-2021: use deprecated NewFloor
            CurveArray curveArray = CurveLoopToCurveArray(flatBoundary[0]);
            FloorType floorType = doc.GetElement(typeId) as FloorType;
            return doc.Create.NewFloor(curveArray, floorType, level, false);
#else
            // Revit 2022+: use Floor.Create
            return Floor.Create(doc, flatBoundary, typeId, levelId);
#endif
        }

        /// <summary>
        /// Create a FootPrintRoof from boundary CurveLoops.
        /// All edges are set to flat (no slope) to allow pure shape editing.
        /// </summary>
        public static FootPrintRoof CreateRoof(Document doc, IList<CurveLoop> boundary, ElementId levelId)
        {
            ElementId typeId = FindBasicRoofTypeId(doc);
            if (typeId == null || typeId == ElementId.InvalidElementId)
                throw new InvalidOperationException("Không tìm thấy Roof Type trong project.");

            Level level = doc.GetElement(levelId) as Level;
            RoofType roofType = doc.GetElement(typeId) as RoofType;

            double levelZ = level?.Elevation ?? 0;
            IList<CurveLoop> flatBoundary = FlattenToZ(boundary, levelZ);

            // NewFootPrintRoof uses CurveArray (outer loop only)
            CurveArray curveArray = CurveLoopToCurveArray(flatBoundary[0]);

            ModelCurveArray modelCurves = new ModelCurveArray();
            FootPrintRoof roof = doc.Create.NewFootPrintRoof(
                curveArray, level, roofType, out modelCurves);

            // Set all edges to flat (no slope) for pure shape editing
            if (roof != null && modelCurves != null)
            {
                foreach (ModelCurve mc in modelCurves)
                {
                    roof.set_DefinesSlope(mc, false);
                }
            }

            return roof;
        }

#if REVIT2024_PLUS
        /// <summary>
        /// Create a Toposolid element from boundary CurveLoops (Revit 2024+).
        /// </summary>
        public static Toposolid CreateTopoSolid(Document doc, IList<CurveLoop> boundary, ElementId levelId)
        {
            return CreateTopoSolid(doc, boundary, levelId, null);
        }

        /// <summary>
        /// Create a Toposolid with optional interior shape points (Revit 2024+).
        /// If interiorPoints is provided, tries to use Toposolid.Create(doc, boundary, points, typeId, levelId)
        /// which directly defines the surface shape — more reliable than using SlabShapeEditor afterward.
        /// </summary>
        public static Toposolid CreateTopoSolid(Document doc, IList<CurveLoop> boundary, ElementId levelId,
            IList<XYZ> interiorPoints)
        {
            ElementId typeId = FindToposolidTypeId(doc);
            if (typeId == null || typeId == ElementId.InvalidElementId)
                throw new InvalidOperationException(
                    "Không tìm thấy Toposolid Type trong project.\n" +
                    "Hãy đảm bảo project có ít nhất một Toposolid Type.");

            Level level = doc.GetElement(levelId) as Level;
            double levelZ = level?.Elevation ?? 0;
            IList<CurveLoop> flatBoundary = FlattenToZ(boundary, levelZ);

            // Try the overload with interior points first
            if (interiorPoints != null && interiorPoints.Count > 0)
            {
                try
                {
                    // Toposolid.Create(Document, IList<CurveLoop>, IList<XYZ>, ElementId, ElementId)
                    var method = typeof(Toposolid).GetMethod("Create",
                        BindingFlags.Public | BindingFlags.Static,
                        null,
                        new Type[] {
                            typeof(Document),
                            typeof(IList<CurveLoop>),
                            typeof(IList<XYZ>),
                            typeof(ElementId),
                            typeof(ElementId)
                        }, null);

                    if (method != null)
                    {
                        var result = method.Invoke(null, new object[] {
                            doc, flatBoundary, interiorPoints, typeId, levelId
                        }) as Toposolid;

                        if (result != null) return result;
                    }
                }
                catch { /* Fallback to flat creation */ }
            }

            return Toposolid.Create(doc, flatBoundary, typeId, levelId);
        }

        /// <summary>
        /// Find a valid ToposolidType ElementId in the document.
        /// </summary>
        private static ElementId FindToposolidTypeId(Document doc)
        {
            try
            {
                var type = new FilteredElementCollector(doc)
                    .OfClass(typeof(ToposolidType))
                    .FirstOrDefault();
                if (type != null) return type.Id;
            }
            catch { }

            // Fallback: by category (Topography)
            try
            {
                return new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Topography)
                    .WhereElementIsElementType()
                    .FirstElementId();
            }
            catch { }

            return ElementId.InvalidElementId;
        }
#endif

        // ================================================================
        //  UTILITY HELPERS
        // ================================================================

        /// <summary>
        /// Convert CurveLoop to legacy CurveArray (for NewFloor / NewFootPrintRoof).
        /// </summary>
        public static CurveArray CurveLoopToCurveArray(CurveLoop loop)
        {
            CurveArray arr = new CurveArray();
            foreach (Curve c in loop)
            {
                arr.Append(c);
            }
            return arr;
        }

        /// <summary>
        /// Get the default ElementType ID for a given type class.
        /// </summary>
        public static ElementId GetDefaultTypeId(Document doc, Type typeClass)
        {
            var collector = new FilteredElementCollector(doc)
                .OfClass(typeClass);
            Element first = collector.FirstOrDefault();
            return first?.Id ?? ElementId.InvalidElementId;
        }

        /// <summary>
        /// Find a basic/generic RoofType, skipping Glazing/Glass/Curtain types.
        /// Priority: "Generic" > any non-glazing type > first available.
        /// </summary>
        public static ElementId FindBasicRoofTypeId(Document doc)
        {
            var allTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(RoofType))
                .Cast<RoofType>()
                .ToList();

            if (allTypes.Count == 0)
                return ElementId.InvalidElementId;

            // Skip keywords for glazing/curtain types
            string[] skipKeywords = { "glazing", "glass", "curtain", "kính" };

            // 1. Prefer "Generic" type
            var generic = allTypes.FirstOrDefault(t =>
                t.Name.IndexOf("Generic", StringComparison.OrdinalIgnoreCase) >= 0 &&
                !skipKeywords.Any(k => t.Name.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0));
            if (generic != null) return generic.Id;

            // 2. Any non-glazing type
            var basic = allTypes.FirstOrDefault(t =>
                !skipKeywords.Any(k => t.Name.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0));
            if (basic != null) return basic.Id;

            // 3. Fallback: first available
            return allTypes[0].Id;
        }

        /// <summary>
        /// Get the Level ID from an element, trying multiple parameter sources.
        /// </summary>
        public static ElementId GetLevelId(Document doc, Element element)
        {
            // 1. Direct LevelId property
            if (element.LevelId != null && element.LevelId != ElementId.InvalidElementId)
                return element.LevelId;

            // 2. LEVEL_PARAM (common for floors)
            Parameter lvlParam = element.get_Parameter(BuiltInParameter.LEVEL_PARAM);
            if (lvlParam != null && lvlParam.AsElementId() != ElementId.InvalidElementId)
                return lvlParam.AsElementId();

            // 3. ROOF_BASE_LEVEL_PARAM (for roofs)
            Parameter roofLvl = element.get_Parameter(BuiltInParameter.ROOF_BASE_LEVEL_PARAM);
            if (roofLvl != null && roofLvl.AsElementId() != ElementId.InvalidElementId)
                return roofLvl.AsElementId();

            // 4. SCHEDULE_LEVEL_PARAM (general fallback)
            Parameter schedLvl = element.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM);
            if (schedLvl != null && schedLvl.AsElementId() != ElementId.InvalidElementId)
                return schedLvl.AsElementId();

            // 5. Last resort: first level in document
            Level firstLevel = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .FirstOrDefault() as Level;
            return firstLevel?.Id ?? ElementId.InvalidElementId;
        }

        // ================================================================
        //  ELEVATION & OFFSET HELPERS
        // ================================================================

        /// <summary>
        /// Get height offset from level for any element type (Floor, Roof, TopoSolid).
        /// Returns the offset in internal units (feet).
        /// </summary>
        public static double GetHeightOffset(Element element)
        {
            if (element == null) return 0;

            // Floor: Height Offset From Level
            Parameter p = element.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
            if (p != null && p.HasValue) return p.AsDouble();

            // Roof: Base Offset From Level
            p = element.get_Parameter(BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM);
            if (p != null && p.HasValue) return p.AsDouble();

            return 0;
        }

        /// <summary>
        /// Set height offset on any element type (Floor or Roof).
        /// Handles cross-type parameter mapping.
        /// </summary>
        public static bool SetHeightOffset(Element element, double offset)
        {
            if (element == null) return false;

            Parameter p = element.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
            if (p != null && !p.IsReadOnly) { p.Set(offset); return true; }

            p = element.get_Parameter(BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM);
            if (p != null && !p.IsReadOnly) { p.Set(offset); return true; }

            return false;
        }

        /// <summary>
        /// Calculate the base elevation of an element: level elevation + height offset.
        /// This is the Z coordinate of the element's base plane.
        /// </summary>
        public static double GetBaseElevation(Document doc, Element element)
        {
            double baseZ = 0;
            ElementId levelId = GetLevelId(doc, element);
            if (levelId != ElementId.InvalidElementId)
            {
                Level level = doc.GetElement(levelId) as Level;
                if (level != null) baseZ = level.Elevation;
            }
            baseZ += GetHeightOffset(element);
            return baseZ;
        }

        /// <summary>
        /// Extract shape points with detailed diagnostic data.
        /// Each point records its absolute position and relative Z offset
        /// from the source base elevation. This enables correct Z-adjustment
        /// when transferring to a different element type.
        /// </summary>
        public static List<ShapePointData> ExtractShapePointsDetailed(
            Element element, double sourceBaseZ)
        {
            var points = new List<ShapePointData>();
            SlabShapeEditor editor = GetEditor(element);
            if (editor == null) return points;

            try
            {
                if (!editor.IsEnabled) return points;

                int idx = 0;
                foreach (SlabShapeVertex vertex in editor.SlabShapeVertices)
                {
                    string vType = "Unknown";
                    try { vType = vertex.VertexType.ToString(); }
                    catch { }

                    points.Add(new ShapePointData
                    {
                        Index = idx++,
                        VertexType = vType,
                        X = vertex.Position.X,
                        Y = vertex.Position.Y,
                        SourceZ = vertex.Position.Z,
                        RelativeZ = vertex.Position.Z - sourceBaseZ,
                        Applied = false
                    });
                }
            }
            catch { }

            return points;
        }

        /// <summary>
        /// Apply shape points with Z adjustment for cross-type conversion.
        /// 
        /// Strategy: each shape point's Z is recalculated as
        ///   targetZ = targetBaseZ + relativeZ
        /// where relativeZ = originalZ - sourceBaseZ.
        /// 
        /// Critical: After enabling the SlabShapeEditor, we MUST call
        /// doc.Regenerate() to ensure the editor's vertex list is fully
        /// initialized. Without this, the vertex list may be empty,
        /// causing DrawPoint to create new points instead of modifying
        /// existing corner/edge vertices.
        /// 
        /// The algorithm uses two passes:
        ///   Pass 1: Match source points to existing target vertices by XY
        ///           proximity and modify them using DrawPoint with the
        ///           target vertex's EXACT XY coordinates.
        ///   Pass 2: Add remaining unmatched source points (interior points)
        ///           as new points.
        /// </summary>
        public static int ApplyShapePointsAdjusted(
            Element element, List<ShapePointData> shapeData, double targetBaseZ)
        {
            if (shapeData == null || shapeData.Count == 0) return 0;

            SlabShapeEditor editor = GetEditor(element);
            if (editor == null) return 0;

            if (!editor.IsEnabled)
                editor.Enable();

            // ── CRITICAL: Regenerate to initialize editor vertices ───────
            // Without this, the vertex list may be empty after Enable(),
            // causing vertex matching to fail completely.
            try { element.Document.Regenerate(); }
            catch { }

            // ── Collect existing target vertices (corners + edges) ────────
            var existingVertices = new List<XYZ>();
            try
            {
                foreach (SlabShapeVertex v in editor.SlabShapeVertices)
                {
                    existingVertices.Add(v.Position);
                }
            }
            catch { }

            // Track which existing vertices and source points have been matched
            var matchedExisting = new bool[existingVertices.Count];
            var matchedSource = new bool[shapeData.Count];

            int applied = 0;

            // ── Pass 1: Match source points to existing target vertices ──
            // For each existing target vertex, find the closest source point.
            // This ensures ALL target vertices get processed (not just those
            // that happen to match a source point in the reverse direction).
            for (int ev = 0; ev < existingVertices.Count; ev++)
            {
                var targetVtx = existingVertices[ev];
                int bestSrcIdx = -1;
                double bestDist = 0.5; // ~150mm tolerance for cross-type matching

                for (int si = 0; si < shapeData.Count; si++)
                {
                    if (matchedSource[si]) continue;

                    double dx = targetVtx.X - shapeData[si].X;
                    double dy = targetVtx.Y - shapeData[si].Y;
                    double dist = Math.Sqrt(dx * dx + dy * dy);

                    if (dist < bestDist)
                    {
                        bestDist = dist;
                        bestSrcIdx = si;
                    }
                }

                if (bestSrcIdx >= 0)
                {
                    var data = shapeData[bestSrcIdx];
                    double adjustedZ = targetBaseZ + data.RelativeZ;
                    data.TargetZ = adjustedZ;

                    // Use target vertex's EXACT XY → DrawPoint recognizes it
                    XYZ adjustedPoint = new XYZ(targetVtx.X, targetVtx.Y, adjustedZ);

                    try
                    {
                        if (AddPointSafe(editor, adjustedPoint))
                        {
                            data.Applied = true;
                            applied++;
                        }
                    }
                    catch { }

                    matchedExisting[ev] = true;
                    matchedSource[bestSrcIdx] = true;
                }
            }

            // ── Pass 2: Add remaining unmatched source points ────────────
            // These are interior points that don't have a corresponding
            // existing target vertex. Add them as new points.
            for (int si = 0; si < shapeData.Count; si++)
            {
                if (matchedSource[si]) continue;

                var data = shapeData[si];
                double adjustedZ = targetBaseZ + data.RelativeZ;
                data.TargetZ = adjustedZ;

                XYZ adjustedPoint = new XYZ(data.X, data.Y, adjustedZ);

                try
                {
                    if (AddPointSafe(editor, adjustedPoint))
                    {
                        data.Applied = true;
                        applied++;
                    }
                }
                catch { }
            }

            return applied;
        }

        // ================================================================
        //  TYPE QUERY HELPERS
        // ================================================================

        /// <summary>
        /// Check if TopoSolid is available in this Revit version.
        /// </summary>
        public static bool IsTopoSolidAvailable()
        {
#if REVIT2024_PLUS
            return true;
#else
            return false;
#endif
        }

        /// <summary>
        /// Check if an element is one of the supported source types.
        /// </summary>
        public static bool IsValidSourceElement(Element element)
        {
            if (element == null) return false;
            if (element is Floor) return true;
            if (element is FootPrintRoof) return true;
#if REVIT2024_PLUS
            if (element is Toposolid) return true;
#endif
            return false;
        }

        /// <summary>
        /// Get human-readable type name for display.
        /// </summary>
        public static string GetElementTypeName(Element element)
        {
            if (element is Floor) return "Floor";
            if (element is FootPrintRoof) return "Roof";
#if REVIT2024_PLUS
            if (element is Toposolid) return "TopoSolid";
#endif
            if (element is RoofBase) return "Roof (Unsupported)";
            return "Unknown";
        }

        /// <summary>
        /// Check if element already matches the target type.
        /// </summary>
        public static bool IsSameType(Element element, ConvertTarget target)
        {
            switch (target)
            {
                case ConvertTarget.Floor: return element is Floor;
                case ConvertTarget.Roof: return element is FootPrintRoof;
#if REVIT2024_PLUS
                case ConvertTarget.TopoSolid: return element is Toposolid;
#endif
                default: return false;
            }
        }

        // ================================================================
        //  MAIN CONVERSION ORCHESTRATOR
        // ================================================================

        /// <summary>
        /// Convert a single source element to the target type.
        /// 
        /// Key improvement: shape points are transferred using relative Z offsets,
        /// preventing the vertical spike issue when source and target have
        /// different base elevations (e.g., Floor with height offset → Roof).
        /// 
        /// Workflow:
        ///   1. Extract boundary + shape points from source
        ///   2. Create target element at level elevation (offset = 0)
        ///   3. Apply shape points with adjusted Z = levelZ + relativeOffset
        ///   4. Set height offset on target = source height offset
        ///      → element shifts, shape points auto-adjust to correct absolute Z
        /// </summary>
        public static ConversionResult Convert(
            Document doc, Element source, ConvertTarget target, bool deleteSource)
        {
            var result = new ConversionResult();
            var diag = new ConversionDiagnosticInfo();
            result.Diagnostics = diag;

            try
            {
                // ── Source info ──────────────────────────────────────────
                diag.SourceType = GetElementTypeName(source);
                diag.SourceId = ElementIdHelper.GetValue(source.Id);
                diag.TargetType = target.ToString();

                // Skip if source and target are the same type
                if (IsSameType(source, target))
                {
                    result.WarningMessage =
                        $"Bỏ qua {diag.SourceType} — đã là loại đích.";
                    result.Success = true;
                    result.NewElementId = source.Id;
                    diag.Success = true;
                    diag.Message = result.WarningMessage;
                    return result;
                }

                // ── Level info ──────────────────────────────────────────
                ElementId levelId = GetLevelId(doc, source);
                if (levelId == ElementId.InvalidElementId)
                {
                    result.ErrorMessage =
                        "Không thể xác định Level của phần tử nguồn.";
                    diag.Success = false;
                    diag.Message = result.ErrorMessage;
                    return result;
                }

                Level level = doc.GetElement(levelId) as Level;
                diag.LevelName = level?.Name ?? "N/A";
                diag.LevelElevation = level?.Elevation ?? 0;
                diag.SourceOffset = GetHeightOffset(source);
                diag.SourceBaseZ = diag.LevelElevation + diag.SourceOffset;

                // ── 1. Extract boundary ─────────────────────────────────
                IList<CurveLoop> boundary = ExtractBoundary(doc, source);
                if (boundary == null || boundary.Count == 0)
                {
                    result.ErrorMessage =
                        "Không thể trích xuất đường biên từ phần tử nguồn.";
                    diag.Success = false;
                    diag.Message = result.ErrorMessage;
                    return result;
                }

                diag.BoundaryLoops = boundary.Count;
                int totalCurves = 0;
                foreach (var loop in boundary)
                    foreach (Curve c in loop)
                        totalCurves++;
                diag.TotalCurves = totalCurves;

                // ── 2. Extract shape points with relative offsets ────────
                List<ShapePointData> shapeData =
                    ExtractShapePointsDetailed(source, diag.SourceBaseZ);
                diag.ShapePoints = shapeData;

                // ── 3. Create target element ────────────────────────────
                Element newElement = null;
                switch (target)
                {
                    case ConvertTarget.Floor:
                        newElement = CreateFloor(doc, boundary, levelId);
                        break;
                    case ConvertTarget.Roof:
                        newElement = CreateRoof(doc, boundary, levelId);
                        break;
                    case ConvertTarget.TopoSolid:
#if REVIT2024_PLUS
                        {
                            // For TopoSolid: pass shape points directly during creation
                            // This is more reliable than post-creation SlabShapeEditor
                            var topoPoints = new List<XYZ>();
                            double topoBaseZ = diag.LevelElevation;
                            foreach (var sp in shapeData)
                            {
                                double adjZ = topoBaseZ + sp.RelativeZ;
                                topoPoints.Add(new XYZ(sp.X, sp.Y, adjZ));
                            }
                            newElement = CreateTopoSolid(doc, boundary, levelId, topoPoints);
                        }
#else
                        result.ErrorMessage =
                            "TopoSolid không khả dụng trong phiên bản Revit này.";
                        diag.Success = false;
                        diag.Message = result.ErrorMessage;
                        return result;
#endif
                        break;
                }

                if (newElement == null)
                {
                    result.ErrorMessage =
                        $"Không thể tạo phần tử {target}.";
                    diag.Success = false;
                    diag.Message = result.ErrorMessage;
                    return result;
                }

                diag.TargetId = ElementIdHelper.GetValue(newElement.Id);

                // ── 4. Regenerate to initialize new element ─────────────
                doc.Regenerate();

                // ── 5. Read ACTUAL target base Z (not assumed) ───────────
                // Some element types may have default offsets from their type.
                // Reading the actual base elevation ensures correct Z math.
                double targetBaseZ = GetBaseElevation(doc, newElement);
                diag.TargetBaseZ = targetBaseZ;

                // ── 6. Apply shape points with Z adjustment ─────────────
                // Points are applied at: targetBaseZ + relativeOffset
                // This ensures correct relative deformation regardless of
                // source height offset.
                if (shapeData.Count > 0)
                {
                    try
                    {
                        int applied = ApplyShapePointsAdjusted(
                            newElement, shapeData, targetBaseZ);
                        diag.AppliedCount = applied;

                        if (applied < shapeData.Count)
                        {
                            result.WarningMessage =
                                $"Đã áp dụng {applied}/{shapeData.Count} điểm nhấc.";
                        }
                    }
                    catch (Exception ex)
                    {
                        result.WarningMessage =
                            $"Điểm nhấc áp dụng một phần: {ex.Message}";
                    }
                }

                // ── 7. Copy height offset (cross-type compatible) ───────
                // After shape points are applied relative to levelZ,
                // setting the offset shifts the entire element (including
                // shape deformations) to match source absolute positions.
                if (Math.Abs(diag.SourceOffset) > 1e-9)
                {
                    SetHeightOffset(newElement, diag.SourceOffset);
                }

                // ── 8. Delete source if requested ───────────────────────
                if (deleteSource)
                {
                    doc.Delete(source.Id);
                }

                result.Success = true;
                result.NewElementId = newElement.Id;
                diag.Success = true;
                diag.Message = "Chuyển đổi thành công";
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"Lỗi chuyển đổi: {ex.Message}";
                diag.Success = false;
                diag.Message = result.ErrorMessage;
            }

            return result;
        }
    }
}
