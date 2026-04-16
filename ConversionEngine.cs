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

        // ── Pending shape editing (applied in a separate transaction) ──
        // On Revit 2026+, SlabShapeEditor.Enable() fails within the same
        // transaction that created the element. Shape data is stored here
        // and applied after the creation transaction is committed.
        public bool HasPendingShapeEditing { get; set; }
        public List<ShapePointData> PendingShapeData { get; set; }
        public double PendingShapeBaseZ { get; set; }
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
        public string ErrorInfo { get; set; }
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
        public string EditorMethods { get; set; }
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

        /// <summary>
        /// Subdivide boundary curves at edge vertex positions from source shape data.
        ///
        /// Problem: SlabShapeEditor can only modify existing boundary vertices or
        /// add interior points. It CANNOT add new points ON boundary edges.
        /// When a TopoSolid has edge vertices (subdivision points along its boundary),
        /// these cannot be transferred to a Floor/Roof because the target boundary
        /// has only corner vertices.
        ///
        /// Solution: Before creating the target element, split each boundary line
        /// segment at positions where source edge vertices exist. This creates
        /// additional corner vertices in the target, allowing shape editing to
        /// modify them via ModifySubElement.
        /// </summary>
        public static IList<CurveLoop> SubdivideBoundaryAtEdgePoints(
            IList<CurveLoop> boundary, List<ShapePointData> shapeData)
        {
            if (boundary == null || shapeData == null) return boundary;

            // Collect edge-type points (XY only)
            var edgePoints = shapeData
                .Where(sp => sp.VertexType == "Edge" || sp.VertexType == "Cạnh")
                .ToList();

            if (edgePoints.Count == 0) return boundary;

            var result = new List<CurveLoop>();

            foreach (CurveLoop loop in boundary)
            {
                CurveLoop newLoop = new CurveLoop();

                foreach (Curve curve in loop)
                {
                    XYZ start = curve.GetEndPoint(0);
                    XYZ end = curve.GetEndPoint(1);

                    if (curve is Line)
                    {
                        // Find edge points that project onto this line segment
                        double segLen = Math.Sqrt(
                            Math.Pow(end.X - start.X, 2) +
                            Math.Pow(end.Y - start.Y, 2));

                        if (segLen < 1e-6)
                        {
                            newLoop.Append(curve);
                            continue;
                        }

                        double dirX = (end.X - start.X) / segLen;
                        double dirY = (end.Y - start.Y) / segLen;

                        // Collect (parameter, edgePoint) pairs for points on this line
                        var splits = new List<Tuple<double, ShapePointData>>();

                        foreach (var ep in edgePoints)
                        {
                            // Project edge point onto line (2D)
                            double vx = ep.X - start.X;
                            double vy = ep.Y - start.Y;
                            double t = vx * dirX + vy * dirY; // dot product

                            // Must be interior to the segment (not at endpoints)
                            if (t < 0.005 || t > segLen - 0.005) continue;

                            // Perpendicular distance
                            double projX = start.X + t * dirX;
                            double projY = start.Y + t * dirY;
                            double perpDist = Math.Sqrt(
                                Math.Pow(ep.X - projX, 2) +
                                Math.Pow(ep.Y - projY, 2));

                            if (perpDist < 0.01) // ~3mm tolerance
                            {
                                splits.Add(Tuple.Create(t, ep));
                            }
                        }

                        if (splits.Count == 0)
                        {
                            newLoop.Append(curve);
                            continue;
                        }

                        // Sort by parameter along the line
                        splits.Sort((a, b) => a.Item1.CompareTo(b.Item1));

                        // Split the line at each edge point
                        XYZ prev = start;
                        foreach (var split in splits)
                        {
                            // Use exact XY from edge point, Z from boundary
                            XYZ splitPt = new XYZ(split.Item2.X, split.Item2.Y, start.Z);

                            if (Math.Sqrt(
                                Math.Pow(prev.X - splitPt.X, 2) +
                                Math.Pow(prev.Y - splitPt.Y, 2)) > 1e-6)
                            {
                                newLoop.Append(Line.CreateBound(prev, splitPt));
                                prev = splitPt;
                            }
                        }

                        // Final segment to end
                        if (Math.Sqrt(
                            Math.Pow(prev.X - end.X, 2) +
                            Math.Pow(prev.Y - end.Y, 2)) > 1e-6)
                        {
                            newLoop.Append(Line.CreateBound(prev, end));
                        }
                    }
                    else
                    {
                        // For arcs/other curves, keep as-is
                        newLoop.Append(curve);
                    }
                }

                if (newLoop.Count() > 0)
                    result.Add(newLoop);
            }

            return result;
        }

        /// <summary>
        /// Ensure the outer CurveLoop is counter-clockwise (required by Floor.Create).
        /// Inner loops (holes) should be clockwise.
        /// Floor.Create throws when loop orientation is wrong.
        /// </summary>
        public static IList<CurveLoop> EnsureCorrectOrientation(IList<CurveLoop> loops)
        {
            if (loops == null || loops.Count == 0) return loops;

            var result = new List<CurveLoop>();

            for (int i = 0; i < loops.Count; i++)
            {
                CurveLoop loop = loops[i];
                bool isCounterClockwise = loop.IsCounterclockwise(
                    new XYZ(0, 0, 1));

                if (i == 0)
                {
                    // Outer loop must be counter-clockwise
                    if (!isCounterClockwise)
                    {
                        loop.Flip();
                    }
                }
                else
                {
                    // Inner loops (holes) must be clockwise
                    if (isCounterClockwise)
                    {
                        loop.Flip();
                    }
                }

                result.Add(loop);
            }

            return result;
        }

        // ================================================================
        //  SLABSHAPEEDITOR HELPERS (Reflection-based cross-version)
        // ================================================================

        /// <summary>
        /// Get SlabShapeEditor from element.
        /// Revit 2024+: direct cast or method GetSlabShapeEditor()
        /// Revit 2019-2023: property SlabShapeEditor (reflection)
        /// IMPORTANT: Always call this to get a FRESH editor reference.
        /// On Revit 2026+, stale references may not reflect Enable() state.
        /// </summary>
        public static SlabShapeEditor GetEditor(Element element)
        {
            if (element == null) return null;

#if REVIT2024_PLUS
            // Direct API calls — avoids reflection issues on .NET 8 (Revit 2025+)
            try
            {
                if (element is Floor floor)
                    return floor.GetSlabShapeEditor();
            }
            catch { }
            try
            {
                if (element is FootPrintRoof roof)
                    return roof.GetSlabShapeEditor();
            }
            catch { }
#endif

            // Reflection fallback for older versions or unexpected types
            try
            {
                var method = element.GetType().GetMethod("GetSlabShapeEditor",
                    BindingFlags.Public | BindingFlags.Instance,
                    null, Type.EmptyTypes, null);
                if (method != null)
                    return method.Invoke(element, null) as SlabShapeEditor;
            }
            catch { }

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

        // ── Thread-local error capture for diagnostics ────────────────
        [ThreadStatic] private static string _lastPointError;

        /// <summary>
        /// Add or modify a point on SlabShapeEditor.
        /// Revit 2020-2025: DrawPoint(XYZ) — Z must be ON the slab face.
        /// Revit 2026+: DrawPoint removed, use AddPoint(XYZ) instead.
        /// </summary>
        private static bool AddPointSafe(SlabShapeEditor editor, XYZ point)
        {
            if (editor == null || point == null) return false;
            try
            {
#if REVIT2026
                editor.AddPoint(point);
#else
                editor.DrawPoint(point);
#endif
                return true;
            }
            catch (Exception ex)
            {
                _lastPointError = $"DrawPoint/AddPoint: {ex.Message}";
                return false;
            }
        }

        /// <summary>
        /// Modify an existing SlabShapeVertex's height offset.
        /// ModifySubElement(SlabShapeVertex, double) exists in ALL Revit versions.
        /// The offset is the elevation relative to the original unedited slab plane.
        /// </summary>
        private static bool ModifyVertexSafe(SlabShapeEditor editor, SlabShapeVertex vertex, double offset)
        {
            if (editor == null || vertex == null) return false;
            try
            {
                editor.ModifySubElement(vertex, offset);
                return true;
            }
            catch (Exception ex)
            {
                _lastPointError = $"ModifySubElement: {ex.Message}";
                return false;
            }
        }

        /// <summary>
        /// Get a summary of editor state for diagnostics.
        /// </summary>
        private static string GetEditorMethodsSummary(SlabShapeEditor editor)
        {
            if (editor == null) return "Editor is null";
            try
            {
                int vtxCount = 0;
                try { vtxCount = editor.SlabShapeVertices.Size; } catch { }
                return $"Enabled={editor.IsEnabled}; Vertices={vtxCount}";
            }
            catch { return "Error reading editor"; }
        }

        // ================================================================
        //  ELEMENT CREATION
        // ================================================================

        /// <summary>
        /// Create a Floor element from boundary CurveLoops.
        /// Handles orientation enforcement and fallback strategies for
        /// boundaries from TopoSolid/Roof that may have incompatible winding.
        /// </summary>
        public static Floor CreateFloor(Document doc, IList<CurveLoop> boundary, ElementId levelId)
        {
            ElementId typeId = FindFloorTypeId(doc);
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
            // Floor.Create is strict about loop orientation:
            //   - Outer loop must be counter-clockwise
            //   - Inner loops (holes) must be clockwise
            // TopoSolid/Roof boundaries may have the wrong winding.

            // Attempt 1: with corrected orientation
            try
            {
                IList<CurveLoop> oriented = EnsureCorrectOrientation(flatBoundary);
                return Floor.Create(doc, oriented, typeId, levelId);
            }
            catch { }

            // Attempt 2: use only the outer loop (skip problematic inner loops)
            try
            {
                // Find the largest loop by area (= outer boundary)
                CurveLoop outerLoop = null;
                double maxArea = 0;
                foreach (var loop in flatBoundary)
                {
                    try
                    {
                        // Approximate area using shoelace on endpoints
                        double area = Math.Abs(GetLoopArea(loop));
                        if (area > maxArea)
                        {
                            maxArea = area;
                            outerLoop = loop;
                        }
                    }
                    catch { if (outerLoop == null) outerLoop = loop; }
                }

                if (outerLoop != null)
                {
                    // Ensure CCW
                    try
                    {
                        if (!outerLoop.IsCounterclockwise(new XYZ(0, 0, 1)))
                            outerLoop.Flip();
                    }
                    catch { }

                    var singleLoop = new List<CurveLoop> { outerLoop };
                    return Floor.Create(doc, singleLoop, typeId, levelId);
                }
            }
            catch { }

            // Attempt 3: try flipped orientation (some TopoSolids have reversed normals)
            try
            {
                foreach (var loop in flatBoundary)
                {
                    loop.Flip();
                }
                return Floor.Create(doc, flatBoundary, typeId, levelId);
            }
            catch { }

            // Attempt 4: last resort — NewFloor via reflection (more tolerant, deprecated API)
            try
            {
                CurveArray curveArray = CurveLoopToCurveArray(flatBoundary[0]);
                FloorType floorType = doc.GetElement(typeId) as FloorType;
                var createObj = doc.Create;
                var newFloorMethod = createObj.GetType().GetMethod("NewFloor",
                    BindingFlags.Public | BindingFlags.Instance,
                    null,
                    new Type[] { typeof(CurveArray), typeof(FloorType), typeof(Level), typeof(bool) },
                    null);
                if (newFloorMethod != null)
                {
                    var result = newFloorMethod.Invoke(createObj,
                        new object[] { curveArray, floorType, level, false }) as Floor;
                    if (result != null) return result;
                }
            }
            catch { }

            throw new InvalidOperationException(
                "Không thể tạo Floor. Boundary có thể không hợp lệ — " +
                "kiểm tra đường biên nguồn có đóng kín và không tự giao cắt.");
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
            return CreateTopoSolid(doc, boundary, levelId, null, null);
        }

        /// <summary>
        /// Create a Toposolid with optional interior shape points (Revit 2024+).
        /// If interiorPoints is provided, tries to use Toposolid.Create(doc, boundary, points, typeId, levelId)
        /// which directly defines the surface shape — more reliable than using SlabShapeEditor afterward.
        /// 
        /// surfaceElevation: the Z to flatten boundary curves to. If null, uses level elevation.
        /// CRITICAL for Floor/Roof→Topo conversion when source has Height Offset:
        ///   TopoSolid has no offset parameter, so boundary must be at sourceBaseZ.
        /// </summary>
        public static Toposolid CreateTopoSolid(Document doc, IList<CurveLoop> boundary, ElementId levelId,
            IList<XYZ> interiorPoints, double? surfaceElevation = null)
        {
            ElementId typeId = FindToposolidTypeId(doc);
            if (typeId == null || typeId == ElementId.InvalidElementId)
                throw new InvalidOperationException(
                    "Không tìm thấy Toposolid Type trong project.\n" +
                    "Hãy đảm bảo project có ít nhất một Toposolid Type.");

            Level level = doc.GetElement(levelId) as Level;
            double levelZ = level?.Elevation ?? 0;
            // Use surfaceElevation if provided (bakes source offset into boundary)
            double flatZ = surfaceElevation ?? levelZ;
            IList<CurveLoop> flatBoundary = FlattenToZ(boundary, flatZ);

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
        /// Calculate signed area of a CurveLoop using the Shoelace formula.
        /// Positive = counter-clockwise, Negative = clockwise.
        /// Used to identify the outer loop (largest area) and check winding.
        /// </summary>
        private static double GetLoopArea(CurveLoop loop)
        {
            double area = 0;
            foreach (Curve curve in loop)
            {
                XYZ p0 = curve.GetEndPoint(0);
                XYZ p1 = curve.GetEndPoint(1);
                area += (p0.X * p1.Y - p1.X * p0.Y);
            }
            return area / 2.0;
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
        /// Find a regular Floor type, excluding Structural Foundation types.
        /// Foundation Slab (OST_StructuralFoundation) does NOT support
        /// SlabShapeEditor — must pick a real Floor (OST_Floors).
        /// Priority: "Generic" > any non-foundation > first available.
        /// </summary>
        public static ElementId FindFloorTypeId(Document doc)
        {
            var allTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(FloorType))
                .Cast<FloorType>()
                .ToList();

            if (allTypes.Count == 0)
                return ElementId.InvalidElementId;

            // Separate regular floors from structural foundations
            var regularFloors = new List<FloorType>();
            var foundations = new List<FloorType>();

            foreach (var ft in allTypes)
            {
                // Check the category of the FloorType
                // Structural Foundation types have category OST_StructuralFoundation
                bool isFoundation = false;
                try
                {
                    if (ft.Category != null &&
                        ft.Category.Id.Equals(
                            new ElementId(BuiltInCategory.OST_StructuralFoundation)))
                    {
                        isFoundation = true;
                    }
                }
                catch { }

                // Also check by name keywords as fallback
                if (!isFoundation)
                {
                    string name = ft.Name ?? "";
                    string famName = "";
                    try { famName = ft.FamilyName ?? ""; } catch { }
                    string combined = name + " " + famName;
                    if (combined.IndexOf("Foundation", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        combined.IndexOf("Structural", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        isFoundation = true;
                    }
                }

                if (isFoundation)
                    foundations.Add(ft);
                else
                    regularFloors.Add(ft);
            }

            // Use regular floors if available
            var candidates = regularFloors.Count > 0 ? regularFloors : allTypes;

            // 1. Prefer "Generic" type
            var generic = candidates.FirstOrDefault(t =>
                t.Name.IndexOf("Generic", StringComparison.OrdinalIgnoreCase) >= 0);
            if (generic != null) return generic.Id;

            // 2. First regular floor
            return candidates[0].Id;
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
        /// Get the thickness of an element from its type's CompoundStructure.
        /// Works for Floor, Roof, and similar layered elements.
        /// Returns 0 if thickness cannot be determined.
        /// </summary>
        public static double GetElementThickness(Document doc, Element element)
        {
            if (element == null) return 0;
            try
            {
                ElementId typeId = element.GetTypeId();
                if (typeId == null || typeId == ElementId.InvalidElementId) return 0;

                Element typeElem = doc.GetElement(typeId);

                // Try RoofType
                RoofType roofType = typeElem as RoofType;
                if (roofType != null)
                {
                    CompoundStructure cs = roofType.GetCompoundStructure();
                    if (cs != null) return cs.GetWidth();
                }

                // Try FloorType
                FloorType floorType = typeElem as FloorType;
                if (floorType != null)
                {
                    CompoundStructure cs = floorType.GetCompoundStructure();
                    if (cs != null) return cs.GetWidth();
                }
            }
            catch { }

            // Fallback: try ROOF/FLOOR_ATTR_THICKNESS_PARAM
            try
            {
                Parameter p = element.get_Parameter(BuiltInParameter.ROOF_ATTR_THICKNESS_PARAM);
                if (p != null && p.HasValue) return p.AsDouble();
                p = element.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM);
                if (p != null && p.HasValue) return p.AsDouble();
            }
            catch { }

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
        /// The algorithm uses three passes:
        ///   Pass 1: Match source points to existing target vertices by XY
        ///           proximity and modify them using target vertex's EXACT XY.
        ///           If a point fails, it is NOT marked as matched so Pass 2
        ///           can retry with the source's original XY coordinates.
        ///   Pass 2: Retry failed + unmatched source points using their
        ///           original XY coordinates (DrawPoint handles matching).
        ///   Pass 3: Final fallback — try DrawPoint-first for any remaining
        ///           unapplied points (handles Revit version API differences).
        /// </summary>
        public static int ApplyShapePointsAdjusted(
            Element element, List<ShapePointData> shapeData, double targetBaseZ,
            ConversionDiagnosticInfo diag = null)
        {
            if (shapeData == null || shapeData.Count == 0) return 0;

            SlabShapeEditor editor = GetEditor(element);
            if (editor == null)
            {
                if (diag != null) diag.EditorMethods = "GetEditor returned null";
                return 0;
            }

            if (!editor.IsEnabled)
            {
                editor.Enable();
                try { element.Document.Regenerate(); } catch { }

                editor = GetEditor(element);
                if (editor == null)
                {
                    if (diag != null) diag.EditorMethods = "GetEditor returned null after Enable+Regen";
                    return 0;
                }

                // Retry once if still not enabled
                if (!editor.IsEnabled)
                {
                    try
                    {
                        editor.Enable();
                        element.Document.Regenerate();
                        editor = GetEditor(element);
                    }
                    catch { }
                }

                // On Revit 2026+, IsEnabled may remain False even after
                // Enable() in a separate committed transaction.
                // Do NOT bail out — proceed anyway and try AddPoint/ModifySubElement.
                // These operations may work despite IsEnabled reporting False.
                if (editor == null)
                {
                    if (diag != null) diag.EditorMethods = "GetEditor null after retries";
                    return 0;
                }
            }

            // Regenerate to ensure vertices are fully initialized
            try { element.Document.Regenerate(); } catch { }
            editor = GetEditor(element);
            if (editor == null)
            {
                if (diag != null) diag.EditorMethods = "GetEditor null after final regen";
                return 0;
            }

            // ── Discover available methods for diagnostics ────────────────
            if (diag != null)
            {
                diag.EditorMethods = GetEditorMethodsSummary(editor);
            }

            int applied = 0;
            _lastPointError = null;

            // ── Helper: collect vertices from editor ─────────────────────
            Func<List<SlabShapeVertex>> collectVertices = () =>
            {
                var verts = new List<SlabShapeVertex>();
                try
                {
                    foreach (SlabShapeVertex v in editor.SlabShapeVertices)
                        verts.Add(v);
                }
                catch { }
                return verts;
            };

            // ── Helper: match and modify vertices against source data ────
            // Returns number of newly applied points.
            Func<List<SlabShapeVertex>, double, int> matchAndModify = (verts, tolerance) =>
            {
                int count = 0;
                foreach (var vtx in verts)
                {
                    XYZ vPos = vtx.Position;
                    int bestSrcIdx = -1;
                    double bestDist = tolerance;

                    for (int si = 0; si < shapeData.Count; si++)
                    {
                        if (shapeData[si].Applied) continue;
                        double dx = vPos.X - shapeData[si].X;
                        double dy = vPos.Y - shapeData[si].Y;
                        double dist = Math.Sqrt(dx * dx + dy * dy);
                        if (dist < bestDist) { bestDist = dist; bestSrcIdx = si; }
                    }

                    if (bestSrcIdx < 0) continue;

                    var data = shapeData[bestSrcIdx];
                    double offset = data.RelativeZ;
                    data.TargetZ = targetBaseZ + offset;

                    if (Math.Abs(offset) < 1e-9)
                    {
                        data.Applied = true;
                        count++;
                        continue;
                    }

                    if (ModifyVertexSafe(editor, vtx, offset))
                    {
                        data.Applied = true;
                        count++;
                    }
                    else
                    {
                        data.ErrorInfo = _lastPointError;
                    }
                }
                return count;
            };

            // ── Pass 1: Modify existing vertices (corners) ──────────────
            var existingVerts = collectVertices();
            applied += matchAndModify(existingVerts, 0.5);

            // ── Pass 1b: Regenerate and match NEW auto-generated vertices ─
            // After modifying corner vertices, Revit may auto-create edge
            // vertices along the boundary. Regenerate + re-collect to match
            // source edge points to these new vertices.
            if (applied > 0 && applied < shapeData.Count)
            {
                try { element.Document.Regenerate(); } catch { }
                editor = GetEditor(element);
                if (editor != null)
                {
                    var newVerts = collectVertices();
                    if (newVerts.Count > existingVerts.Count)
                    {
                        applied += matchAndModify(newVerts, 0.5);
                    }
                }
            }

            // ── Pass 2: Add remaining points via DrawPoint/AddPoint ─────
            // For points with no matching vertex (interior or edge points
            // not auto-generated), create them via DrawPoint/AddPoint.
            if (editor == null) editor = GetEditor(element);
            for (int si = 0; si < shapeData.Count; si++)
            {
                if (shapeData[si].Applied) continue;
                var data = shapeData[si];
                double offset = data.RelativeZ;
                data.TargetZ = targetBaseZ + offset;

                try
                {
                    XYZ facePoint = new XYZ(data.X, data.Y, targetBaseZ);
                    if (AddPointSafe(editor, facePoint))
                    {
                        data.Applied = true;
                        applied++;

                        if (Math.Abs(offset) > 1e-9)
                        {
                            try { element.Document.Regenerate(); } catch { }
                            editor = GetEditor(element);
                            if (editor != null)
                            {
                                foreach (SlabShapeVertex v in editor.SlabShapeVertices)
                                {
                                    double vdx = v.Position.X - data.X;
                                    double vdy = v.Position.Y - data.Y;
                                    if (Math.Sqrt(vdx * vdx + vdy * vdy) < 0.01)
                                    {
                                        ModifyVertexSafe(editor, v, offset);
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        data.ErrorInfo = _lastPointError;
                    }
                }
                catch { }
            }

#if REVIT2026
            // ── Revit 2026 Fallback: Try AddPoint with absolute Z ───────
            // On Revit 2026, SlabShapeEditor.Enable() may not work for
            // newly created elements. As a fallback, try AddPoint with
            // the ABSOLUTE desired Z coordinate (not face Z).
            // Revit 2026's AddPoint may accept absolute coordinates and
            // compute the offset internally, unlike old DrawPoint which
            // required the point to be ON the slab face.
            if (applied == 0)
            {
                // Re-acquire editor in case previous attempts changed state
                editor = GetEditor(element);
                if (editor != null)
                {
                    // Try Enable one more time before the fallback
                    if (!editor.IsEnabled)
                    {
                        try { editor.Enable(); } catch { }
                        try { element.Document.Regenerate(); } catch { }
                        editor = GetEditor(element);
                    }

                    if (editor != null)
                    {
                        for (int si = 0; si < shapeData.Count; si++)
                        {
                            if (shapeData[si].Applied) continue;
                            var data = shapeData[si];
                            double desiredZ = targetBaseZ + data.RelativeZ;
                            data.TargetZ = desiredZ;

                            // Strategy A: AddPoint with absolute Z
                            try
                            {
                                XYZ absPoint = new XYZ(data.X, data.Y, desiredZ);
                                editor.AddPoint(absPoint);
                                data.Applied = true;
                                applied++;
                                continue;
                            }
                            catch (Exception ex)
                            {
                                data.ErrorInfo = $"AddPoint(absZ): {ex.GetType().Name}: {ex.Message}";
                            }

                            // Strategy B: AddPoint with offset as Z
                            try
                            {
                                XYZ offsetPoint = new XYZ(data.X, data.Y, data.RelativeZ);
                                editor.AddPoint(offsetPoint);
                                data.Applied = true;
                                applied++;
                                continue;
                            }
                            catch (Exception ex)
                            {
                                data.ErrorInfo = $"AddPoint(offsetZ): {ex.GetType().Name}: {ex.Message}";
                            }
                        }
                    }
                }
            }
#endif

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

                // ── 2b. Subdivide boundary at edge vertex positions ─────
                // Edge vertices lie ON boundary edges. SlabShapeEditor cannot
                // add new points on edges — only modify existing vertices.
                // By splitting boundary curves at edge positions, each edge
                // point becomes a corner vertex on the target element.
                if (shapeData.Count > 0 && target != ConvertTarget.TopoSolid)
                {
                    try
                    {
                        boundary = SubdivideBoundaryAtEdgePoints(boundary, shapeData);
                    }
                    catch { /* Keep original boundary if subdivision fails */ }
                }

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
                            // For TopoSolid: pass shape points directly during creation.
                            // CRITICAL: TopoSolid does NOT have "Height Offset From Level"
                            // parameter, so we must bake the source offset into:
                            //   1. Boundary Z (surfaceElevation = sourceBaseZ)
                            //   2. Interior point Z (adjZ = sourceBaseZ + relativeZ)
                            var topoPoints = new List<XYZ>();
                            double topoBaseZ = diag.SourceBaseZ;
                            foreach (var sp in shapeData)
                            {
                                double adjZ = topoBaseZ + sp.RelativeZ;
                                topoPoints.Add(new XYZ(sp.X, sp.Y, adjZ));
                            }

                            // If no shape points but source has offset,
                            // generate a centroid point at sourceBaseZ.
                            // This forces Toposolid.Create to respect the
                            // surface elevation (some Revit versions ignore boundary Z).
                            if (topoPoints.Count == 0 && Math.Abs(diag.SourceOffset) > 1e-9)
                            {
                                double cx = 0, cy = 0;
                                int ptCount = 0;
                                foreach (var loop in boundary)
                                {
                                    foreach (Curve c in loop)
                                    {
                                        XYZ p = c.GetEndPoint(0);
                                        cx += p.X;
                                        cy += p.Y;
                                        ptCount++;
                                    }
                                }
                                if (ptCount > 0)
                                {
                                    cx /= ptCount;
                                    cy /= ptCount;
                                    topoPoints.Add(new XYZ(cx, cy, topoBaseZ));
                                }
                            }

                            // Pass sourceBaseZ as surfaceElevation so boundary is
                            // flattened to the correct absolute Z (not just levelZ)
                            newElement = CreateTopoSolid(doc, boundary, levelId,
                                topoPoints.Count > 0 ? topoPoints : null,
                                diag.SourceBaseZ);
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

                // ── 5b. For TopoSolid target: override base Z ────────────
                // TopoSolid has NO "Height Offset From Level" parameter,
                // so SetHeightOffset (step 7) will do nothing.
                // We must bake the source offset into the Z math here:
                //   adjustedZ = sourceBaseZ + relativeZ = original absolute Z
                // For Floor/Roof targets, keep using targetBaseZ because
                // step 7 will apply the offset via the parameter.
                double shapeBaseZ = targetBaseZ;
#if REVIT2024_PLUS
                if (target == ConvertTarget.TopoSolid)
                {
                    shapeBaseZ = diag.SourceBaseZ;
                }
#endif

                // ── 6. Apply shape points with Z adjustment ─────────────
                // Try applying within this transaction first.
                // If Enable fails (Revit 2026+), store data for phase 2.
                if (shapeData.Count > 0)
                {
                    try
                    {
                        int applied = ApplyShapePointsAdjusted(
                            newElement, shapeData, shapeBaseZ, diag);
                        diag.AppliedCount = applied;

                        if (applied < shapeData.Count)
                        {
                            // Check if editor couldn't be enabled (Revit 2026+)
                            // If so, defer to phase 2 (separate transaction)
                            bool editorFailed = diag.EditorMethods != null &&
                                diag.EditorMethods.Contains("Enabled=False");

                            if (editorFailed && applied == 0)
                            {
                                // Store for phase 2 — separate transaction
                                // Reset Applied flags for retry
                                foreach (var sp in shapeData)
                                {
                                    sp.Applied = false;
                                    sp.ErrorInfo = null;
                                }
                                result.HasPendingShapeEditing = true;
                                result.PendingShapeData = shapeData;
                                result.PendingShapeBaseZ = shapeBaseZ;
                            }
                            else
                            {
                                var errors = shapeData
                                    .Where(sp => !sp.Applied && !string.IsNullOrEmpty(sp.ErrorInfo))
                                    .Select(sp => sp.ErrorInfo)
                                    .Distinct()
                                    .Take(3)
                                    .ToList();
                                string errDetail = errors.Count > 0
                                    ? " | " + string.Join("; ", errors)
                                    : "";
                                result.WarningMessage =
                                    $"Đã áp dụng {applied}/{shapeData.Count} điểm nhấc{errDetail}";
                            }
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
                // NOTE: Skip for TopoSolid — it has no offset parameter,
                // and the offset was already baked into shape points (step 5b/6).
                if (target != ConvertTarget.TopoSolid)
                {
                    double finalOffset = diag.SourceOffset;

                    // ── Roof thickness compensation ─────────────────────
                    // Roof reference plane is at the BOTTOM. Top = bottom + thickness.
                    // To align roof TOP surface with source surface (TopoSolid/Floor),
                    // we must shift the roof DOWN by its thickness.
                    // TopoSolid → Roof: offset = 0 - thickness = -thickness
                    // Floor → Roof:     offset = sourceOffset - thickness
                    if (target == ConvertTarget.Roof)
                    {
                        double roofThickness = GetElementThickness(doc, newElement);
                        if (roofThickness > 1e-9)
                        {
                            finalOffset -= roofThickness;
                        }
                    }

                    if (Math.Abs(finalOffset) > 1e-9)
                    {
                        SetHeightOffset(newElement, finalOffset);
                    }
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

        /// <summary>
        /// Enable SlabShapeEditor on an element.
        /// Must be called inside its own Transaction and committed
        /// BEFORE attempting to modify vertices (Revit 2026+ requirement).
        /// </summary>
        public static void EnableShapeEditor(Element element)
        {
            if (element == null) return;

            SlabShapeEditor editor = GetEditor(element);
            if (editor == null) return;

            if (!editor.IsEnabled)
            {
                editor.Enable();
            }
        }

        /// <summary>
        /// Apply pending shape points AFTER the editor has been enabled
        /// in a separate committed transaction.
        /// Called in Phase 3 (Revit 2026+) or Phase 2 (older versions).
        /// Must be called inside a Transaction.
        /// </summary>
        public static void ApplyPendingShapeEditing(Document doc, ConversionResult result)
        {
            if (!result.HasPendingShapeEditing) return;
            if (result.PendingShapeData == null || result.PendingShapeData.Count == 0) return;
            if (result.NewElementId == null || result.NewElementId == ElementId.InvalidElementId) return;

            Element element = doc.GetElement(result.NewElementId);
            if (element == null) return;

            var diag = result.Diagnostics;

            try
            {
                try { doc.Regenerate(); } catch { }

                int applied = ApplyShapePointsAdjusted(
                    element, result.PendingShapeData, result.PendingShapeBaseZ, diag);

                if (diag != null)
                    diag.AppliedCount = applied;

                if (applied < result.PendingShapeData.Count)
                {
                    var errors = result.PendingShapeData
                        .Where(sp => !sp.Applied && !string.IsNullOrEmpty(sp.ErrorInfo))
                        .Select(sp => sp.ErrorInfo)
                        .Distinct()
                        .Take(3)
                        .ToList();
                    string errDetail = errors.Count > 0
                        ? " | " + string.Join("; ", errors)
                        : "";
                    result.WarningMessage =
                        $"Phase 2: {applied}/{result.PendingShapeData.Count} điểm nhấc{errDetail}";
                }
                else
                {
                    // Clear warning — all points applied successfully
                    result.WarningMessage = null;
                }
            }
            catch (Exception ex)
            {
                result.WarningMessage = $"Phase 2 error: {ex.Message}";
            }

            result.HasPendingShapeEditing = false;
        }
    }
}
