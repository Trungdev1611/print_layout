using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace PrintLayoutAddin.Core
{
    public class SttAssignResult
    {
        public int Assigned;
        public int VertexMissed;
        public int FramesSkipped;
        /// <summary>Number of codes in the supplied list that were never applied.</summary>
        public int CodesUnused;
        public int DrawingNamesAssigned;
        public int DrawingNameAttributesMissing;
        /// <summary>True when no drawing change was committed (validation abort).</summary>
        public bool Aborted;
        /// <summary>Short human-readable explanation when <see cref="Aborted"/> is true.</summary>
        public string Message;
    }

    public static class SttAssigner
    {
        // ------------------------------------------------------------
        // Walk the guide polyline's vertices in order. For each vertex
        // landing inside a not-yet-numbered frame, take the next code
        // from <paramref name="codes"/> and write it into the frame's
        // INNO-STT attribute (with XData fallback).
        //
        // If the number of codes and the number of matched vertices do
        // not agree:
        //   - when <paramref name="allowMismatch"/> is false (default),
        //     the transaction is rolled back and no frame is modified;
        //   - when true, as many codes as possible are applied and the
        //     leftover frames/codes are reported via the result.
        // ------------------------------------------------------------
        public static SttAssignResult Run(
            Database db,
            Editor ed,
            ObjectId polyId,
            string blockName,
            List<string> codes,
            List<string> drawingNames = null,
            bool allowMismatch = false)
        {
            var cfg = Config.Instance;
            var result = new SttAssignResult();
            if (codes == null) codes = new List<string>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var pts = ReadPolylinePoints(tr, polyId);

                var ms = (BlockTableRecord)tr.GetObject(
                    SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

                var frames = new List<(ObjectId Id, Point3d Min, Point3d Max, bool Used)>();
                foreach (ObjectId id in ms)
                {
                    if (id.ObjectClass.DxfName != "INSERT") continue;
                    var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                    if (br == null) continue;
                    if (!string.Equals(br.Name, blockName, System.StringComparison.OrdinalIgnoreCase)) continue;

                    Extents3d ext;
                    try { ext = br.GeometricExtents; } catch { continue; }
                    frames.Add((id, ext.MinPoint, ext.MaxPoint, false));
                }

                // PASS 1 — walk vertices, record which frame each match should take
                // (in the order the polyline traverses them). No writes yet.
                var pending = new List<ObjectId>();
                int vertexMissed = 0;
                var used = new HashSet<ObjectId>();

                for (int i = 0; i < pts.Count; i++)
                {
                    var p = pts[i];
                    ObjectId? matchId = null;
                    for (int j = 0; j < frames.Count; j++)
                    {
                        var f = frames[j];
                        if (used.Contains(f.Id)) continue;
                        if (Contains(f.Min, f.Max, p)) { matchId = f.Id; break; }
                    }
                    if (!matchId.HasValue) { vertexMissed++; continue; }
                    used.Add(matchId.Value);
                    pending.Add(matchId.Value);
                }

                result.VertexMissed = vertexMissed;

                // Mismatch check — bail out without touching anything if the caller
                // hasn't opted into partial assignment.
                if (!allowMismatch && pending.Count != codes.Count)
                {
                    result.Aborted = true;
                    result.Message =
                        $"Count mismatch: polyline matches {pending.Count} frame(s), " +
                        $"but {codes.Count} code(s) were supplied. No frame was modified.";
                    // tr.Dispose without commit = rollback
                    return result;
                }

                // PASS 2 — apply. Walk the shorter of the two lists.
                int n = System.Math.Min(pending.Count, codes.Count);
                for (int i = 0; i < n; i++)
                {
                    var br = (BlockReference)tr.GetObject(pending[i], OpenMode.ForWrite);
                    FrameScanner.WriteStt(br, codes[i], cfg, tr);
                    if (drawingNames != null && i < drawingNames.Count
                        && !string.IsNullOrWhiteSpace(drawingNames[i]))
                    {
                        if (FrameScanner.WriteDrawingName(br, drawingNames[i], cfg, tr))
                            result.DrawingNamesAssigned++;
                        else
                            result.DrawingNameAttributesMissing++;
                    }
                }
                result.Assigned = n;
                result.FramesSkipped = frames.Count - n; // all non-written frames
                result.CodesUnused = codes.Count - n;

                tr.Commit();
            }
            return result;
        }

        // ------------------------------------------------------------
        // Public alias matching the naming in the upgrade spec.
        // ------------------------------------------------------------
        public static SttAssignResult ApplyNumbersToSelectedFrames(
            Database db,
            Editor ed,
            ObjectId polyId,
            string blockName,
            List<string> codes,
            List<string> drawingNames,
            bool allowMismatch = false)
            => Run(db, ed, polyId, blockName, codes, drawingNames, allowMismatch);

        public static SttAssignResult ApplyNumbersToSelectedFrames(
            Database db,
            Editor ed,
            ObjectId polyId,
            string blockName,
            List<string> codes,
            bool allowMismatch = false)
            => Run(db, ed, polyId, blockName, codes, null, allowMismatch);

        // ------------------------------------------------------------
        // Helper: count how many frames of <paramref name="blockName"/>
        // the polyline would hit, without writing anything. Used by the
        // UI to pre-seed the "Count" / "End" fields in the dialog.
        // ------------------------------------------------------------
        public static int CountMatches(Database db, ObjectId polyId, string blockName)
        {
            int hit = 0;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var pts = ReadPolylinePoints(tr, polyId);
                var ms = (BlockTableRecord)tr.GetObject(
                    SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

                var frames = new List<(ObjectId Id, Point3d Min, Point3d Max)>();
                foreach (ObjectId id in ms)
                {
                    if (id.ObjectClass.DxfName != "INSERT") continue;
                    var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                    if (br == null) continue;
                    if (!string.Equals(br.Name, blockName, System.StringComparison.OrdinalIgnoreCase)) continue;
                    Extents3d ext;
                    try { ext = br.GeometricExtents; } catch { continue; }
                    frames.Add((id, ext.MinPoint, ext.MaxPoint));
                }

                var used = new HashSet<ObjectId>();
                foreach (var p in pts)
                {
                    for (int j = 0; j < frames.Count; j++)
                    {
                        if (used.Contains(frames[j].Id)) continue;
                        if (Contains(frames[j].Min, frames[j].Max, p))
                        {
                            used.Add(frames[j].Id);
                            hit++;
                            break;
                        }
                    }
                }
                tr.Commit();
            }
            return hit;
        }

        private static List<Point3d> ReadPolylinePoints(Transaction tr, ObjectId id)
        {
            var pts = new List<Point3d>();
            var obj = tr.GetObject(id, OpenMode.ForRead);
            switch (obj)
            {
                case Polyline pl:
                    for (int i = 0; i < pl.NumberOfVertices; i++) pts.Add(pl.GetPoint3dAt(i));
                    break;
                case Polyline2d p2:
                    foreach (ObjectId vid in p2)
                    {
                        var v = tr.GetObject(vid, OpenMode.ForRead) as Vertex2d;
                        if (v != null) pts.Add(v.Position);
                    }
                    break;
                case Polyline3d p3:
                    foreach (ObjectId vid in p3)
                    {
                        var v = tr.GetObject(vid, OpenMode.ForRead) as PolylineVertex3d;
                        if (v != null) pts.Add(v.Position);
                    }
                    break;
            }
            return pts;
        }

        private static bool Contains(Point3d min, Point3d max, Point3d p)
        {
            const double eps = 1e-6;
            return p.X >= min.X - eps && p.X <= max.X + eps
                && p.Y >= min.Y - eps && p.Y <= max.Y + eps;
        }
    }
}
