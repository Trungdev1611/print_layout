using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace PrintLayoutAddin.Core
{
    /// <summary>
    /// One detected nested frame instance.
    /// Contract for <see cref="NativeFrameBuilder.InsertFrames"/>:
    ///   xform = ParentToWcs * Displacement(LocalMin)
    /// where LocalMin/Max are the frame AABB in the coordinate system of the
    /// containing BTR, and ParentToWcs maps that space to WCS (composed through
    /// any nesting depth).
    /// </summary>
    public class NestedFrameHit
    {
        public ObjectId ParentRefId;
        public ObjectId NestedRefId;

        /// <summary>Containing-BTR space → WCS (product of ancestor BlockTransforms).</summary>
        public Matrix3d ParentToWcs;

        /// <summary>Frame AABB in containing-BTR coordinates.</summary>
        public Point3d LocalMin;
        public Point3d LocalMax;

        public double LocalWidth => LocalMax.X - LocalMin.X;
        public double LocalHeight => LocalMax.Y - LocalMin.Y;
    }

    public static class NestedFrameScanner
    {
        /// <summary>Skip tiny junk geometry when ranking picker candidates.</summary>
        public const double MinCandidateSize = 50.0;

        /// <summary>Safety cap on nesting depth (path length).</summary>
        public const int MaxWalkDepth = 16;

        private static readonly string[] PreferredNameParts =
        {
            "khung", "frame", "title", "a0", "a1", "a2", "a3", "a4", "khổ", "kho"
        };

        /// <summary>
        /// Skip PLAUTO wrappers, anonymous/dynamic effective names, and xref temp names
        /// from picker lists (e.g. PL_*, A$*, *..., Name|A$..., Name|Name).
        /// </summary>
        public static bool IsNoiseBlockName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return true;
            var n = name.Trim();
            if (n.StartsWith("PL_", StringComparison.OrdinalIgnoreCase)) return true;
            if (n.StartsWith("A$", StringComparison.OrdinalIgnoreCase)) return true;
            if (n.StartsWith("*", StringComparison.Ordinal)) return true;
            // Dynamic / nested effective names: "D2.Khung TB|A$..." or "D2.Khung TB|D2.Khung TB"
            if (n.IndexOf('|') >= 0) return true;
            if (n.IndexOf("A$", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            return false;
        }

        /// <summary>
        /// Auto-scan: every nested block name under all ModelSpace inserts (any depth),
        /// excluding noise names (PL_ / A$ / *). Sorted for the single PLAUTO picker.
        /// </summary>
        public static List<BlockChoice> ListFramesInModelSpace(Database db)
        {
            var result = new Dictionary<string, BlockChoice>(StringComparer.OrdinalIgnoreCase);
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var ms = (BlockTableRecord)tr.GetObject(
                    SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

                foreach (ObjectId pid in ms)
                {
                    if (idIsNotInsert(pid)) continue;
                    BlockReference br;
                    try { br = tr.GetObject(pid, OpenMode.ForRead, false) as BlockReference; }
                    catch { continue; }
                    if (br == null) continue;

                    string topName = br.Name ?? "";
                    // Do not walk into existing PL_ wrappers (avoid listing garbage / double-wrap).
                    if (topName.StartsWith("PL_", StringComparison.OrdinalIgnoreCase))
                        continue;

                    BlockTableRecord topBtr;
                    try { topBtr = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead); }
                    catch { continue; }
                    if (!CanRecurseInto(topBtr)) continue;

                    // Direct MS insert can itself be a frame candidate (non-noise).
                    if (!IsNoiseBlockName(topName))
                        RecordChoice(result, topName, topBtr, br, depth: 0);

                    var path = new HashSet<ObjectId> { topBtr.ObjectId };
                    WalkList(tr, topBtr, path, depth: 1, result);
                }

                tr.Commit();
            }

            return FilterAndSortCandidates(result.Values);
        }

        /// <summary>
        /// All instances of <paramref name="frameBlockName"/> anywhere under ModelSpace
        /// (direct MS inserts + nested at any depth). Skips PL_ wrapper roots.
        /// </summary>
        public static List<NestedFrameHit> CollectHitsInModelSpace(Database db, string frameBlockName)
        {
            var hits = new List<NestedFrameHit>();
            if (string.IsNullOrWhiteSpace(frameBlockName) || IsNoiseBlockName(frameBlockName))
                return hits;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                // Shared BTR-local extents for direct MS hits of this frame.
                Extents3d? btrLocalExt = null;
                try
                {
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    if (bt.Has(frameBlockName))
                    {
                        var def = (BlockTableRecord)tr.GetObject(bt[frameBlockName], OpenMode.ForRead);
                        foreach (ObjectId eid in def)
                        {
                            var ent = tr.GetObject(eid, OpenMode.ForRead, false) as Entity;
                            if (ent == null) continue;
                            Extents3d ext;
                            try { ext = ent.GeometricExtents; } catch { continue; }
                            if (btrLocalExt == null) btrLocalExt = ext;
                            else { var t = btrLocalExt.Value; t.AddExtents(ext); btrLocalExt = t; }
                        }
                    }
                }
                catch { }

                var ms = (BlockTableRecord)tr.GetObject(
                    SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

                foreach (ObjectId pid in ms)
                {
                    if (idIsNotInsert(pid)) continue;
                    BlockReference br;
                    try { br = tr.GetObject(pid, OpenMode.ForRead, false) as BlockReference; }
                    catch { continue; }
                    if (br == null) continue;

                    string topName = br.Name ?? "";
                    if (topName.StartsWith("PL_", StringComparison.OrdinalIgnoreCase))
                        continue;

                    // Frame sitting directly in ModelSpace.
                    if (string.Equals(topName, frameBlockName, StringComparison.OrdinalIgnoreCase)
                        && btrLocalExt.HasValue)
                    {
                        hits.Add(new NestedFrameHit
                        {
                            ParentRefId = pid,
                            NestedRefId = pid,
                            ParentToWcs = br.BlockTransform,
                            LocalMin = btrLocalExt.Value.MinPoint,
                            LocalMax = btrLocalExt.Value.MaxPoint,
                        });
                    }

                    BlockTableRecord topBtr;
                    try { topBtr = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead); }
                    catch { continue; }
                    if (!CanRecurseInto(topBtr)) continue;

                    var path = new HashSet<ObjectId> { topBtr.ObjectId };
                    WalkCollect(
                        tr,
                        topBtr,
                        br.BlockTransform,
                        pid,
                        frameBlockName,
                        path,
                        depth: 1,
                        hits);
                }

                tr.Commit();
            }

            return hits;
        }

        /// <summary>
        /// Distinct block/xref names nested under <paramref name="parentBlockName"/>
        /// at any depth (recursive). Kept for compatibility; PLAUTO uses
        /// <see cref="ListFramesInModelSpace"/>.
        /// </summary>
        public static List<BlockChoice> ListNestedBlocks(Database db, string parentBlockName)
        {
            var result = new Dictionary<string, BlockChoice>(StringComparer.OrdinalIgnoreCase);
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (!bt.Has(parentBlockName))
                {
                    tr.Commit();
                    return new List<BlockChoice>();
                }

                var parent = (BlockTableRecord)tr.GetObject(bt[parentBlockName], OpenMode.ForRead);
                var path = new HashSet<ObjectId>();
                path.Add(parent.ObjectId);
                WalkList(tr, parent, path, depth: 1, result);
                tr.Commit();
            }

            return FilterAndSortCandidates(result.Values);
        }

        /// <summary>
        /// Rank / soft-filter candidates: drop noise + tiny sizes, prefer frame-like names,
        /// sort by count descending then name.
        /// </summary>
        public static List<BlockChoice> FilterAndSortCandidates(IEnumerable<BlockChoice> raw)
        {
            var list = (raw ?? Enumerable.Empty<BlockChoice>())
                .Where(c => c != null && !IsNoiseBlockName(c.Name))
                .ToList();

            bool AnyPreferred(BlockChoice c)
            {
                var n = c.Name ?? "";
                foreach (var p in PreferredNameParts)
                {
                    if (n.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0)
                        return true;
                }
                return false;
            }

            bool SizeOk(BlockChoice c)
            {
                if (c.Width <= 0 || c.Height <= 0) return true; // unknown — keep
                return c.Width >= MinCandidateSize && c.Height >= MinCandidateSize;
            }

            var sized = list.Where(SizeOk).ToList();
            if (sized.Count == 0) sized = list;

            var preferred = sized.Where(AnyPreferred).ToList();
            var pool = preferred.Count > 0 ? preferred : sized;

            return pool
                .OrderByDescending(c => c.Count)
                .ThenBy(c => c.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        /// <summary>
        /// For every BlockReference of <paramref name="blockName"/> sitting directly in ModelSpace,
        /// build a NestedFrameHit using the BR's own BlockTransform as ParentToWcs and the BTR's
        /// own bounding box as LocalMin/LocalMax.
        /// </summary>
        public static List<NestedFrameHit> CollectDirectHits(Database db, string blockName)
        {
            var hits = new List<NestedFrameHit>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (!bt.Has(blockName))
                {
                    tr.Commit();
                    return hits;
                }
                var btrId = bt[blockName];
                var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

                Extents3d? localExt = null;
                foreach (ObjectId eid in btr)
                {
                    var ent = tr.GetObject(eid, OpenMode.ForRead) as Entity;
                    if (ent == null) continue;
                    Extents3d ext;
                    try { ext = ent.GeometricExtents; } catch { continue; }
                    if (localExt == null) localExt = ext;
                    else { var t = localExt.Value; t.AddExtents(ext); localExt = t; }
                }
                if (!localExt.HasValue)
                {
                    tr.Commit();
                    return hits;
                }
                var localMin = localExt.Value.MinPoint;
                var localMax = localExt.Value.MaxPoint;

                var ms = (BlockTableRecord)tr.GetObject(
                    SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);
                foreach (ObjectId pid in ms)
                {
                    if (idIsNotInsert(pid)) continue;
                    var br = tr.GetObject(pid, OpenMode.ForRead, false) as BlockReference;
                    if (br == null) continue;
                    if (!string.Equals(br.Name, blockName, StringComparison.OrdinalIgnoreCase)) continue;

                    hits.Add(new NestedFrameHit
                    {
                        ParentRefId = pid,
                        NestedRefId = pid,
                        ParentToWcs = br.BlockTransform,
                        LocalMin = localMin,
                        LocalMax = localMax,
                    });
                }
                tr.Commit();
            }
            return hits;
        }

        /// <summary>
        /// Every instance of <paramref name="nestedBlockName"/> under each ModelSpace insert of
        /// <paramref name="parentBlockName"/>, at any nesting depth. Matrix chain matches InsertFrames.
        /// </summary>
        public static List<NestedFrameHit> CollectHits(
            Database db, string parentBlockName, string nestedBlockName)
        {
            var hits = new List<NestedFrameHit>();
            if (string.IsNullOrWhiteSpace(parentBlockName) || string.IsNullOrWhiteSpace(nestedBlockName))
                return hits;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var ms = (BlockTableRecord)tr.GetObject(
                    SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

                foreach (ObjectId pid in ms)
                {
                    if (idIsNotInsert(pid)) continue;
                    BlockReference parentBR;
                    try { parentBR = tr.GetObject(pid, OpenMode.ForRead, false) as BlockReference; }
                    catch { continue; }
                    if (parentBR == null) continue;
                    if (!string.Equals(parentBR.Name, parentBlockName, StringComparison.OrdinalIgnoreCase))
                        continue;

                    BlockTableRecord parentBtr;
                    try
                    {
                        parentBtr = (BlockTableRecord)tr.GetObject(
                            parentBR.BlockTableRecord, OpenMode.ForRead);
                    }
                    catch { continue; }

                    if (!CanRecurseInto(parentBtr)) continue;

                    var path = new HashSet<ObjectId> { parentBtr.ObjectId };
                    WalkCollect(
                        tr,
                        parentBtr,
                        parentBR.BlockTransform,
                        pid,
                        nestedBlockName,
                        path,
                        depth: 1,
                        hits);
                }

                tr.Commit();
            }
            return hits;
        }

        // ------------------------------------------------------------------
        // Recursive walk
        // ------------------------------------------------------------------

        private static void WalkList(
            Transaction tr,
            BlockTableRecord btr,
            HashSet<ObjectId> path,
            int depth,
            Dictionary<string, BlockChoice> result)
        {
            if (btr == null || depth > MaxWalkDepth) return;

            foreach (ObjectId id in btr)
            {
                if (idIsNotInsert(id)) continue;
                BlockReference br;
                try { br = tr.GetObject(id, OpenMode.ForRead, false) as BlockReference; }
                catch { continue; }
                if (br == null) continue;

                BlockTableRecord childBtr = null;
                try { childBtr = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead); }
                catch { continue; }
                if (childBtr == null) continue;

                string name = br.Name ?? "";
                if (string.IsNullOrWhiteSpace(name)) continue;

                // Never offer PL_ / A$ / * as wrap targets; still may walk non-PL_ children.
                bool noise = IsNoiseBlockName(name);
                if (!noise)
                    RecordChoice(result, name, childBtr, br, depth);

                // Do not recurse into PLAUTO wrappers.
                if (name.StartsWith("PL_", StringComparison.OrdinalIgnoreCase))
                    continue;

                if (!CanRecurseInto(childBtr)) continue;
                if (path.Contains(childBtr.ObjectId)) continue; // circular

                path.Add(childBtr.ObjectId);
                try { WalkList(tr, childBtr, path, depth + 1, result); }
                finally { path.Remove(childBtr.ObjectId); }
            }
        }

        private static void RecordChoice(
            Dictionary<string, BlockChoice> result,
            string name,
            BlockTableRecord childBtr,
            BlockReference br,
            int depth)
        {
            if (!result.TryGetValue(name, out var choice))
            {
                choice = new BlockChoice
                {
                    Name = name,
                    IsXref = childBtr != null && childBtr.IsFromExternalReference,
                    Count = 0,
                    Depth = depth,
                };
                result[name] = choice;
            }
            choice.Count++;
            if (depth > choice.Depth) choice.Depth = depth;
            TrySampleSize(br, choice);
        }

        private static void WalkCollect(
            Transaction tr,
            BlockTableRecord btr,
            Matrix3d spaceToWcs,
            ObjectId rootParentId,
            string targetName,
            HashSet<ObjectId> path,
            int depth,
            List<NestedFrameHit> hits)
        {
            if (btr == null || depth > MaxWalkDepth) return;

            foreach (ObjectId nid in btr)
            {
                if (idIsNotInsert(nid)) continue;
                BlockReference nestedBR;
                try { nestedBR = tr.GetObject(nid, OpenMode.ForRead, false) as BlockReference; }
                catch { continue; }
                if (nestedBR == null) continue;

                if (string.Equals(nestedBR.Name, targetName, StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        var ext = nestedBR.GeometricExtents;
                        hits.Add(new NestedFrameHit
                        {
                            ParentRefId = rootParentId,
                            NestedRefId = nid,
                            ParentToWcs = spaceToWcs,
                            LocalMin = ext.MinPoint,
                            LocalMax = ext.MaxPoint,
                        });
                    }
                    catch { /* skip bad extents */ }
                }

                // Do not walk into existing PL_ wrappers.
                if ((nestedBR.Name ?? "").StartsWith("PL_", StringComparison.OrdinalIgnoreCase))
                    continue;

                BlockTableRecord childBtr;
                try { childBtr = (BlockTableRecord)tr.GetObject(nestedBR.BlockTableRecord, OpenMode.ForRead); }
                catch { continue; }
                if (!CanRecurseInto(childBtr)) continue;
                if (path.Contains(childBtr.ObjectId)) continue;

                Matrix3d childSpace;
                try { childSpace = spaceToWcs * nestedBR.BlockTransform; }
                catch { continue; }

                path.Add(childBtr.ObjectId);
                try
                {
                    WalkCollect(
                        tr, childBtr, childSpace, rootParentId, targetName, path, depth + 1, hits);
                }
                finally { path.Remove(childBtr.ObjectId); }
            }
        }

        private static void TrySampleSize(BlockReference br, BlockChoice choice)
        {
            if (choice.Width > 0 && choice.Height > 0) return;
            try
            {
                var ext = br.GeometricExtents;
                double w = ext.MaxPoint.X - ext.MinPoint.X;
                double h = ext.MaxPoint.Y - ext.MinPoint.Y;
                if (w > 0 && h > 0)
                {
                    choice.Width = w;
                    choice.Height = h;
                }
            }
            catch { }
        }

        private static bool CanRecurseInto(BlockTableRecord btr)
        {
            if (btr == null) return false;
            try
            {
                if (btr.IsLayout) return false;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool idIsNotInsert(ObjectId id)
        {
            try
            {
                if (id.IsNull || id.IsErased) return true;
                return id.ObjectClass.DxfName != "INSERT";
            }
            catch
            {
                return true;
            }
        }
    }
}
