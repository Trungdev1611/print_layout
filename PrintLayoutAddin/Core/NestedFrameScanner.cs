using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace PrintLayoutAddin.Core
{
    /// <summary>
    /// One detected nested frame instance. All coordinates are in WCS unless noted.
    /// </summary>
    public class NestedFrameHit
    {
        public ObjectId ParentRefId;
        public ObjectId NestedRefId;

        // Transform taking points from nestedBR's parent BTR coord system into WCS.
        public Matrix3d ParentToWcs;

        // Nested BR's extents — expressed in nestedBR's parent BTR coord system.
        public Point3d LocalMin;
        public Point3d LocalMax;

        public double LocalWidth => LocalMax.X - LocalMin.X;
        public double LocalHeight => LocalMax.Y - LocalMin.Y;
    }

    public static class NestedFrameScanner
    {
        /// <summary>
        /// List every distinct block/xref reference name that sits INSIDE the given parent block's BTR.
        /// Used to let the user pick which nested block is the frame (e.g. "Z1.Khung TB").
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
                    return result.Values.ToList();
                }
                var parent = (BlockTableRecord)tr.GetObject(bt[parentBlockName], OpenMode.ForRead);
                foreach (ObjectId id in parent)
                {
                    if (id.ObjectClass.DxfName != "INSERT") continue;
                    var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                    if (br == null) continue;
                    var btr = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                    var name = br.Name;
                    if (!result.TryGetValue(name, out var choice))
                    {
                        choice = new BlockChoice
                        {
                            Name = name,
                            IsXref = btr.IsFromExternalReference,
                            Count = 0
                        };
                        result[name] = choice;
                    }
                    choice.Count++;
                }
                tr.Commit();
            }
            return result.Values.OrderBy(c => c.Name).ToList();
        }

        /// <summary>
        /// For every BlockReference of <paramref name="blockName"/> sitting directly in ModelSpace,
        /// build a NestedFrameHit using the BR's own BlockTransform as ParentToWcs and the BTR's
        /// own bounding box (computed once from the BTR's entities) as LocalMin/LocalMax.
        ///
        /// Used by PLAUTO when the user wants to wrap a block that lives in the current DWG, not
        /// nested inside an xref. Downstream NativeFrameBuilder.InsertFrames uses the same
        /// formula (xform = ParentToWcs * T(LocalMin)) so PL_ origin lines up with the source
        /// block's local min corner in WCS — identical to the nested case.
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

                // BTR-local extents: aggregate entity extents in BTR coord system.
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
                    if (pid.ObjectClass.DxfName != "INSERT") continue;
                    var br = tr.GetObject(pid, OpenMode.ForRead) as BlockReference;
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
        /// For every BlockReference of parentBlockName in ModelSpace, find every nested BlockReference of
        /// nestedBlockName inside it. Return one NestedFrameHit per nested instance, with composed transform.
        /// </summary>
        public static List<NestedFrameHit> CollectHits(Database db, string parentBlockName, string nestedBlockName)
        {
            var hits = new List<NestedFrameHit>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var ms = (BlockTableRecord)tr.GetObject(
                    SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

                foreach (ObjectId pid in ms)
                {
                    if (pid.ObjectClass.DxfName != "INSERT") continue;
                    var parentBR = tr.GetObject(pid, OpenMode.ForRead) as BlockReference;
                    if (parentBR == null) continue;
                    if (!string.Equals(parentBR.Name, parentBlockName, StringComparison.OrdinalIgnoreCase)) continue;

                    var parentBtr = (BlockTableRecord)tr.GetObject(parentBR.BlockTableRecord, OpenMode.ForRead);
                    var parentXf = parentBR.BlockTransform;

                    foreach (ObjectId nid in parentBtr)
                    {
                        if (nid.ObjectClass.DxfName != "INSERT") continue;
                        var nestedBR = tr.GetObject(nid, OpenMode.ForRead) as BlockReference;
                        if (nestedBR == null) continue;
                        if (!string.Equals(nestedBR.Name, nestedBlockName, StringComparison.OrdinalIgnoreCase)) continue;

                        Extents3d ext;
                        try { ext = nestedBR.GeometricExtents; }
                        catch { continue; }

                        hits.Add(new NestedFrameHit
                        {
                            ParentRefId = pid,
                            NestedRefId = nid,
                            ParentToWcs = parentXf,
                            LocalMin = ext.MinPoint,
                            LocalMax = ext.MaxPoint,
                        });
                    }
                }
                tr.Commit();
            }
            return hits;
        }
    }
}
