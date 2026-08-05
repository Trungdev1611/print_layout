using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace PrintLayoutAddin.Core
{
    /// <summary>
    /// One block-cleanup candidate: a block definition in the current DB that looks like
    /// something PLAUTO generated (name starts with "PL_"), plus the count of its BR
    /// instances sitting in ModelSpace.
    /// </summary>
    public class CleanupCandidate
    {
        public string BlockName;
        public ObjectId BlockDefId;
        public int InstanceCount;

        public override string ToString()
        {
            return $"{BlockName}  ({InstanceCount} instance{(InstanceCount == 1 ? "" : "s")})";
        }
    }

    /// <summary>
    /// Creates a native block definition containing a rectangle + INNO-STT attribute,
    /// then inserts one BlockReference per detected nested-xref frame hit at the matching
    /// world-space position / rotation / scale, so PLSTT + PLAYOUT can operate on them.
    /// </summary>
    public static class NativeFrameBuilder
    {
        /// <summary>
        /// List all block definitions in <paramref name="db"/> whose name starts with "PL_"
        /// (convention for PLAUTO-generated blocks), plus ModelSpace instance count.
        /// </summary>
        public static List<CleanupCandidate> ListAutoBlocks(Database db)
        {
            var byName = new Dictionary<string, CleanupCandidate>(StringComparer.OrdinalIgnoreCase);
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                foreach (ObjectId id in bt)
                {
                    var btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    if (btr.IsLayout) continue;
                    if (btr.IsFromExternalReference) continue;
                    if (!btr.Name.StartsWith("PL_", StringComparison.OrdinalIgnoreCase)) continue;

                    byName[btr.Name] = new CleanupCandidate
                    {
                        BlockName = btr.Name,
                        BlockDefId = id,
                        InstanceCount = 0
                    };
                }

                // Count instances in ModelSpace
                var ms = (BlockTableRecord)tr.GetObject(
                    SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);
                foreach (ObjectId id in ms)
                {
                    if (id.ObjectClass.DxfName != "INSERT") continue;
                    var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                    if (br == null) continue;
                    if (byName.TryGetValue(br.Name, out var cand))
                        cand.InstanceCount++;
                }
                tr.Commit();
            }
            return byName.Values.OrderBy(c => c.BlockName).ToList();
        }

        /// <summary>
        /// Erase every BlockReference of <paramref name="blockName"/> in ModelSpace,
        /// then purge the block definition. Returns (instancesErased, blockPurged).
        /// Also outputs the list of INNO-STT values read from the instances (so caller
        /// can optionally delete matching layouts).
        /// </summary>
        public static (int InstancesErased, bool BlockPurged, List<string> Stts) RemoveAutoBlock(
            Database db, string blockName, string attributeTag)
        {
            int erased = 0;
            bool purged = false;
            var stts = new List<string>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (!bt.Has(blockName))
                {
                    tr.Commit();
                    return (0, false, stts);
                }
                var btrId = bt[blockName];

                var ms = (BlockTableRecord)tr.GetObject(
                    SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);

                // First pass: collect STTs before erasing so we can return them.
                foreach (ObjectId id in ms)
                {
                    if (id.ObjectClass.DxfName != "INSERT") continue;
                    var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                    if (br == null) continue;
                    if (!string.Equals(br.Name, blockName, StringComparison.OrdinalIgnoreCase)) continue;

                    foreach (ObjectId attId in br.AttributeCollection)
                    {
                        var att = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                        if (att == null) continue;
                        if (string.Equals(att.Tag, attributeTag, StringComparison.OrdinalIgnoreCase))
                        {
                            if (!string.IsNullOrWhiteSpace(att.TextString))
                                stts.Add(att.TextString);
                        }
                    }
                }

                // Second pass: erase. (Do separately to avoid mutating the iteration source.)
                foreach (ObjectId id in ms)
                {
                    if (id.ObjectClass.DxfName != "INSERT") continue;
                    var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                    if (br == null) continue;
                    if (!string.Equals(br.Name, blockName, StringComparison.OrdinalIgnoreCase)) continue;

                    var brW = (BlockReference)tr.GetObject(id, OpenMode.ForWrite);
                    brW.Erase();
                    erased++;
                }

                // Purge the BTR itself if nothing else references it.
                var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForWrite);
                var refIds = new ObjectIdCollection();
                btr.GetBlockReferenceIds(true, true);
                // GetBlockReferenceIds returns *live* references. After erasing above, there should be none.
                try
                {
                    btr.Erase();
                    purged = true;
                }
                catch
                {
                    // e.g. still referenced somewhere (inside another block definition). Leave it.
                    purged = false;
                }

                tr.Commit();
            }
            return (erased, purged, stts);
        }

        /// <summary>
        /// Delete layouts whose name matches any value in <paramref name="layoutNames"/>.
        /// Returns count of layouts actually deleted.
        /// </summary>
        public static int DeleteLayoutsByName(Database db, IEnumerable<string> layoutNames)
        {
            int deleted = 0;
            var lm = LayoutManager.Current;
            foreach (var name in layoutNames)
            {
                if (string.IsNullOrWhiteSpace(name)) continue;
                try
                {
                    if (lm.LayoutExists(name))
                    {
                        lm.DeleteLayout(name);
                        deleted++;
                    }
                }
                catch { }
            }
            return deleted;
        }

        /// <summary>
        /// Ensure a block definition named <paramref name="blockName"/> exists in <paramref name="db"/>,
        /// sized width x height with an INNO-STT attribute centred inside. If the block already exists,
        /// it is replaced so size stays in sync with detected frames.
        /// </summary>
        public static ObjectId EnsureFrameBlock(Database db, string blockName, double width, double height, string attributeTag)
        {
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);

                BlockTableRecord btr;
                ObjectId btrId;
                if (bt.Has(blockName))
                {
                    btrId = bt[blockName];
                    btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForWrite);
                    // Clear old contents so we can rebuild the geometry
                    foreach (ObjectId eid in btr)
                    {
                        var ent = tr.GetObject(eid, OpenMode.ForWrite) as Entity;
                        ent?.Erase();
                    }
                }
                else
                {
                    btr = new BlockTableRecord
                    {
                        Name = blockName,
                        Origin = Point3d.Origin
                    };
                    btrId = bt.Add(btr);
                    tr.AddNewlyCreatedDBObject(btr, true);
                }

                // Rectangle from (0,0) to (width, height) — closed LWPolyline
                var pl = new Polyline(4);
                pl.AddVertexAt(0, new Point2d(0, 0), 0, 0, 0);
                pl.AddVertexAt(1, new Point2d(width, 0), 0, 0, 0);
                pl.AddVertexAt(2, new Point2d(width, height), 0, 0, 0);
                pl.AddVertexAt(3, new Point2d(0, height), 0, 0, 0);
                pl.Closed = true;
                btr.AppendEntity(pl);
                tr.AddNewlyCreatedDBObject(pl, true);

                // AttributeDefinition for INNO-STT — centred, height ≈ 5% of frame height
                double h = Math.Max(1.0, Math.Min(width, height) * 0.05);
                var att = new AttributeDefinition
                {
                    Tag = attributeTag,
                    Prompt = attributeTag,
                    TextString = "",
                    Position = new Point3d(width / 2.0, height / 2.0, 0),
                    Height = h,
                    Justify = AttachmentPoint.MiddleCenter,
                    AlignmentPoint = new Point3d(width / 2.0, height / 2.0, 0),
                    LockPositionInBlock = true,
                };
                btr.AppendEntity(att);
                tr.AddNewlyCreatedDBObject(att, true);

                tr.Commit();
                return btrId;
            }
        }

        /// <summary>
        /// Insert one BlockReference per hit into ModelSpace, matching each hit's WCS placement.
        /// Returns the number of references created.
        /// </summary>
        public static int InsertFrames(Database db, string blockName, IList<NestedFrameHit> hits)
        {
            int count = 0;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (!bt.Has(blockName))
                    throw new InvalidOperationException($"Block '{blockName}' not defined.");
                var btrId = bt[blockName];
                var nativeBtr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

                var ms = (BlockTableRecord)tr.GetObject(
                    SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);

                foreach (var hit in hits)
                {
                    // The native block's (0,0) should map to the nested frame's local MinPoint,
                    // then run through parent → WCS. Compose: parent * T(localMin).
                    var shift = Matrix3d.Displacement(
                        new Vector3d(hit.LocalMin.X, hit.LocalMin.Y, 0));
                    var xform = hit.ParentToWcs * shift;

                    var br = new BlockReference(Point3d.Origin, btrId);
                    br.BlockTransform = xform;
                    ms.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);

                    // Create an AttributeReference for each non-constant AttributeDefinition in the native BTR.
                    foreach (ObjectId eid in nativeBtr)
                    {
                        var ad = tr.GetObject(eid, OpenMode.ForRead) as AttributeDefinition;
                        if (ad == null) continue;
                        if (ad.Constant) continue;

                        var ar = new AttributeReference();
                        ar.SetAttributeFromBlock(ad, br.BlockTransform);
                        ar.TextString = ad.TextString ?? "";
                        br.AttributeCollection.AppendAttribute(ar);
                        tr.AddNewlyCreatedDBObject(ar, true);
                    }

                    count++;
                }

                tr.Commit();
            }
            return count;
        }
    }
}
