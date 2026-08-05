using System;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace PrintLayoutAddin.Core
{
    public static class LayoutBuilder
    {
        /// <summary>
        /// Copy an existing template layout (e.g. "Layout1") to a new layout with the given name.
        /// Any user viewports inherited from the template are erased so the command can add its own.
        /// </summary>
        public static ObjectId CreateLayoutFromTemplate(Database db, string layoutName, string templateLayoutName)
        {
            var lm = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;
            if (!lm.LayoutExists(templateLayoutName))
                throw new InvalidOperationException(
                    $"Template layout '{templateLayoutName}' not found. " +
                    "Edit 'templateLayout' in config.json or rename your source layout.");

            // CopyLayout clones paper space BTR contents (xrefs, annotations, viewports, view state).
            lm.CopyLayout(templateLayoutName, layoutName);
            var layoutId = lm.GetLayoutId(layoutName);

            // Remove any user viewports inherited from the template so we can place our own cleanly.
            // Temporarily unlock any layers involved — otherwise Erase throws eOnLockedLayer.
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                var paperBtr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                foreach (ObjectId id in paperBtr)
                {
                    if (id.ObjectClass.DxfName != "VIEWPORT") continue;
                    var vp = tr.GetObject(id, OpenMode.ForWrite) as Viewport;
                    if (vp == null) continue;
                    if (vp.Number == 1) continue; // keep the paper space overall viewport

                    TemporarilyUnlockLayer(vp.LayerId, tr);
                    vp.Erase();
                }
                tr.Commit();
            }

            return layoutId;
        }

        private static void TemporarilyUnlockLayer(ObjectId layerId, Transaction tr)
        {
            if (layerId.IsNull) return;
            var ltr = tr.GetObject(layerId, OpenMode.ForRead) as LayerTableRecord;
            if (ltr == null) return;
            if (ltr.IsLocked)
            {
                ltr.UpgradeOpen();
                ltr.IsLocked = false;
            }
        }

        /// <summary>
        /// Caller can inspect <see cref="LastLayerAction"/> right after <see cref="AddViewport"/>
        /// to report whether the VP layer was freshly created or reused from an existing one.
        /// </summary>
        public static string LastLayerAction { get; private set; } = "";

        public static string AddViewport(Database db, ObjectId layoutId,
            Point3d p1, Point3d p2, FrameInfo frame, string vpLayerName)
        {
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                var paperBtr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite);

                ObjectId layerId = EnsureLayer(db, vpLayerName, tr);

                double minX = Math.Min(p1.X, p2.X);
                double maxX = Math.Max(p1.X, p2.X);
                double minY = Math.Min(p1.Y, p2.Y);
                double maxY = Math.Max(p1.Y, p2.Y);
                double vpWidth = maxX - minX;
                double vpHeight = maxY - minY;
                if (vpWidth <= 0 || vpHeight <= 0) return "vpSize<=0";

                var vp = new Viewport
                {
                    CenterPoint = new Point3d((minX + maxX) / 2.0, (minY + maxY) / 2.0, 0),
                    Width = vpWidth,
                    Height = vpHeight,
                    LayerId = layerId
                };
                paperBtr.AppendEntity(vp);
                tr.AddNewlyCreatedDBObject(vp, true);

                // Turn viewport on BEFORE setting view properties (required by AutoCAD).
                vp.On = true;

                // Reset every view-related property explicitly so nothing inherits from template.
                vp.PerspectiveOn = false;
                vp.FrontClipOn = false;
                vp.BackClipOn = false;
                vp.TwistAngle = 0.0;
                vp.ViewDirection = new Vector3d(0, 0, 1);

                // Target = frame center in WCS; DCS center = origin relative to target.
                vp.ViewTarget = new Point3d(frame.Center.X, frame.Center.Y, 0);
                vp.ViewCenter = new Point2d(0, 0);

                // Fit the frame rectangle into the paper viewport, preserving aspect.
                double paperAspect = vpWidth / vpHeight; // paper W/H
                double modelWidth = frame.Width;
                double modelHeight = frame.Height;
                double viewH;
                if (modelWidth > 0 && modelHeight > 0 && paperAspect > 0)
                {
                    double byHeight = modelHeight;
                    double byWidth = modelWidth / paperAspect;
                    viewH = Math.Max(byHeight, byWidth);
                }
                else
                {
                    viewH = modelHeight > 0 ? modelHeight : 1.0;
                }

                vp.ViewHeight = viewH;

                // Force view recompute via on/off toggle (AutoCAD quirk).
                vp.On = false;
                vp.On = true;

                vp.Locked = true;
                tr.Commit();

                return string.Format(
                    "paper({0:F1}x{1:F1} @ c({2:F1},{3:F1})) | model({4:F1}x{5:F1} @ c({6:F1},{7:F1})) viewH={8:F1}",
                    vpWidth, vpHeight, (minX + maxX) / 2.0, (minY + maxY) / 2.0,
                    modelWidth, modelHeight, frame.Center.X, frame.Center.Y, viewH);
            }
        }

        private static ObjectId EnsureLayer(Database db, string layerName, Transaction tr)
        {
            var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            if (lt.Has(layerName))
            {
                // Reuse the existing layer as-is; keep its colour / linetype / plottable
                // exactly as the user defined them. We only touch lock/freeze because a locked
                // or frozen layer would block viewport creation — that's a technical constraint,
                // not a stylistic change.
                var id = lt[layerName];
                var existing = (LayerTableRecord)tr.GetObject(id, OpenMode.ForRead);
                bool touched = false;
                if (existing.IsLocked)
                {
                    existing.UpgradeOpen();
                    existing.IsLocked = false;
                    touched = true;
                }
                if (existing.IsFrozen)
                {
                    if (!existing.IsWriteEnabled) existing.UpgradeOpen();
                    existing.IsFrozen = false;
                    touched = true;
                }
                LastLayerAction = touched
                    ? $"reused existing layer '{layerName}' (unlocked/unfrozen to allow viewport)"
                    : $"reused existing layer '{layerName}'";
                return id;
            }

            lt.UpgradeOpen();
            var rec = new LayerTableRecord
            {
                Name = layerName,
                IsPlottable = false
            };
            var newId = lt.Add(rec);
            tr.AddNewlyCreatedDBObject(rec, true);
            LastLayerAction = $"created layer '{layerName}'";
            return newId;
        }
    }
}
