using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using PrintLayoutAddin.Core;
using PrintLayoutAddin.UI;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Registry = Microsoft.Win32.Registry;

namespace PrintLayoutAddin
{
    public class Commands
    {
        private const string RegKey = @"Software\PrintLayoutAddin";
        private static TitleBlockSetupPalette _titleBlockSetupPalette;

        [CommandMethod("PLSTT")]
        public void PlStt()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            if (!LicenseGate.Allow(ed)) return;

            if (!IsInModelSpace(db))
            {
                ed.WriteMessage("\nPLSTT must be run in ModelSpace.");
                return;
            }

            // Step 1 — polyline
            var peo = new PromptEntityOptions("\nSelect guide polyline crossing the frames: ");
            peo.SetRejectMessage("\nObject must be a polyline.");
            peo.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), true);
            peo.AddAllowedClass(typeof(Polyline2d), true);
            peo.AddAllowedClass(typeof(Polyline3d), true);
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;
            var polyId = per.ObjectId;

            // Step 2 — block
            var choice = PickBlock(db);
            if (choice == null) return;

            // Step 3 — pre-count how many frames the polyline actually hits.
            // This is what the dialog validates the generated code list against.
            int expectedCount;
            try
            {
                expectedCount = SttAssigner.CountMatches(db, polyId, choice.Name);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nPLSTT pre-scan failed: {ex.Message}");
                return;
            }
            if (expectedCount == 0)
            {
                ed.WriteMessage($"\nPolyline does not pass through any '{choice.Name}' frame. Abort.");
                return;
            }
            ed.WriteMessage($"\nPolyline will number {expectedCount} frame(s).");

            // Step 4 — dialog: choose mode, generate / import codes, validate preview.
            System.Collections.Generic.List<string> codes;
            System.Collections.Generic.List<string> drawingNames;
            bool allowMismatch;
            using (var dlg = new SttOptionsDialog(expectedCount))
            {
                if (AcadApp.ShowModalDialog(dlg) != DialogResult.OK) return;
                codes = dlg.Codes;
                drawingNames = dlg.DrawingNames;
                allowMismatch = dlg.AllowCountMismatch;
            }
            if (codes == null || codes.Count == 0)
            {
                ed.WriteMessage("\nNo codes to apply.");
                return;
            }

            // Step 5 — apply. SttAssigner rolls back the transaction if
            // the matched-frame count != codes.Count and allowMismatch is false.
            try
            {
                var result = SttAssigner.ApplyNumbersToSelectedFrames(
                    db, ed, polyId, choice.Name, codes, drawingNames, allowMismatch);
                if (result.Aborted)
                {
                    ed.WriteMessage("\nPLSTT aborted — " + result.Message);
                    ed.WriteMessage("\nTip: tick 'Allow count mismatch' in the dialog to apply the overlap anyway.");
                    return;
                }
                ed.WriteMessage($"\nNumbered {result.Assigned} frame(s). " +
                                $"Vertices missing a frame: {result.VertexMissed}. " +
                                $"Frames left unnumbered: {result.FramesSkipped}. " +
                                $"Codes unused: {result.CodesUnused}.");
                if (drawingNames != null)
                {
                    ed.WriteMessage(
                        $"\nDrawing names assigned: {result.DrawingNamesAssigned}. " +
                        $"Frames missing attribute '{Config.Instance.DrawingNameTag}': " +
                        $"{result.DrawingNameAttributesMissing}.");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nPLSTT error: {ex.Message}");
            }
        }

        [CommandMethod("PLAYOUT")]
        public void PlLayout()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            if (!LicenseGate.Allow(ed)) return;

            var choice = PickBlock(db);
            if (choice == null) return;

            string templateLayout = Config.Instance.TemplateLayout;
            var lmCheck = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;
            if (!lmCheck.LayoutExists(templateLayout))
            {
                ed.WriteMessage($"\nTemplate layout '{templateLayout}' not found. " +
                    "Prepare a layout containing the title block xref at the correct position (default name 'Layout1'), " +
                    "or change 'templateLayout' in config.json.");
                return;
            }

            var frames = FrameScanner.CollectFrames(db, choice.Name, requireStt: true);
            if (frames.Count == 0)
            {
                ed.WriteMessage("\nNo frame carries an INNO-STT value. Run PLSTT first.");
                return;
            }
            frames = frames.OrderBy(f => f.Stt, FrameScanner.SttComparer).ToList();

            string dwgPath = null;
            try { dwgPath = doc.Name; } catch { }
            var corners = ViewportCornerStore.Load(dwgPath);
            if (!corners.HasValue)
            {
                ed.WriteMessage(
                    "\nNo saved viewport corners for this drawing.\n"
                    + "Open Title Block Setup, pick P1/P2 on the template layout, then run Build Layouts again.");
                return;
            }

            int created = 0, skipped = 0, errors = 0;
            DuplicateAction? rememberedAction = null;
            var lm = Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;
            string vpLayer = Config.Instance.VpLayer;
            var orderedNames = new List<string>();

            // Unlock any locked layers for the duration of the command —
            // CopyLayout + Viewport creation + template-viewport erase all fail with eOnLockedLayer
            // if a relevant layer is locked (typically layer "0" or the VP layer from a previous run).
            var relockIds = UnlockAllLayers(db);

            using (LayoutDstSyncWatcher.Suppress())
            try
            {
            foreach (var frame in frames)
            {
                string name = frame.Stt;

                if (lm.LayoutExists(name))
                {
                    DuplicateAction act;
                    if (rememberedAction.HasValue)
                    {
                        act = rememberedAction.Value;
                    }
                    else
                    {
                        using (var dlg = new DuplicateLayoutDialog(name))
                        {
                            if (AcadApp.ShowModalDialog(dlg) != DialogResult.OK) return;
                            act = dlg.Action;
                            if (dlg.ApplyToAll) rememberedAction = act;
                        }
                    }

                    switch (act)
                    {
                        case DuplicateAction.Abort:
                            ed.WriteMessage("\nCommand aborted.");
                            return;
                        case DuplicateAction.Skip:
                            skipped++;
                            if (lm.LayoutExists(name))
                            {
                                orderedNames.Add(name);
                                // Kept as-is, but it is still one of ours — stamp it so a later
                                // delete offers the DST cleanup (also backfills pre-2.0 drawings).
                                PlayoutLayoutStamp.Stamp(db, lm.GetLayoutId(name));
                            }
                            continue;
                        case DuplicateAction.Overwrite:
                            lm.DeleteLayout(name);
                            break;
                        case DuplicateAction.Rename:
                            name = UniqueLayoutName(lm, name);
                            break;
                    }
                }

                try
                {
                    var layoutId = LayoutBuilder.CreateLayoutFromTemplate(db, name, templateLayout);
                    orderedNames.Add(name);

                    // Viewport entity creation requires TILEMODE = 0 (paper space active).
                    // Switch current layout for every frame, not only when picking corners.
                    lm.CurrentLayout = name;

                    var diag = LayoutBuilder.AddViewport(db, layoutId, corners.Value.P1, corners.Value.P2, frame, vpLayer);
                    if (created == 0 && !string.IsNullOrEmpty(LayoutBuilder.LastLayerAction))
                        ed.WriteMessage("\n  " + LayoutBuilder.LastLayerAction + ".");
                    ed.WriteMessage($"\n  [{name}] {diag}");
                    created++;
                }
                catch (System.Exception ex)
                {
                    errors++;
                    ed.WriteMessage($"\nFailed to create layout {name}: {ex.Message}");
                }
            }

            }
            finally
            {
                try
                {
                    var tabs = LayoutBuilder.PlaceLayoutsAfterTemplate(db, templateLayout, orderedNames);
                    if (!string.IsNullOrWhiteSpace(tabs))
                        ed.WriteMessage("\nTab order: Model, " + tabs);
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nCould not reorder layout tabs: {ex.Message}");
                }
                RelockLayers(db, relockIds);
            }

            ed.WriteMessage($"\nDone. Created {created} layouts, skipped {skipped}, errors {errors}.");
            ed.WriteMessage("\nTitle-block Sheet Set fields may show #### until you Create / Update a DST.");
            try { LayoutDstSyncWatcher.RefreshSnapshot(); } catch { }

            if (!Config.Instance.AutoSheetSetAfterLayout)
            {
                SheetSetAutoLog.Write(ed, doc?.Name,
                    "skipped — no auto DST after PLAYOUT (open Create Sheet Set to write a .dst)");
            }
            else
            {
                TryAutoSheetSetAfterLayout(doc, ed, created, skipped, errors);
            }
        }

        /// <summary>
        /// After PLAYOUT: write/update default DST so Sheet Set fields on the title block
        /// resolve (instead of ####) without opening PLSHEETSET.
        /// </summary>
        private static void TryAutoSheetSetAfterLayout(
            Document doc, Editor ed, int created, int skipped, int errors)
        {
            string dwgPath = null;
            try { dwgPath = doc?.Name; } catch { }

            SheetSetAutoLog.Write(ed, dwgPath,
                $"start after PLAYOUT (created={created}, skipped={skipped}, errors={errors})");

            if (doc == null || ed == null)
            {
                SheetSetAutoLog.Write(ed, dwgPath, "abort — no active document");
                return;
            }

            bool savedToDisk = false;
            try { savedToDisk = !string.IsNullOrWhiteSpace(dwgPath) && File.Exists(dwgPath); }
            catch { }
            if (!savedToDisk)
            {
                SheetSetAutoLog.Write(ed, dwgPath,
                    "abort — DWG not saved to disk. Save then run PLAYOUT / PLSHEETSET again.");
                return;
            }

            try
            {
                var layouts = LayoutPlotter.GetPrintableLayouts(doc.Database);
                var titled = SheetSetFolderImport.GetImportableLayoutNames(doc.Database);
                layouts = (layouts ?? new System.Collections.Generic.List<PrintableLayout>())
                    .Where(l => l != null && titled.Contains(l.Name))
                    .ToList();
                int layoutCount = layouts.Count;
                SheetSetAutoLog.Write(ed, dwgPath,
                    $"printable titled layouts={layoutCount}, dst={PublishPaths.DefaultDstPath(dwgPath)}");

                if (layoutCount == 0)
                {
                    SheetSetAutoLog.Write(ed, dwgPath,
                        "abort — no layouts with DrawingName (template/untitled excluded)");
                    return;
                }

                var drawingNames = FrameScanner.CollectDrawingNamesByStt(doc.Database);
                SheetSetAutoLog.Write(ed, dwgPath,
                    $"drawing-name map entries={drawingNames?.Count ?? 0}");

                if (!EnsureSavedForSheetSet(doc, ed, dwgPath))
                {
                    NotifyAutoSheetSetFailed(
                        "The drawing could not be saved, and AcSm can only import layouts that "
                        + "exist in the saved .dwg. See the log for details.");
                    return;
                }

                SheetSetAutoLog.Write(ed, dwgPath, "calling CreateOrReplace (silent)…");
                var sync = SheetSetService.TryAutoSyncFromLayouts(dwgPath, layouts, drawingNames);
                SheetSetAutoLog.Write(ed, dwgPath,
                    sync.Ok
                        ? $"OK sheets={sync.SheetCount} file={sync.DstPath}"
                        : $"FAIL {sync.Message}");

                if (!sync.Ok)
                {
                    NotifyAutoSheetSetFailed(sync.Message);
                    return;
                }

                // Remember which DST these layouts belong to, so deleting a tab later can
                // point the .dst picker straight at it.
                try { PlayoutLayoutStamp.SetDstPathForStamped(doc.Database, sync.DstPath); } catch { }
                try { LayoutDstSyncWatcher.RefreshSnapshot(); } catch { }

                // Deliberately no ReloadForUser/SHEETSET here. Opening the DST in Sheet Set
                // Manager on every PLAYOUT planted a second owner on the file, and the
                // release/reopen churn around it was what broke later Create/Update calls.
                // UpdateInMemoryDwgHints (inside the sync) is what the Fields need — not SSM.
                try
                {
                    // Queue REGEN so Sheet Set fields re-evaluate after DST association.
                    doc.SendStringToExecute("_.REGEN ", true, false, false);
                    SheetSetAutoLog.Write(ed, dwgPath, "queued _.REGEN");
                }
                catch (System.Exception regenEx)
                {
                    SheetSetAutoLog.Write(ed, dwgPath, "REGEN error: " + regenEx.Message);
                }
            }
            catch (System.Exception ex)
            {
                SheetSetAutoLog.Write(ed, dwgPath, "exception: " + ex.Message);
                NotifyAutoSheetSetFailed(ex.Message);
            }
        }

        /// <summary>
        /// AcSm's ImportSheet resolves each layout through the .dwg as SAVED ON DISK, so the
        /// layouts PLAYOUT just created stay invisible to it until the drawing is written out —
        /// the sync then fails with a bare HRESULT 0x80040211 that names no cause. Save first.
        /// </summary>
        internal static bool EnsureSavedForSheetSet(Document doc, Editor ed, string dwgPath)
        {
            int dbmod = 0;
            try { dbmod = Convert.ToInt32(AcadApp.GetSystemVariable("DBMOD")); }
            catch { }
            if (dbmod == 0)
            {
                SheetSetAutoLog.Write(ed, dwgPath, "DWG has no unsaved changes (DBMOD=0)");
                return true;
            }

            try
            {
                // Lock the document — SaveAs without it throws when called from this context.
                using (doc.LockDocument())
                {
                    doc.Database.SaveAs(
                        doc.Name,
                        true,
                        doc.Database.OriginalFileVersion,
                        doc.Database.SecurityParameters);
                }
                SheetSetAutoLog.Write(ed, dwgPath,
                    $"saved DWG before sheet set sync (DBMOD was {dbmod}) — "
                    + "ImportSheet reads layouts from the file on disk");
                return true;
            }
            catch (System.Exception ex)
            {
                SheetSetAutoLog.Write(ed, dwgPath,
                    "could not save the DWG before sheet set sync: " + ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Layouts from PLAYOUT are fine; only the silent DST sync failed.
        /// </summary>
        private static void NotifyAutoSheetSetFailed(string detail)
        {
            string reason = string.IsNullOrWhiteSpace(detail) ? "(unknown error)" : detail.Trim();
            try
            {
                MessageBox.Show(
                    "Build Layouts finished, but auto Sheet Set (DST) failed.\n\n"
                    + reason
                    + "\n\nLayouts are OK. To fix title-block fields (####):\n"
                    + "1. Save the drawing (Ctrl+S) — the sheet set can only import layouts\n"
                    + "   that exist in the saved .dwg.\n"
                    + "2. Run PLSHEETSET (Create / Update) manually.\n"
                    + "3. REGEN the layout.\n\n"
                    + "Details: "
                    + SheetSetAutoLog.GetLogFilePath(
                        AcadApp.DocumentManager.MdiActiveDocument?.Name),
                    "Print Layout — Sheet Set",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            catch { }
        }

        private static System.Collections.Generic.List<ObjectId> UnlockAllLayers(Database db)
        {
            var relock = new System.Collections.Generic.List<ObjectId>();
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    foreach (ObjectId id in lt)
                    {
                        var ltr = (LayerTableRecord)tr.GetObject(id, OpenMode.ForRead);
                        if (ltr.IsLocked)
                        {
                            ltr.UpgradeOpen();
                            ltr.IsLocked = false;
                            relock.Add(id);
                        }
                    }
                    tr.Commit();
                }
            }
            catch { }
            return relock;
        }

        private static void RelockLayers(Database db, System.Collections.Generic.List<ObjectId> ids)
        {
            if (ids == null || ids.Count == 0) return;
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    foreach (var id in ids)
                    {
                        if (id.IsNull) continue;
                        var ltr = tr.GetObject(id, OpenMode.ForWrite) as LayerTableRecord;
                        if (ltr == null) continue;
                        ltr.IsLocked = true;
                    }
                    tr.Commit();
                }
            }
            catch { }
        }

        [CommandMethod("PLVP")]
        public void PlVpReset()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            string dwgPath = null;
            try { dwgPath = doc.Name; } catch { }

            string normalized = ViewportCornerStore.TryNormalizePath(dwgPath);
            if (normalized == null)
            {
                doc.Editor.WriteMessage(
                    "\nSave the drawing first — viewport corners are stored per DWG file path.");
                return;
            }

            if (ViewportCornerStore.Clear(dwgPath))
            {
                doc.Editor.WriteMessage(
                    "\nCleared viewport corners for this drawing:\n  " + normalized
                    + "\nPick P1/P2 again in Title Block Setup before Auto or Build Layouts.");
            }
            else
            {
                doc.Editor.WriteMessage(
                    "\nNo saved viewport corners for this drawing:\n  " + normalized);
            }
        }

        /// <summary>Diagnostic: print normalized viewport geometry from saved P1/P2 (step 2 check).</summary>
        [CommandMethod("PLVPGEO")]
        public void PlVpGeo()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            string dwgPath = null;
            try { dwgPath = doc.Name; } catch { }

            var corners = ViewportCornerStore.Load(dwgPath);
            if (!corners.HasValue)
            {
                ed.WriteMessage(
                    "\nNo saved viewport corners for this DWG. Pick P1/P2 in Title Block Setup first.");
                return;
            }

            var b = ViewportCornerGeometry.Normalize(corners.Value);
            ed.WriteMessage("\n[PLVPGEO] " + ViewportCornerGeometry.Describe(b));
            ed.WriteMessage(
                "\n  raw P1=(" + corners.Value.P1.X.ToString("F2") + "," + corners.Value.P1.Y.ToString("F2")
                + ") P2=(" + corners.Value.P2.X.ToString("F2") + "," + corners.Value.P2.Y.ToString("F2") + ")");
        }

        /// <summary>
        /// Diagnostic (step 3): scan title-block labels from config, print hits, highlight found texts.
        /// Requires saved P1/P2 (PLAYOUT) and Paper Space.
        /// </summary>
        [CommandMethod("PLFRAMESCAN")]
        public void PlFrameScan()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            if (db.TileMode)
            {
                ed.WriteMessage("\nSwitch to a paper-space layout, then run PLFRAMESCAN.");
                return;
            }

            string dwgPath = null;
            try { dwgPath = doc.Name; } catch { }
            var corners = ViewportCornerStore.Load(dwgPath);
            if (!corners.HasValue)
            {
                ed.WriteMessage(
                    "\nNo saved viewport corners. Pick P1/P2 in Title Block Setup first, then PLFRAMESCAN.");
                return;
            }

            var bounds = ViewportCornerGeometry.Normalize(corners.Value);
            TitleBlockLabelScanner.ScanResult scan;
            try
            {
                scan = TitleBlockLabelScanner.Scan(db, bounds);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nPLFRAMESCAN failed: " + ex.Message);
                return;
            }

            ed.WriteMessage("\n[PLFRAMESCAN] labels from config:");
            ed.WriteMessage("\n  sheetTitleLabel    = " + (Config.Instance.SheetTitleLabel ?? ""));
            ed.WriteMessage("\n  sheetNumberLabel   = " + (Config.Instance.SheetNumberLabel ?? ""));
            ed.WriteMessage("\n  sheetRevisionLabel = " + (Config.Instance.SheetRevisionLabel ?? ""));
            foreach (var line in (scan.Summarize() ?? "").Split('\n'))
                ed.WriteMessage("\n" + line);

            ed.WriteMessage("\n[PLFRAMESCAN] cell corners + place point (method B):");
            foreach (var h in scan.Hits)
            {
                if (TitleBlockLabelScanner.TryResolveCellCenter(
                        db, h, scan.FullStrip, out var placeAt, out var cell, out var detail))
                {
                    ed.WriteMessage("\n  --- " + h.Kind + " ---");
                    ed.WriteMessage("\n    label@ (" + h.Position.X.ToString("F2") + "," + h.Position.Y.ToString("F2") + ")");
                    ed.WriteMessage("\n    " + cell.FormatCorners());
                    ed.WriteMessage(string.Format(
                        "\n    place@ ({0:F2},{1:F2})  W={2:F2} H={3:F2}",
                        placeAt.X, placeAt.Y, cell.Width, cell.Height));
                }
                else
                {
                    ed.WriteMessage("\n  " + h.Kind + ": NO CELL — " + detail);
                }
            }

            var ids = new System.Collections.Generic.List<ObjectId>();
            foreach (var h in scan.Hits)
            {
                if (!h.EntityId.IsNull && !h.EntityId.IsErased)
                    ids.Add(h.EntityId);
            }

            if (ids.Count > 0)
            {
                try
                {
                    ed.SetImpliedSelection(ids.ToArray());
                    ed.WriteMessage(
                        "\nHighlighted " + ids.Count
                        + " label text(s). Check grips/selection — then Esc to clear.");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage("\nCould not highlight: " + ex.Message
                        + " (nested block text may not select directly).");
                }
            }
            else
            {
                ed.WriteMessage(
                    "\nNo labels highlighted. Widen titleStripScanWidth or check Paper Space / config wording.");
            }
        }

        /// <summary>
        /// Step 3 place test: scan labels, find line-bounded cells, insert centered DRAWING_* placeholders.
        /// </summary>
        [CommandMethod("PLFRAMEPLACE")]
        public void PlFramePlace()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            if (db.TileMode)
            {
                ed.WriteMessage("\nSwitch to a paper-space layout, then run PLFRAMEPLACE.");
                return;
            }

            string dwgPath = null;
            try { dwgPath = doc.Name; } catch { }
            var corners = ViewportCornerStore.Load(dwgPath);
            if (!corners.HasValue)
            {
                ed.WriteMessage("\nNo saved viewport corners. Pick P1/P2 in Title Block Setup first.");
                return;
            }

            var bounds = ViewportCornerGeometry.Normalize(corners.Value);
            var scan = TitleBlockLabelScanner.Scan(db, bounds);
            ed.WriteMessage("\n[PLFRAMEPLACE]\n" + scan.Summarize());

            bool ok = TitleBlockSetupService.PlaceFieldsNearScannedLabels(
                db, scan, out string message);
            ed.WriteMessage("\n" + (message ?? (ok ? "Done." : "Failed.")));
            if (ok)
                ed.WriteMessage("\nTip: run button 5 / Activate later to link Sheet Set fields.");
        }

        /// <summary>
        /// Auto Frame Setup: place title fields, rev table, center title, activate Sheet Set.
        /// </summary>
        [CommandMethod("PLFRAMEAUTO")]
        public void PlFrameAuto()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            if (db.TileMode)
            {
                ed.WriteMessage("\nSwitch to a paper-space layout, then run PLFRAMEAUTO.");
                return;
            }

            string dwgPath = null;
            try { dwgPath = doc.Name; } catch { }
            var corners = ViewportCornerStore.Load(dwgPath);
            if (!corners.HasValue)
            {
                ed.WriteMessage("\nNo saved viewport corners. Pick P1/P2 in Title Block Setup first.");
                return;
            }

            var bounds = ViewportCornerGeometry.Normalize(corners.Value);
            bool ok;
            string message;
            using (doc.LockDocument())
            {
                ok = TitleBlockSetupService.RunAutoFrameSetup(db, bounds, out message);
            }
            ed.WriteMessage("\n[PLFRAMEAUTO]\n" + (message ?? (ok ? "Done." : "Failed.")));
        }

        [CommandMethod("PLPRINT")]
        public void PlPrint()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            if (!LicenseGate.Allow(ed)) return;

            System.Collections.Generic.List<PrintableLayout> layouts;
            try
            {
                layouts = LayoutPlotter.GetPrintableLayouts(db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nCould not read layouts: " + ex.Message);
                return;
            }

            if (layouts.Count == 0)
            {
                ed.WriteMessage("\nNo paper-space layouts found to print.");
                return;
            }

            PrintJobOptions options;
            try
            {
                using (var dlg = new PrintOptionsDialog(
                    db,
                    layouts,
                    Config.Instance.TemplateLayout,
                    DefaultPdfPath(doc)))
                {
                    if (AcadApp.ShowModalDialog(dlg) != DialogResult.OK) return;
                    options = dlg.Options;
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nCould not open the Print dialog (a plotter/printer driver may be broken or offline): " + ex.Message);
                return;
            }

            if (options == null) return;

            if (options.PlotToFile && !EnsureSavedForPublish(doc, ed)) return;

            if (options.PlotToFile && options.PdfOutputMode == PdfOutputMode.Combined && File.Exists(options.OutputPath))
            {
                var answer = MessageBox.Show(
                    $"The output file already exists:\n\n{options.OutputPath}\n\nOverwrite it?",
                    "Print / Export PDF",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if (answer != DialogResult.Yes) return;
                try { File.Delete(options.OutputPath); }
                catch (System.Exception ex)
                {
                    ed.WriteMessage("\nCould not overwrite output file: " + ex.Message);
                    return;
                }
            }
            else if (options.PlotToFile && options.PdfOutputMode == PdfOutputMode.SeparatePerLayout)
            {
                var existing = options.Layouts
                    .Select(l => Path.Combine(options.OutputPath, LayoutPlotter.MakeSafeFileName(l.Name) + ".pdf"))
                    .Where(File.Exists)
                    .Take(5)
                    .ToList();
                if (existing.Count > 0)
                {
                    var answer = MessageBox.Show(
                        "Some per-layout PDF files already exist and will be overwritten.\n\n" +
                        string.Join("\n", existing) +
                        "\n\nContinue?",
                        "Print / Export PDF",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);
                    if (answer != DialogResult.Yes) return;
                }
            }

            try
            {
                using (doc.LockDocument())
                {
                    LayoutPlotter.Plot(doc, options, msg => ed.WriteMessage("\n" + msg));
                }
                if (options.PlotToFile && options.PdfOutputMode == PdfOutputMode.SeparatePerLayout)
                    ed.WriteMessage($"\nPrint complete. Output folder: {options.OutputPath}");
                else
                    ed.WriteMessage(options.PlotToFile
                        ? $"\nPrint complete. Output: {options.OutputPath}"
                        : "\nPrint complete.");

                if (options.PlotToFile && options.OpenAfterExport)
                    OpenOutput(options, ed);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nPLPRINT failed: " + ex.Message);
            }
        }

        /// <summary>
        /// Internal: run DST layout-delete sync prompt off the Idle event
        /// (MessageBox inside Idle re-enters and loops).
        /// Keep attribute as a plain string literal — CadAddinManager/Mono.Cecil
        /// can crash on CommandMethod(..., CommandFlags) when resolving duplicates.
        /// </summary>
        [CommandMethod("PLDSTLAYOUTSYNC")]
        public void PlDstLayoutSync()
        {
            try { LayoutDstSyncWatcher.ProcessPendingFromCommand(); }
            catch (System.Exception ex)
            {
                var ed = AcadApp.DocumentManager?.MdiActiveDocument?.Editor;
                ed?.WriteMessage("\n" + LayoutDstSyncWatcher.SyncCommand + " failed: " + ex.Message);
            }
        }

        [CommandMethod("PLSHEETSET")]
        public void PlSheetSet()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            if (!LicenseGate.Allow(ed)) return;
            if (!EnsureSavedForPublish(doc, ed)) return;

            try
            {
                // Open whenever we are in paper space — do not require DrawingName layouts.
                // Untitled / template tabs are only omitted from the dialog seed list.
                int tileMode = 1;
                try { tileMode = Convert.ToInt32(AcadApp.GetSystemVariable("TILEMODE")); }
                catch { }
                if (tileMode != 0)
                {
                    ed.WriteMessage("\nSwitch to a paper-space layout, then run PLSHEETSET.");
                    return;
                }

                var titled = SheetSetFolderImport.GetImportableLayoutNames(doc.Database);
                var layouts = LayoutPlotter.GetPrintableLayouts(doc.Database)
                    .Where(x => x != null && !string.IsNullOrWhiteSpace(x.Name))
                    .Where(x => !string.Equals(
                        x.Name,
                        Config.Instance.TemplateLayout,
                        StringComparison.OrdinalIgnoreCase))
                    .Where(x => titled.Contains(x.Name))
                    .ToList();

                var drawingNames = FrameScanner.CollectDrawingNamesByStt(doc.Database);
                var defaultDstPath = PublishPaths.DefaultDstPath(doc.Name);
                using (var dlg = new SheetSetDialog(
                    layouts,
                    drawingNames,
                    doc.Name,
                    defaultDstPath))
                {
                    var sheetResult = AcadApp.ShowModalDialog(dlg);
                    if (sheetResult != DialogResult.OK) return;
                    if (dlg.ExportLayouts == null || dlg.ExportLayouts.Count == 0) return;
                    SavePrintLayoutSelection(dlg.ExportLayouts);
                }

                // Queue the existing, proven PLPRINT workflow. Its dialog opens
                // with exactly the order and selection chosen in the sheet-set table.
                doc.SendStringToExecute("_PLPRINT ", true, false, true);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nPLSHEETSET failed: " + ex.Message);
            }
        }

        private static void SavePrintLayoutSelection(
            System.Collections.Generic.IEnumerable<PrintableLayout> layouts)
        {
            try
            {
                var names = layouts
                    .Where(x => x != null && !string.IsNullOrWhiteSpace(x.Name))
                    .Select(x => x.Name)
                    .ToArray();
                using (var key = Registry.CurrentUser.CreateSubKey(RegKey))
                {
                    key?.SetValue(
                        "PrintLayoutOrder",
                        names,
                        Microsoft.Win32.RegistryValueKind.MultiString);
                    key?.SetValue(
                        "PrintLayoutChecked",
                        names,
                        Microsoft.Win32.RegistryValueKind.MultiString);
                }
            }
            catch { }
        }

        [CommandMethod("PLAUTO")]
        public void PlAuto()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            if (!LicenseGate.Allow(ed)) return;

            if (!IsInModelSpace(db))
            {
                ed.WriteMessage("\nPLAUTO must be run in ModelSpace.");
                return;
            }

            // Single-popup flow: auto-scan all ModelSpace xrefs/blocks (recursive),
            // user picks one frame name, then insert PL_ wrappers on every hit.
            ed.WriteMessage("\nPLAUTO — scanning ModelSpace for nested frames (all xrefs, any depth)...");
            List<BlockChoice> candidates;
            try
            {
                candidates = NestedFrameScanner.ListFramesInModelSpace(db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nPLAUTO scan failed: " + ex.Message);
                return;
            }

            if (candidates == null || candidates.Count == 0)
            {
                ed.WriteMessage(
                    "\nNo frame candidates found under ModelSpace "
                    + "(after skipping PL_ / A$ / * and tiny blocks).");
                return;
            }

            string lastNested = null;
            try
            {
                using (var k = Registry.CurrentUser.OpenSubKey(RegKey))
                    lastNested = k?.GetValue("LastNestedFrame") as string;
            }
            catch { }

            ed.WriteMessage($"\nFound {candidates.Count} candidate block name(s). Select the frame to wrap.");
            BlockChoice frameChoice;
            using (var dlg = new BlockPickerDialog(candidates, lastNested, "Select frame to wrap"))
            {
                if (AcadApp.ShowModalDialog(dlg) != DialogResult.OK) return;
                frameChoice = dlg.Selected;
            }
            if (frameChoice == null || string.IsNullOrWhiteSpace(frameChoice.Name))
                return;

            try
            {
                using (var k = Registry.CurrentUser.CreateSubKey(RegKey))
                    k?.SetValue("LastNestedFrame", frameChoice.Name);
            }
            catch { }

            string sourceBlockName = frameChoice.Name;
            var hits = NestedFrameScanner.CollectHitsInModelSpace(db, sourceBlockName);
            if (hits.Count == 0)
            {
                ed.WriteMessage($"\nFound 0 instances of '{sourceBlockName}' under ModelSpace.");
                return;
            }

            double w = hits[0].LocalWidth;
            double h = hits[0].LocalHeight;
            if (w <= 0 || h <= 0)
            {
                ed.WriteMessage($"\nSource block '{sourceBlockName}' has invalid extents (width/height <= 0). Abort.");
                return;
            }

            string safeNested = SymbolUtilityServices.RepairSymbolName(sourceBlockName, false);
            string nativeBlockName = "PL_" + safeNested;

            try
            {
                NativeFrameBuilder.EnsureFrameBlock(
                    db,
                    nativeBlockName,
                    w,
                    h,
                    Config.Instance.AttributeTag,
                    Config.Instance.DrawingNameTag);
                int inserted = NativeFrameBuilder.InsertFrames(db, nativeBlockName, hits);
                ed.WriteMessage(
                    $"\nCreated native block '{nativeBlockName}' ({w:F1} x {h:F1}) and inserted {inserted} frame(s) in ModelSpace.");
                ed.WriteMessage("\nNow run PLSTT to number them, then PLAYOUT to generate layouts.");

                try
                {
                    using (var k = Registry.CurrentUser.CreateSubKey(RegKey))
                        k?.SetValue("LastBlock", nativeBlockName);
                }
                catch { }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nPLAUTO failed: {ex.Message}");
            }
        }

        [CommandMethod("PLCLEAN")]
        public void PlClean()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            if (!LicenseGate.Allow(ed)) return;

            var candidates = NativeFrameBuilder.ListAutoBlocks(db);
            if (candidates.Count == 0)
            {
                ed.WriteMessage("\nNo PLAUTO-generated blocks found (no block names starting with 'PL_').");
                return;
            }

            System.Collections.Generic.List<CleanupCandidate> targets;
            bool alsoLayouts;
            using (var dlg = new CleanupDialog(candidates))
            {
                if (AcadApp.ShowModalDialog(dlg) != DialogResult.OK) return;
                targets = dlg.SelectedBlocks;
                alsoLayouts = dlg.AlsoDeleteLayouts;
            }

            string tag = Config.Instance.AttributeTag;
            int totalErased = 0, totalPurged = 0, totalLayoutsDeleted = 0;

            // Unlock layers for safety — erasing BRs on locked layers also throws eOnLockedLayer.
            var relockIds = UnlockAllLayers(db);
            try
            {
                foreach (var c in targets)
                {
                    try
                    {
                        var (erased, purged, stts) = NativeFrameBuilder.RemoveAutoBlock(db, c.BlockName, tag);
                        totalErased += erased;
                        if (purged) totalPurged++;

                        if (alsoLayouts && stts.Count > 0)
                        {
                            int deleted = NativeFrameBuilder.DeleteLayoutsByName(db, stts);
                            totalLayoutsDeleted += deleted;
                        }

                        ed.WriteMessage(
                            $"\n  {c.BlockName}: erased {erased} instance(s), " +
                            (purged ? "block purged" : "block kept (still referenced elsewhere)") +
                            (alsoLayouts ? $", {stts.Count} STT value(s) collected." : "."));
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\n  {c.BlockName}: failed — {ex.Message}");
                    }
                }
            }
            finally
            {
                RelockLayers(db, relockIds);
            }

            ed.WriteMessage(
                $"\nCleanup done. Erased {totalErased} instance(s), purged {totalPurged} block(s)" +
                (alsoLayouts ? $", deleted {totalLayoutsDeleted} layout(s)." : "."));
        }

        private static void ZoomPaperSpaceExtents(Document doc, ObjectId layoutId)
        {
            var db = doc.Database;
            var ed = doc.Editor;

            Extents3d? total = null;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                var paperBtr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                foreach (ObjectId id in paperBtr)
                {
                    var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent == null) continue;
                    if (ent is Viewport) continue;
                    Extents3d ext;
                    try { ext = ent.GeometricExtents; }
                    catch { continue; }
                    if (total == null) { total = ext; }
                    else { var t = total.Value; t.AddExtents(ext); total = t; }
                }
                tr.Commit();
            }

            if (!total.HasValue) return;
            var min = total.Value.MinPoint;
            var max = total.Value.MaxPoint;
            double cx = (min.X + max.X) / 2.0;
            double cy = (min.Y + max.Y) / 2.0;
            double w = Math.Max(1e-6, (max.X - min.X) * 1.1);
            double h = Math.Max(1e-6, (max.Y - min.Y) * 1.1);

            try
            {
                using (var view = ed.GetCurrentView())
                {
                    view.CenterPoint = new Point2d(cx, cy);
                    view.Width = w;
                    view.Height = h;
                    view.ViewDirection = new Vector3d(0, 0, 1);
                    ed.SetCurrentView(view);
                }
                ed.UpdateScreen();
            }
            catch { /* non-fatal: user can zoom manually */ }
        }

        private static (Point3d P1, Point3d P2)? PromptViewportCorners(Editor ed) =>
            ViewportCornerPicker.TryPrompt(ed, out var c) ? c : ((Point3d, Point3d)?)null;

        private static string UniqueLayoutName(Autodesk.AutoCAD.DatabaseServices.LayoutManager lm, string baseName)
        {
            int i = 1;
            while (lm.LayoutExists(baseName + "_" + i)) i++;
            return baseName + "_" + i;
        }

        private static BlockChoice PickBlock(Database db)
        {
            var list = FrameScanner.ListBlocksInModelSpace(db);
            if (list.Count == 0)
            {
                AcadApp.DocumentManager.MdiActiveDocument?.Editor
                    .WriteMessage("\nModelSpace contains no block/xref references.");
                return null;
            }

            string last = null;
            try
            {
                using (var k = Registry.CurrentUser.OpenSubKey(RegKey))
                    last = k?.GetValue("LastBlock") as string;
            }
            catch { }

            BlockChoice chosen;
            using (var dlg = new BlockPickerDialog(list, last))
            {
                if (AcadApp.ShowModalDialog(dlg) != DialogResult.OK) return null;
                chosen = dlg.Selected;
            }

            try
            {
                using (var k = Registry.CurrentUser.CreateSubKey(RegKey))
                    k?.SetValue("LastBlock", chosen.Name);
            }
            catch { }
            return chosen;
        }

        private static bool IsInModelSpace(Database db) => db.TileMode;

        private static string DefaultPdfPath(Document doc)
        {
            return PublishPaths.DefaultPdfPath(doc?.Name);
        }

        private static bool EnsureSavedForPublish(Document doc, Editor ed)
        {
            // Publishing reads the DWG from disk, so any edits made since the last save would be
            // silently left out of the PDF. Warn (and optionally save) when the drawing is dirty.
            bool savedToDisk = false;
            try { savedToDisk = !string.IsNullOrWhiteSpace(doc.Name) && File.Exists(doc.Name); }
            catch { }
            if (!savedToDisk) return true; // never saved: PublishPdf will report the proper message

            int dbmod = 0;
            try { dbmod = Convert.ToInt32(AcadApp.GetSystemVariable("DBMOD")); }
            catch { }
            if (dbmod == 0) return true;

            var answer = MessageBox.Show(
                "The drawing has unsaved changes.\n\n" +
                "PDF export reads the saved DWG from disk, so unsaved changes will NOT appear in the PDF.\n\n" +
                "Save the drawing now before exporting?",
                "Print / Export PDF",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Warning);
            if (answer == DialogResult.Cancel) return false;
            if (answer == DialogResult.No) return true; // proceed with the on-disk version

            try
            {
                // Must lock the document — SaveAs from a WinForms MessageBox callback
                // otherwise often throws and the command aborts with no dialog.
                using (doc.LockDocument())
                {
                    doc.Database.SaveAs(
                        doc.Name,
                        true,
                        doc.Database.OriginalFileVersion,
                        doc.Database.SecurityParameters);
                }
                ed.WriteMessage("\nDrawing saved.");
                return true;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nCould not save the drawing: " + ex.Message);
                MessageBox.Show(
                    "Could not save the drawing:\n\n" + ex.Message +
                    "\n\nTip: save manually (Ctrl+S), then run the command again.\n" +
                    "Or click No on the previous prompt to export the last saved version.",
                    "Print / Export PDF",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return false;
            }
        }

        private static void OpenOutput(PrintJobOptions options, Editor ed)
        {
            try
            {
                if (options.PdfOutputMode == PdfOutputMode.SeparatePerLayout)
                    System.Diagnostics.Process.Start("explorer.exe", "\"" + options.OutputPath + "\"");
                else
                    System.Diagnostics.Process.Start(
                        new System.Diagnostics.ProcessStartInfo(options.OutputPath) { UseShellExecute = true });
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nCould not open the output: " + ex.Message);
            }
        }

        [CommandMethod("PLFRAME_SETUP")]
        public void PlFrameSetup()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            if (!LicenseGate.Allow(ed)) return;

            ShowTitleBlockSetupPalette();
        }

        private static void ShowTitleBlockSetupPalette()
        {
            if (_titleBlockSetupPalette == null || _titleBlockSetupPalette.IsDisposed)
            {
                _titleBlockSetupPalette = new TitleBlockSetupPalette();
                AcadApp.ShowModelessDialog(_titleBlockSetupPalette);
            }
            else
            {
                _titleBlockSetupPalette.RefreshContext();
                if (!_titleBlockSetupPalette.Visible)
                    _titleBlockSetupPalette.Show();
                _titleBlockSetupPalette.BringToFront();
            }
        }

        [CommandMethod("PLLICENSE")]
        public void PlLicense()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc?.Editor;
            try
            {
                using (var dlg = new LicenseDialog())
                    AcadApp.ShowModalDialog(dlg);
            }
            catch (System.Exception ex)
            {
                ed?.WriteMessage("\nPLLICENSE error: " + ex.Message);
                return;
            }

            var info = LicenseManager.Current;
            ed?.WriteMessage("\n[License] " + (info.Message ?? info.State.ToString()));
        }

        [CommandMethod("PLKEYS")]
        public void PlKeys()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;

            ShortcutManager.Suspend(true);
            try
            {
                // Pick up changes made by another AutoCAD session before editing.
                ShortcutManager.Reload();
                using (var dlg = new ShortcutsDialog(ShortcutManager.CurrentConfig(), ShortcutManager.DefaultConfig()))
                {
                    if (AcadApp.ShowModalDialog(dlg) != DialogResult.OK) return;
                    ShortcutManager.Apply(dlg.Result);
                    ed.WriteMessage("\nKeyboard shortcuts & aliases updated. Hotkeys work immediately; "
                        + "typed aliases are active now (acad.pgp reloaded).");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nPLKEYS error: " + ex.Message);
            }
            finally
            {
                ShortcutManager.Suspend(false);
            }
        }
    }
}
