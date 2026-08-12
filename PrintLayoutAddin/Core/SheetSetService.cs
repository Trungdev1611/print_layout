using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

#if ACSM_INTEROP
using Autodesk.AutoCAD.Interop;
#endif

namespace PrintLayoutAddin.Core
{
    public enum SheetSetRowKind
    {
        Sheet = 0,
        Subset = 1,
    }

    public class SheetSetEntry
    {
        public bool Include { get; set; } = true;
        /// <summary>UI checkbox for Delete Selected (not written to DST).</summary>
        public bool Selected { get; set; }
        public int Order { get; set; }
        public SheetSetRowKind Kind { get; set; } = SheetSetRowKind.Sheet;
        /// <summary>Subset name for sheet rows; equals Title for subset header rows.</summary>
        public string SubsetName { get; set; }
        /// <summary>
        /// Nesting depth: 0 = root sheet; 1+ = subset level (SubsetLevel1…).
        /// Sheets inherit their parent subset's level for Title indent/tint.
        /// </summary>
        public int SubsetLevel { get; set; }
        public string SheetNumber { get; set; }
        public string Title { get; set; }
        /// <summary>Current sheet revision for CurrentSheetRevisionNumber (title block).</summary>
        public string Revision { get; set; } = "";
        public string DwgPath { get; set; }
        public PrintableLayout Layout { get; set; }
        /// <summary>Latest revision-history summary (table) for the clickable Rev history column.</summary>
        public string LastRevisionSummary { get; set; } = "";

        public bool IsSubset => Kind == SheetSetRowKind.Subset;
        public string LayoutName => Layout?.Name ?? "";
        public string DwgName => Path.GetFileName(DwgPath ?? "");
        public string DisplayLayout
        {
            get
            {
                // No leading spaces — nesting indent is only on the Title column in the UI.
                if (IsSubset)
                    return "▸ " + (Title ?? SubsetName ?? "Subset");
                return LayoutName ?? "";
            }
        }
    }

    /// <summary>One node read back from an existing .dst (subset header or sheet).</summary>
    public class SheetSetNodeInfo
    {
        public SheetSetRowKind Kind { get; set; }
        public string SubsetName { get; set; }
        /// <summary>0 = root sheet; 1+ = subset depth when Kind is Subset (or parent depth for sheets).</summary>
        public int SubsetLevel { get; set; }
        public string LayoutName { get; set; }
        /// <summary>Source DWG/DWT path from the sheet's layout reference (when available).</summary>
        public string DwgPath { get; set; }
        public string Number { get; set; }
        public string Title { get; set; }
        public string Revision { get; set; }
    }

    public class SheetSetReadResult
    {
        public string SheetSetName { get; set; }
        /// <summary>Ordered nodes: subset headers and sheets (nested levels preserved).</summary>
        public List<SheetSetNodeInfo> Nodes { get; } = new List<SheetSetNodeInfo>();

        /// <summary>Flat sheet list for callers that only need Number/Title by layout.</summary>
        public IEnumerable<SheetSetNodeInfo> Sheets =>
            Nodes.Where(n => n != null && n.Kind == SheetSetRowKind.Sheet);
    }

    /// <summary>Dialog colors for subset nesting depth (UI only — not stored in DST).</summary>
    public static class SubsetLevelColors
    {
        public static readonly Color Level1 = Color.FromArgb(0xDB, 0xEA, 0xFE); // blue
        public static readonly Color Level2 = Color.FromArgb(0xDC, 0xFC, 0xE7); // green
        public static readonly Color Level3 = Color.FromArgb(0xFF, 0xED, 0xD5); // orange
        public static readonly Color Level4Plus = Color.FromArgb(0xF3, 0xE8, 0xFF); // purple

        public static Color ForLevel(int level)
        {
            if (level <= 1) return Level1;
            if (level == 2) return Level2;
            if (level == 3) return Level3;
            return Level4Plus;
        }

        /// <summary>Lighter tint for sheet rows under a subset.</summary>
        public static Color SheetTint(int level)
        {
            var c = ForLevel(Math.Max(1, level));
            return Color.FromArgb(
                (c.R + 510) / 3,
                (c.G + 510) / 3,
                (c.B + 510) / 3);
        }

        public static Color SelectionForLevel(int level)
        {
            var c = ForLevel(Math.Max(1, level));
            return Color.FromArgb(
                Math.Max(0, c.R - 40),
                Math.Max(0, c.G - 40),
                Math.Max(0, c.B - 40));
        }
    }

    /// <summary>
    /// Creates native AutoCAD .dst files through the in-process Sheet Set COM API.
    /// Uses typed AcSm* interop (vtable) rather than IDispatch late binding —
    /// AcSmComponents on AutoCAD 2024+ does not implement IDispatch.
    /// </summary>
    public static class SheetSetService
    {
        public static void CreateOrReplace(
            string dstPath,
            string sheetSetName,
            IList<SheetSetEntry> entries)
        {
            if (string.IsNullOrWhiteSpace(dstPath))
                throw new ArgumentException("A .dst output path is required.", nameof(dstPath));
            if (!dstPath.EndsWith(".dst", StringComparison.OrdinalIgnoreCase))
                dstPath += ".dst";

            var ordered = (entries ?? Array.Empty<SheetSetEntry>())
                .Where(e => e != null)
                .ToList();
            // Empty table is allowed — clears sheets from the DST on rebuild.

            // Rebuild from the dialog table (subset headers + sheet order).
            // Close SSM's hold on the file first so overwrite can succeed.
            try
            {
                RebuildFromTable(dstPath, sheetSetName, ordered);
            }
            catch (Exception ex)
            {
                SheetSetAutoLog.WriteException(
                    null, null, "PLSHEETSET", "CreateOrReplace failed", ex, dstPath);
                throw;
            }
        }

        public class AutoSyncResult
        {
            public bool Ok { get; set; }
            public string Message { get; set; }
            public string DstPath { get; set; }
            public int SheetCount { get; set; }
        }

        /// <summary>
        /// Silent Create/Update of the default DST for the active DWG after PLAYOUT.
        /// Preserves existing subset tree when the DST already exists; appends new layouts;
        /// refreshes Title from DrawingName for sheets of this DWG.
        /// </summary>
        public static AutoSyncResult TryAutoSyncFromLayouts(
            string dwgPath,
            IList<PrintableLayout> layouts,
            IDictionary<string, string> drawingNames)
        {
            var result = new AutoSyncResult();
            if (string.IsNullOrWhiteSpace(dwgPath))
            {
                result.Message = "DWG path missing — save the drawing before auto sheet set.";
                return result;
            }

            string template = Config.Instance.TemplateLayout ?? "";
            var layoutList = (layouts ?? Array.Empty<PrintableLayout>())
                .Where(l => l != null && !string.IsNullOrWhiteSpace(l.Name))
                .Where(l => !l.Name.Equals(template, StringComparison.OrdinalIgnoreCase))
                .GroupBy(l => l.Name, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .ToList();
            if (layoutList.Count == 0)
            {
                result.Message = "No printable layouts to sync.";
                return result;
            }

            var dstPath = PublishPaths.DefaultDstPath(dwgPath);
            result.DstPath = dstPath;
            var names = drawingNames ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            string ResolveTitle(string layoutName, string fallback = null)
            {
                if (!string.IsNullOrWhiteSpace(layoutName)
                    && names.TryGetValue(layoutName, out var dn)
                    && !string.IsNullOrWhiteSpace(dn))
                    return dn.Trim();
                if (!string.IsNullOrWhiteSpace(fallback)) return fallback.Trim();
                return layoutName ?? "";
            }

            var layoutByName = layoutList
                .ToDictionary(l => l.Name, l => l, StringComparer.OrdinalIgnoreCase);
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var entries = new List<SheetSetEntry>();
            string sheetSetName = Path.GetFileNameWithoutExtension(dwgPath) ?? "SheetSet";

            var existing = TryRead(dstPath);
            if (existing != null && existing.Nodes.Count > 0)
            {
                if (!string.IsNullOrWhiteSpace(existing.SheetSetName))
                    sheetSetName = existing.SheetSetName;

                foreach (var node in existing.Nodes)
                {
                    if (node == null) continue;
                    if (node.Kind == SheetSetRowKind.Subset)
                    {
                        entries.Add(new SheetSetEntry
                        {
                            Kind = SheetSetRowKind.Subset,
                            Include = false,
                            SubsetName = node.SubsetName ?? node.Title,
                            SubsetLevel = Math.Max(1, node.SubsetLevel),
                            Title = node.Title ?? node.SubsetName,
                            SheetNumber = "",
                            DwgPath = dwgPath,
                        });
                        continue;
                    }

                    layoutByName.TryGetValue(node.LayoutName ?? "", out var layout);
                    string nodeDwg = !string.IsNullOrWhiteSpace(node.DwgPath) ? node.DwgPath : dwgPath;
                    bool sameDwg = PathsEqual(nodeDwg, dwgPath);

                    if (layout != null) used.Add(layout.Name);
                    if (layout == null && !string.IsNullOrWhiteSpace(node.LayoutName))
                        layout = new PrintableLayout { Name = node.LayoutName };

                    string title = sameDwg
                        ? ResolveTitle(node.LayoutName, node.Title)
                        : (!string.IsNullOrWhiteSpace(node.Title) ? node.Title : node.LayoutName);

                    entries.Add(new SheetSetEntry
                    {
                        Kind = SheetSetRowKind.Sheet,
                        Include = layout != null && !string.IsNullOrWhiteSpace(nodeDwg),
                        SubsetName = node.SubsetName,
                        SubsetLevel = Math.Max(0, node.SubsetLevel),
                        SheetNumber = !string.IsNullOrWhiteSpace(node.Number)
                            ? node.Number
                            : (node.LayoutName ?? ""),
                        Title = title ?? "",
                        Revision = node.Revision ?? "",
                        DwgPath = nodeDwg,
                        Layout = layout,
                    });
                }

                foreach (var layout in layoutList)
                {
                    if (used.Contains(layout.Name)) continue;
                    entries.Add(new SheetSetEntry
                    {
                        Kind = SheetSetRowKind.Sheet,
                        Include = true,
                        SheetNumber = layout.Name,
                        Title = ResolveTitle(layout.Name),
                        Revision = "",
                        DwgPath = dwgPath,
                        Layout = layout,
                    });
                }
            }
            else
            {
                foreach (var layout in layoutList)
                {
                    entries.Add(new SheetSetEntry
                    {
                        Kind = SheetSetRowKind.Sheet,
                        Include = true,
                        SheetNumber = layout.Name,
                        Title = ResolveTitle(layout.Name),
                        Revision = "",
                        DwgPath = dwgPath,
                        Layout = layout,
                    });
                }
            }

            try
            {
                CreateOrReplace(dstPath, sheetSetName, entries);
                result.Ok = true;
                result.SheetCount = entries.Count(e => e != null && !e.IsSubset);
                result.Message = $"Auto sheet set: {result.SheetCount} sheet(s) → {dstPath}";
            }
            catch (Exception ex)
            {
                result.Message = "Auto sheet set failed: " + ex.Message;
            }
            return result;
        }

        public class FolderImportWriteResult
        {
            public bool Ok { get; set; }
            public string Message { get; set; }
            public int SubsetsCreated { get; set; }
            public int SheetsAdded { get; set; }
            public int SheetsReplaced { get; set; }
        }

        /// <summary>
        /// Merge a scanned folder tree into an existing (or new) DST under
        /// <paramref name="parentSubsetPath"/> (empty = sheet-set root).
        /// Conflicts (same normalized DWG path + layout name) are removed then re-imported.
        /// </summary>
        public static FolderImportWriteResult ImportFolderTree(
            string dstPath,
            string sheetSetName,
            string parentSubsetPath,
            ImportFolderNode root)
        {
            var result = new FolderImportWriteResult();
            if (root == null || !root.HasContent)
            {
                result.Message = "Nothing to import.";
                return result;
            }
            if (string.IsNullOrWhiteSpace(dstPath))
            {
                result.Message = "DST path is missing.";
                return result;
            }
            if (!dstPath.EndsWith(".dst", StringComparison.OrdinalIgnoreCase))
                dstPath += ".dst";

#if !ACSM_INTEROP
            result.Message = "Sheet Set interop unavailable.";
            return result;
#else
            var dir = Path.GetDirectoryName(dstPath);
            if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);

            IAcSmSheetSetMgr manager = null;
            IAcSmDatabase database = null;
            bool weOpened = false;
            bool locked = false;
            bool commit = false;
            bool createdNew = false;
            try
            {
                manager = (IAcSmSheetSetMgr)CreateComObject("AcSmSheetSetMgr");
                if (File.Exists(dstPath))
                {
                    database = AcquireDatabase(manager, dstPath, "import-folder", out weOpened);
                }
                else
                {
                    database = manager.CreateDatabase(dstPath, "", true);
                    createdNew = true;
                    weOpened = true;
                }

                if (database == null)
                {
                    result.Message = "Could not open or create DST.";
                    return result;
                }

                database.LockDb(database);
                locked = true;

                IAcSmSheetSet sheetSet = database.GetSheetSet();
                if (sheetSet == null)
                {
                    result.Message = "Sheet set is null.";
                    return result;
                }

                if (createdNew || string.IsNullOrWhiteSpace(sheetSet.GetName()))
                {
                    sheetSet.SetName(
                        string.IsNullOrWhiteSpace(sheetSetName)
                            ? Path.GetFileNameWithoutExtension(dstPath)
                            : sheetSetName.Trim());
                }

                string newSheetFolder = dir;
                ApplyNewSheetLocation(sheetSet, sheetSet, newSheetFolder);
                try { sheetSet.SetPromptForDwt(true); } catch { }

                // Remove conflicting sheets (DWG+Layout) anywhere in the DST first.
                var importKeys = new HashSet<string>(StringComparer.Ordinal);
                foreach (var pair in root.EnumerateLayouts())
                    importKeys.Add(SheetSetFolderImport.SheetKey(pair.DwgPath, pair.Layout.LayoutName));

                result.SheetsReplaced = RemoveSheetsByKeys(sheetSet, importKeys);

                object parentHost = ResolveOrCreateSubsetPath(sheetSet, parentSubsetPath, newSheetFolder, result);
                if (parentHost == null)
                {
                    result.Message = "Could not resolve parent subset.";
                    return result;
                }

                // Selected folder becomes a subset under the parent (merge by name).
                MergeFolderNode(parentHost, root, newSheetFolder, result);

                try { database.UpdateInMemoryDwgHints(); } catch { }
                try { sheetSet.UpdateInMemoryDwgHints(); } catch { }
                ReleaseCom(sheetSet);
                commit = true;
                result.Ok = true;
                result.Message =
                    $"Import done. Subsets touched/created: {result.SubsetsCreated}, "
                    + $"sheets added: {result.SheetsAdded}, replaced: {result.SheetsReplaced}.";
                return result;
            }
            catch (COMException ex)
            {
                result.Message = WrapCom(ex, "ImportFolderTree", dstPath).Message;
                return result;
            }
            catch (Exception ex)
            {
                result.Message = ex.Message;
                return result;
            }
            finally
            {
                ReleaseDatabase(manager, database, weOpened, locked, commit);
            }
#endif
        }

        /// <summary>
        /// Import selected drawings' layouts as sheets directly under
        /// <paramref name="parentSubsetPath"/> (no subset per file).
        /// </summary>
        public static FolderImportWriteResult ImportDrawingFiles(
            string dstPath,
            string sheetSetName,
            string parentSubsetPath,
            IList<ImportDrawingFile> drawings)
        {
            var result = new FolderImportWriteResult();
            var list = (drawings ?? Array.Empty<ImportDrawingFile>())
                .Where(d => d != null
                    && !string.IsNullOrWhiteSpace(d.DwgPath)
                    && d.Layouts != null
                    && d.Layouts.Count > 0)
                .ToList();
            if (list.Count == 0)
            {
                result.Message = "Nothing to import.";
                return result;
            }
            if (string.IsNullOrWhiteSpace(dstPath))
            {
                result.Message = "DST path is missing.";
                return result;
            }
            if (!dstPath.EndsWith(".dst", StringComparison.OrdinalIgnoreCase))
                dstPath += ".dst";

#if !ACSM_INTEROP
            result.Message = "Sheet Set interop unavailable.";
            return result;
#else
            var dir = Path.GetDirectoryName(dstPath);
            if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);

            IAcSmSheetSetMgr manager = null;
            IAcSmDatabase database = null;
            bool weOpened = false;
            bool locked = false;
            bool commit = false;
            bool createdNew = false;
            try
            {
                manager = (IAcSmSheetSetMgr)CreateComObject("AcSmSheetSetMgr");
                if (File.Exists(dstPath))
                {
                    database = AcquireDatabase(manager, dstPath, "import-dwg", out weOpened);
                }
                else
                {
                    database = manager.CreateDatabase(dstPath, "", true);
                    createdNew = true;
                    weOpened = true;
                }

                if (database == null)
                {
                    result.Message = "Could not open or create DST.";
                    return result;
                }

                database.LockDb(database);
                locked = true;

                IAcSmSheetSet sheetSet = database.GetSheetSet();
                if (sheetSet == null)
                {
                    result.Message = "Sheet set is null.";
                    return result;
                }

                if (createdNew || string.IsNullOrWhiteSpace(sheetSet.GetName()))
                {
                    sheetSet.SetName(
                        string.IsNullOrWhiteSpace(sheetSetName)
                            ? Path.GetFileNameWithoutExtension(dstPath)
                            : sheetSetName.Trim());
                }

                string newSheetFolder = dir;
                ApplyNewSheetLocation(sheetSet, sheetSet, newSheetFolder);
                try { sheetSet.SetPromptForDwt(true); } catch { }

                var importKeys = new HashSet<string>(StringComparer.Ordinal);
                foreach (var d in list)
                {
                    foreach (var layout in d.Layouts ?? Enumerable.Empty<ImportLayoutSheet>())
                    {
                        if (layout == null || string.IsNullOrWhiteSpace(layout.LayoutName)) continue;
                        importKeys.Add(SheetSetFolderImport.SheetKey(d.DwgPath, layout.LayoutName));
                    }
                }

                result.SheetsReplaced = RemoveSheetsByKeys(sheetSet, importKeys);

                object parentHost = ResolveOrCreateSubsetPath(
                    sheetSet, parentSubsetPath, newSheetFolder, result);
                if (parentHost == null)
                {
                    result.Message = "Could not resolve parent subset.";
                    return result;
                }

                foreach (var drawing in list)
                {
                    if (!File.Exists(drawing.DwgPath)) continue;
                    foreach (var layout in drawing.Layouts ?? Enumerable.Empty<ImportLayoutSheet>())
                    {
                        if (layout == null || string.IsNullOrWhiteSpace(layout.LayoutName)) continue;
                        if (ImportOneSheet(parentHost, drawing.DwgPath, layout, result))
                            result.SheetsAdded++;
                    }
                }

                try { database.UpdateInMemoryDwgHints(); } catch { }
                try { sheetSet.UpdateInMemoryDwgHints(); } catch { }
                ReleaseCom(sheetSet);
                commit = true;
                result.Ok = true;
                result.Message =
                    $"Import done. Sheets added: {result.SheetsAdded}, replaced: {result.SheetsReplaced}.";
                return result;
            }
            catch (COMException ex)
            {
                result.Message = WrapCom(ex, "ImportFolderTree", dstPath).Message;
                return result;
            }
            catch (Exception ex)
            {
                result.Message = ex.Message;
                return result;
            }
            finally
            {
                ReleaseDatabase(manager, database, weOpened, locked, commit);
            }
#endif
        }

        /// <summary>
        /// Writes the flat table into the DST: subset header rows create subsets, and the
        /// sheets after a header go into it (or into the root until the next header).
        /// <para>
        /// A path with no .dst yet goes through <c>CreateDatabase</c> once; an existing .dst
        /// is edited in place through <c>OpenDatabase</c>. We deliberately do NOT build into
        /// a temp DST and swap it over the target: the swap ran
        /// <c>UpdateInMemoryDwgHints</c> while the file still had its temp name, so the
        /// title-block Fields pointed at a path that no longer existed (#### stayed), and it
        /// meant File.Delete on a .dst that Sheet Set Manager still owned.
        /// </para>
        /// </summary>
        private static void RebuildFromTable(
            string dstPath,
            string sheetSetName,
            IList<SheetSetEntry> ordered)
        {
#if !ACSM_INTEROP
            throw MissingInterop();
#else
            var dir = Path.GetDirectoryName(dstPath);
            if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);

            bool exists = File.Exists(dstPath);
            string mode = exists ? "in-place" : "create-new";
            SheetSetAutoLog.WriteStep(null, null, dstPath, $"RebuildFromTable start [{mode}]",
                $"entries={ordered?.Count ?? 0} dstBytes={SafeLength(dstPath)} "
                + $"dst$={File.Exists(dstPath + "$")}");

            if (exists)
                RewriteExistingDst(dstPath, sheetSetName, ordered);
            else
                CreateNewDst(dstPath, sheetSetName, ordered);

            SheetSetAutoLog.WriteStep(null, null, dstPath, $"RebuildFromTable OK [{mode}]",
                $"dstBytes={SafeLength(dstPath)}");
#endif
        }

        /// <summary>
        /// First write for a path that has no .dst yet — the only place <c>CreateDatabase</c>
        /// is allowed to run. If the write fails the half-built file is removed; it did not
        /// exist beforehand, so deleting it cannot destroy user data.
        /// </summary>
        private static void CreateNewDst(
            string dstPath,
            string sheetSetName,
            IList<SheetSetEntry> ordered)
        {
            IAcSmSheetSetMgr manager = null;
            IAcSmDatabase database = null;
            bool locked = false;
            bool commit = false;
            try
            {
                manager = (IAcSmSheetSetMgr)CreateComObject("AcSmSheetSetMgr");
                SheetSetAutoLog.WriteStep(null, null, dstPath, "CreateDatabase [create-new]",
                    $"exists={File.Exists(dstPath)}");
                database = manager.CreateDatabase(dstPath, "", true);
                if (database == null)
                    throw new InvalidOperationException("AutoCAD did not create the sheet set database.");

                database.LockDb(database);
                locked = true;

                IAcSmSheetSet sheetSet = database.GetSheetSet();
                if (sheetSet == null)
                    throw new InvalidOperationException("Sheet set is null after creating DST.");

                sheetSet.SetName(ResolveSheetSetName(dstPath, sheetSetName));

                string newSheetFolder = ResolveNewSheetFolder(dstPath, ordered);
                ApplyNewSheetLocation(sheetSet, sheetSet, newSheetFolder);
                try { sheetSet.SetPromptForDwt(true); } catch { }

                PopulateSheetSet(sheetSet, ordered, newSheetFolder, dstPath);

                PushDwgHints(database, sheetSet, dstPath);
                ReleaseCom(sheetSet);
                commit = true;
            }
            catch (COMException ex)
            {
                // Label the method, not a guessed API — the old "CreateDatabase" label was
                // hard-coded and hid the fact that ImportSheet was the call being rejected.
                throw WrapCom(ex, "CreateNewDst", dstPath);
            }
            finally
            {
                if (database != null && locked)
                {
                    try { database.UnlockDb(database, commit); } catch { }
                }
                // We created this database, so we own it and must close it.
                if (manager != null && database != null)
                {
                    try { manager.Close((AcSmDatabase)database); } catch { }
                }
                ReleaseCom(database);
                ReleaseCom(manager);

                // Keep whatever CreateDatabase produced even when populating failed. The file
                // holds a valid (possibly empty) sheet set, so the next run takes the in-place
                // path instead of re-creating from scratch forever.
                if (!commit)
                {
                    SheetSetAutoLog.WriteStep(null, null, dstPath, "CreateNewDst failed",
                        $"leaving the file in place — dstBytes={SafeLength(dstPath)}");
                }
            }
        }

        /// <summary>
        /// Rewrites a .dst that already exists, in place: empty the sheet set, then re-add
        /// subsets/sheets from the table. No temp file, no File.Move, no File.Delete — an
        /// existing .dst is only ever touched through AcSm.
        /// <para>
        /// When this AutoCAD session already has the .dst open (Sheet Set Manager, or our own
        /// UI hold) we reuse that database object rather than opening a second handle on the
        /// same path.
        /// </para>
        /// </summary>
        private static void RewriteExistingDst(
            string dstPath,
            string sheetSetName,
            IList<SheetSetEntry> ordered)
        {
            IAcSmSheetSetMgr manager = null;
            IAcSmDatabase database = null;
            bool weOpened = false;
            bool locked = false;
            bool commit = false;
            try
            {
                manager = (IAcSmSheetSetMgr)CreateComObject("AcSmSheetSetMgr");
                database = AcquireDatabase(manager, dstPath, "in-place", out weOpened);
                if (database == null)
                    throw new InvalidOperationException(
                        "Could not open the sheet set database:\n\n" + dstPath);

                database.LockDb(database);
                locked = true;

                IAcSmSheetSet sheetSet = database.GetSheetSet();
                if (sheetSet == null)
                    throw new InvalidOperationException("Sheet set is null for:\n\n" + dstPath);

                sheetSet.SetName(ResolveSheetSetName(dstPath, sheetSetName));

                ClearSheetSetContents(sheetSet);
                int leftover = CountSheetsAndSubsets(sheetSet);
                SheetSetAutoLog.WriteStep(null, null, dstPath, "ClearSheetSetContents [in-place]",
                    $"sheets/subsets left={leftover}");
                if (leftover > 0)
                {
                    // Stop rather than re-add on top of the old tree — a half-cleared rebuild
                    // would duplicate every sheet.
                    throw new InvalidOperationException(
                        $"Could not empty the sheet set before rebuilding — {leftover} "
                        + "sheet(s)/subset(s) could not be removed.\n\n"
                        + "Close the sheet set in Sheet Set Manager (dropdown → Close), then "
                        + "Create / Update again.\n\n" + dstPath);
                }

                string newSheetFolder = ResolveNewSheetFolder(dstPath, ordered);
                ApplyNewSheetLocation(sheetSet, sheetSet, newSheetFolder);
                try { sheetSet.SetPromptForDwt(true); } catch { }

                PopulateSheetSet(sheetSet, ordered, newSheetFolder, dstPath);

                PushDwgHints(database, sheetSet, dstPath);
                ReleaseCom(sheetSet);
                commit = true;
            }
            catch (COMException ex)
            {
                throw WrapCom(ex, "RewriteExistingDst", dstPath);
            }
            finally
            {
                ReleaseDatabase(manager, database, weOpened, locked, commit);
            }
        }

        /// <summary>
        /// Adds the table's subset headers and sheets to <paramref name="sheetSet"/>.
        /// Shared by the create-new and in-place paths.
        /// </summary>
        private static void PopulateSheetSet(
            IAcSmSheetSet sheetSet,
            IList<SheetSetEntry> ordered,
            string newSheetFolder,
            string dstPath)
        {
            // Stack of open subsets by depth (index 0 = level 1).
            var subsetStack = new List<IAcSmSubset>();
            foreach (var entry in ordered)
            {
                if (entry == null) continue;

                if (entry.IsSubset)
                {
                    int level = Math.Max(1, entry.SubsetLevel);
                    while (subsetStack.Count >= level)
                    {
                        ReleaseCom(subsetStack[subsetStack.Count - 1]);
                        subsetStack.RemoveAt(subsetStack.Count - 1);
                    }

                    var name = (entry.Title ?? entry.SubsetName ?? "Subset").Trim();
                    if (string.IsNullOrWhiteSpace(name)) name = "Subset";

                    IAcSmSubset created;
                    if (subsetStack.Count == 0)
                    {
                        created = sheetSet.CreateSubset(name, "");
                    }
                    else
                    {
                        created = subsetStack[subsetStack.Count - 1].CreateSubset(name, "");
                    }

                    if (created == null)
                        throw new InvalidOperationException("CreateSubset returned null for: " + name);
                    ApplyNewSheetLocation(created, created, newSheetFolder);
                    try { created.SetPromptForDwt(true); } catch { }
                    subsetStack.Add(created);
                    continue;
                }

                if (entry.Layout == null) continue;
                if (string.IsNullOrWhiteSpace(entry.DwgPath) || !File.Exists(entry.DwgPath))
                    throw new FileNotFoundException(
                        "Save the source DWG before creating a sheet set.", entry.DwgPath);

                IAcSmSubset currentSubset = subsetStack.Count > 0
                    ? subsetStack[subsetStack.Count - 1]
                    : null;
                IAcSmPersist initOwner = currentSubset != null
                    ? (IAcSmPersist)currentSubset
                    : (IAcSmPersist)sheetSet;
                var layoutRef = (IAcSmAcDbLayoutReference)CreateComObject("AcSmAcDbLayoutReference");
                layoutRef.InitNew(initOwner);
                layoutRef.SetFileName(entry.DwgPath);
                layoutRef.SetName(entry.Layout.Name);

                IAcSmSheet sheet;
                try
                {
                    sheet = currentSubset != null
                        ? currentSubset.ImportSheet((AcSmAcDbLayoutReference)layoutRef)
                        : sheetSet.ImportSheet((AcSmAcDbLayoutReference)layoutRef);
                }
                catch (COMException)
                {
                    LogImportSheetFailure(entry, dstPath);
                    throw;
                }

                if (sheet == null)
                    throw new InvalidOperationException(
                        $"Could not import layout '{entry.Layout.Name}'.");

                sheet.SetNumber(entry.SheetNumber ?? "");
                sheet.SetTitle(
                    string.IsNullOrWhiteSpace(entry.Title)
                        ? entry.Layout.Name
                        : entry.Title.Trim());
                TrySetRevisionNumber(sheet, entry.Revision ?? "");

                try
                {
                    if (currentSubset != null)
                        currentSubset.InsertComponent(sheet, null);
                    else
                        sheetSet.InsertComponent(sheet, null);
                }
                catch (COMException)
                {
                    // Already in the set — safe to continue.
                }

                ReleaseCom(sheet);
                ReleaseCom(layoutRef);
            }

            foreach (var s in subsetStack) ReleaseCom(s);
            subsetStack.Clear();
        }

        /// <summary>
        /// Records everything needed to explain an ImportSheet rejection: AcSm resolves the
        /// layout through the .dwg as SAVED ON DISK, so a layout PLAYOUT just created but
        /// nobody saved is invisible to it, and the raw HRESULT says nothing about that.
        /// </summary>
        private static void LogImportSheetFailure(SheetSetEntry entry, string dstPath)
        {
            string layoutName = entry?.Layout?.Name ?? "";
            string dwgPath = entry?.DwgPath ?? "";
            var onDisk = ReadLayoutNamesFromDisk(dwgPath);

            string diskState = onDisk == null
                ? "unreadable"
                : (onDisk.Any(n => string.Equals(n, layoutName, StringComparison.OrdinalIgnoreCase))
                    ? "PRESENT"
                    : "MISSING");

            string dbmod = "?";
            try { dbmod = Convert.ToString(AcadApp.GetSystemVariable("DBMOD"), CultureInfo.InvariantCulture); }
            catch { }

            string savedAt = "?";
            try { if (File.Exists(dwgPath)) savedAt = File.GetLastWriteTime(dwgPath).ToString("HH:mm:ss"); }
            catch { }

            SheetSetAutoLog.WriteTagged(null, null, "PLSHEETSET-STEP", dstPath:dstPath, message:
                $"ImportSheet REJECTED | layout='{layoutName}' number='{entry?.SheetNumber}' "
                + $"layoutInSavedDwg={diskState} dwgSavedAt={savedAt} DBMOD={dbmod} "
                + $"(DBMOD!=0 means the drawing has unsaved changes) dwg={dwgPath}");

            if (onDisk != null)
            {
                SheetSetAutoLog.WriteTagged(null, null, "PLSHEETSET-STEP", dstPath:dstPath, message:
                    "  layouts in the saved .dwg: " + string.Join(", ", onDisk));
            }
        }

        /// <summary>
        /// Layout names in the .dwg as it sits ON DISK — deliberately not the in-memory
        /// document, because that is what AcSm reads. Null when the file cannot be read.
        /// </summary>
        private static List<string> ReadLayoutNamesFromDisk(string dwgPath)
        {
            if (string.IsNullOrWhiteSpace(dwgPath) || !File.Exists(dwgPath)) return null;

            Autodesk.AutoCAD.DatabaseServices.Database db = null;
            try
            {
                db = new Autodesk.AutoCAD.DatabaseServices.Database(false, true);
                db.ReadDwgFile(
                    dwgPath,
                    Autodesk.AutoCAD.DatabaseServices.FileOpenMode.OpenForReadAndReadShare,
                    true,
                    "");
                db.CloseInput(true);

                var names = new List<string>();
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var dict = (Autodesk.AutoCAD.DatabaseServices.DBDictionary)tr.GetObject(
                        db.LayoutDictionaryId,
                        Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                    foreach (var item in dict)
                    {
                        if (!string.Equals(item.Key, "Model", StringComparison.OrdinalIgnoreCase))
                            names.Add(item.Key);
                    }
                    tr.Commit();
                }
                return names;
            }
            catch
            {
                return null;
            }
            finally
            {
                try { db?.Dispose(); } catch { }
            }
        }

        /// <summary>
        /// Pushes the sheet-set association into the DWGs open in memory so title-block Fields
        /// (CurrentSheetNumber / CurrentSheetTitle) resolve instead of showing ####.
        /// Must run while the DST still sits at its final path.
        /// </summary>
        private static void PushDwgHints(
            IAcSmDatabase database, IAcSmSheetSet sheetSet, string dstPath)
        {
            bool dbOk = false, setOk = false;
            try { database?.UpdateInMemoryDwgHints(); dbOk = true; } catch { }
            try { sheetSet?.UpdateInMemoryDwgHints(); setOk = true; } catch { }
            SheetSetAutoLog.WriteStep(null, null, dstPath, "UpdateInMemoryDwgHints",
                $"database={dbOk} sheetSet={setOk}");
        }

        /// <summary>
        /// Sheets and subsets still owned by the sheet set — 0 means the tree is empty.
        /// <para>
        /// Counts only those two kinds on purpose. GetDirectlyOwnedObjects also hands back the
        /// sheet set's own structural objects (new-sheet location, custom property bag, sheet
        /// selection sets, view categories…), which ClearSheetSetContents must never remove —
        /// counting them made an empty sheet set look like it still had 10 items.
        /// </para>
        /// </summary>
        private static int CountSheetsAndSubsets(IAcSmSheetSet sheetSet)
        {
            if (sheetSet == null) return 0;
            object[] owned = null;
            try { sheetSet.GetDirectlyOwnedObjects(out owned); }
            catch { return 0; }
            if (owned == null) return 0;

            int count = 0;
            foreach (var item in owned)
            {
                if (item is IAcSmSheet || item is IAcSmSubset) count++;
                ReleaseCom(item);
            }
            return count;
        }

        private static string ResolveSheetSetName(string dstPath, string sheetSetName) =>
            string.IsNullOrWhiteSpace(sheetSetName)
                ? Path.GetFileNameWithoutExtension(dstPath)
                : sheetSetName.Trim();

        /// <summary>File size for logging; -1 when it cannot be read.</summary>
        private static long SafeLength(string path)
        {
            try { return File.Exists(path) ? new FileInfo(path).Length : 0L; }
            catch { return -1L; }
        }

        /// <summary>
        /// Removes all sheets/subsets from the sheet set so we can rebuild from the dialog table.
        /// Do NOT call <c>IAcSmPersist.Clear()</c> — that clears the component itself and
        /// leads to NullReferenceException on later CreateSubset / ImportSheet.
        /// </summary>
        private static void ClearSheetSetContents(IAcSmSheetSet sheetSet)
        {
            if (sheetSet == null) return;

            // Repeat until empty — removing can invalidate ownership lists mid-pass.
            for (int pass = 0; pass < 50; pass++)
            {
                object[] owned = null;
                try { sheetSet.GetDirectlyOwnedObjects(out owned); }
                catch { break; }
                if (owned == null || owned.Length == 0) break;

                bool removedAny = false;
                foreach (var item in owned)
                {
                    try
                    {
                        if (item is IAcSmSheet sheet)
                        {
                            try { sheetSet.RemoveSheet((AcSmSheet)sheet); removedAny = true; }
                            catch { }
                        }
                        else if (item is IAcSmSubset subset)
                        {
                            // Remove nested sheets first when possible.
                            try
                            {
                                object[] nested = null;
                                subset.GetDirectlyOwnedObjects(out nested);
                                if (nested != null)
                                {
                                    foreach (var n in nested)
                                    {
                                        try
                                        {
                                            if (n is IAcSmSheet nestedSheet)
                                                subset.RemoveSheet((AcSmSheet)nestedSheet);
                                        }
                                        catch { }
                                        finally { ReleaseCom(n); }
                                    }
                                }
                            }
                            catch { }

                            try { sheetSet.RemoveSubset((AcSmSubset)subset); removedAny = true; }
                            catch { }
                        }
                    }
                    finally
                    {
                        ReleaseCom(item);
                    }
                }

                if (!removedAny) break;
            }
        }

        /// <summary>
        /// Removes sheet entries whose layout name matches any of
        /// <paramref name="layoutNames"/> (case-insensitive). Subset headers are kept.
        /// </summary>
        public static bool TryRemoveSheetsByLayoutNames(
            string dstPath,
            IEnumerable<string> layoutNames,
            out string message)
        {
            message = "";
#if !ACSM_INTEROP
            message = "Sheet Set interop unavailable.";
            return false;
#else
            if (string.IsNullOrWhiteSpace(dstPath) || !File.Exists(dstPath))
            {
                message = "DST file not found.";
                return false;
            }

            var targets = new HashSet<string>(
                (layoutNames ?? Enumerable.Empty<string>())
                    .Where(n => !string.IsNullOrWhiteSpace(n))
                    .Select(n => n.Trim()),
                StringComparer.OrdinalIgnoreCase);
            if (targets.Count == 0)
            {
                message = "No layout names to remove.";
                return false;
            }

            IAcSmSheetSetMgr manager = null;
            IAcSmDatabase database = null;
            bool weOpened = false;
            bool locked = false;
            bool commit = false;
            int removed = 0;
            try
            {
                manager = (IAcSmSheetSetMgr)CreateComObject("AcSmSheetSetMgr");
                database = AcquireDatabase(manager, dstPath, "remove-sheets", out weOpened);
                if (database == null)
                {
                    message = "Could not open DST: " + dstPath;
                    return false;
                }

                database.LockDb(database);
                locked = true;

                var sheetSet = database.GetSheetSet();
                if (sheetSet == null)
                {
                    message = "Sheet set is null.";
                    return false;
                }

                removed += RemoveMatchingSheets(sheetSet, targets);

                try { database.UpdateInMemoryDwgHints(); } catch { }
                try { sheetSet.UpdateInMemoryDwgHints(); } catch { }
                ReleaseCom(sheetSet);
                commit = true;
                message = removed == 0
                    ? "No matching sheets found in DST."
                    : $"Removed {removed} sheet(s) from DST for deleted layout(s).";
                return removed > 0;
            }
            catch (COMException ex)
            {
                message = WrapCom(ex, "TryRemoveSheetsByLayoutNames", dstPath).Message;
                return false;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return false;
            }
            finally
            {
                ReleaseDatabase(manager, database, weOpened, locked, commit);
            }
#endif
        }

#if ACSM_INTEROP
        private static int RemoveMatchingSheets(
            IAcSmSheetSet sheetSet,
            HashSet<string> layoutNames)
        {
            if (sheetSet == null || layoutNames == null || layoutNames.Count == 0)
                return 0;

            // Collect (sheet, host) first — same ownership walk as TryRead.
            var victims = new List<Tuple<IAcSmSheet, object>>();

            void Consider(IAcSmSheet sheet, object host)
            {
                if (sheet == null) return;
                string layoutName = null;
                try
                {
                    var layout = sheet.GetLayout();
                    layoutName = layout?.GetName();
                    ReleaseCom(layout);
                }
                catch { }

                string number = null;
                try { number = sheet.GetNumber(); } catch { }

                bool match =
                    (!string.IsNullOrWhiteSpace(layoutName) && layoutNames.Contains(layoutName))
                    || (!string.IsNullOrWhiteSpace(number) && layoutNames.Contains(number));

                if (match)
                    victims.Add(Tuple.Create(sheet, host));
                else
                    ReleaseCom(sheet);
            }

            object[] owned = null;
            try { sheetSet.GetDirectlyOwnedObjects(out owned); } catch { owned = null; }
            if (owned != null)
            {
                foreach (var item in owned)
                {
                    try
                    {
                        if (item is IAcSmSheet sheet)
                        {
                            Consider(sheet, sheetSet);
                            continue;
                        }
                        if (item is IAcSmSubset subset)
                        {
                            object[] nested = null;
                            try { subset.GetDirectlyOwnedObjects(out nested); } catch { nested = null; }
                            if (nested != null)
                            {
                                foreach (var n in nested)
                                {
                                    try
                                    {
                                        if (n is IAcSmSheet nestedSheet)
                                            Consider(nestedSheet, subset);
                                        else
                                            ReleaseCom(n);
                                    }
                                    catch { ReleaseCom(n); }
                                }
                            }

                            // Fallback: sheet enumerator on subset (some DST shapes).
                            try
                            {
                                var en = subset.GetSheetEnumerator();
                                if (en != null)
                                {
                                    IAcSmComponent next;
                                    while ((next = en.Next()) != null)
                                    {
                                        try
                                        {
                                            if (next is IAcSmSheet nestedSheet)
                                                Consider(nestedSheet, subset);
                                            else
                                                ReleaseCom(next);
                                        }
                                        catch { ReleaseCom(next); }
                                    }
                                    ReleaseCom(en);
                                }
                            }
                            catch { }

                            ReleaseCom(subset);
                            continue;
                        }
                        ReleaseCom(item);
                    }
                    catch { ReleaseCom(item); }
                }
            }

            int removed = 0;
            foreach (var pair in victims)
            {
                var sheet = pair.Item1;
                var host = pair.Item2;
                try
                {
                    if (host is IAcSmSubset subset)
                        subset.RemoveSheet((AcSmSheet)sheet);
                    else if (host is IAcSmSheetSet root)
                        root.RemoveSheet((AcSmSheet)sheet);
                    else
                    {
                        // Last resort: ask the sheet who owns it.
                        try
                        {
                            var owner = sheet.GetOwner();
                            if (owner is IAcSmSubset ownSub)
                                ownSub.RemoveSheet((AcSmSheet)sheet);
                            else if (owner is IAcSmSheetSet ownRoot)
                                ownRoot.RemoveSheet((AcSmSheet)sheet);
                            ReleaseCom(owner);
                        }
                        catch { }
                    }
                    removed++;
                }
                catch { }
                finally
                {
                    ReleaseCom(sheet);
                }
            }

            return removed;
        }
#endif

        /// <summary>
        /// Reads nested subsets + sheet Number/Title from an existing DST.
        /// Returns null when the file is missing or cannot be opened.
        /// </summary>
        public static SheetSetReadResult TryRead(string dstPath)
        {
#if !ACSM_INTEROP
            return null;
#else
            if (string.IsNullOrWhiteSpace(dstPath) || !File.Exists(dstPath))
                return null;

            // Reading while SSM has the DST open is fine — we reuse its database object.
            // If that fails, the caller falls back to seeding from the model.
            IAcSmSheetSetMgr manager = null;
            IAcSmDatabase database = null;
            bool weOpened = false;
            bool locked = false;
            try
            {
                manager = (IAcSmSheetSetMgr)CreateComObject("AcSmSheetSetMgr");
                database = AcquireDatabase(manager, dstPath, "read", out weOpened);
                if (database == null) return null;

                database.LockDb(database);
                locked = true;

                var sheetSet = database.GetSheetSet();
                var result = new SheetSetReadResult
                {
                    SheetSetName = sheetSet?.GetName()
                };

                if (!TryReadOwnedTree(sheetSet, parentSubsetName: null, depth: 0, result))
                    ReadLevel(sheetSet, subsetName: null, result, depth: 0);

                ReleaseCom(sheetSet);
                return result;
            }
            catch
            {
                return null;
            }
            finally
            {
                // Read-only: commit true keeps AcSm's own bookkeeping happy, nothing was edited.
                ReleaseDatabase(manager, database, weOpened, locked, commit: true);
            }
#endif
        }

#if ACSM_INTEROP
        /// <summary>
        /// Preferred: GetDirectlyOwnedObjects preserves subset order and nesting.
        /// </summary>
        private static bool TryReadOwnedTree(
            object container,
            string parentSubsetName,
            int depth,
            SheetSetReadResult result)
        {
            if (container == null || result == null) return false;
            object[] owned = null;
            try
            {
                if (container is IAcSmSheetSet sheetSet)
                    sheetSet.GetDirectlyOwnedObjects(out owned);
                else if (container is IAcSmSubset subset)
                    subset.GetDirectlyOwnedObjects(out owned);
                else
                    return false;
            }
            catch
            {
                return false;
            }

            if (owned == null || owned.Length == 0) return false;

            bool any = false;
            foreach (var item in owned)
            {
                try
                {
                    if (item is IAcSmSubset childSubset)
                    {
                        any = true;
                        int level = depth + 1;
                        var name = childSubset.GetName() ?? "Subset";
                        result.Nodes.Add(new SheetSetNodeInfo
                        {
                            Kind = SheetSetRowKind.Subset,
                            SubsetName = name,
                            Title = name,
                            SubsetLevel = level,
                        });
                        // Recurse; ignore bool — empty nested subset still counts as structure.
                        TryReadOwnedTree(childSubset, name, level, result);
                        continue;
                    }

                        if (item is IAcSmSheet sheet)
                        {
                            string layoutName = null;
                            string dwgPath = null;
                            try
                            {
                                var layout = sheet.GetLayout();
                                layoutName = layout?.GetName();
                                try { dwgPath = layout?.ResolveFileName(); } catch { }
                                if (string.IsNullOrWhiteSpace(dwgPath))
                                {
                                    try { dwgPath = layout?.GetFileName(); } catch { }
                                }
                                ReleaseCom(layout);
                            }
                            catch { }
                            if (string.IsNullOrWhiteSpace(layoutName)) continue;
                            any = true;
                            result.Nodes.Add(new SheetSetNodeInfo
                            {
                                Kind = SheetSetRowKind.Sheet,
                                SubsetName = parentSubsetName,
                                SubsetLevel = depth,
                                LayoutName = layoutName,
                                DwgPath = dwgPath,
                                Number = sheet.GetNumber() ?? "",
                                Title = sheet.GetTitle() ?? "",
                                Revision = TryGetRevisionNumber(sheet),
                            });
                        }
                }
                finally
                {
                    ReleaseCom(item);
                }
            }

            return any || result.Nodes.Count > 0;
        }

        private static void ReadLevel(
            object container,
            string subsetName,
            SheetSetReadResult result,
            int depth)
        {
            // Fallback when GetDirectlyOwnedObjects is unavailable.
            IAcSmEnumComponent enumerator = null;
            try
            {
                if (container is IAcSmSheetSet ss)
                    enumerator = ss.GetSheetEnumerator();
                else if (container is IAcSmSubset sub)
                    enumerator = sub.GetSheetEnumerator();
            }
            catch { }

            if (enumerator == null) return;

            try
            {
                IAcSmComponent next;
                while ((next = enumerator.Next()) != null)
                {
                    try
                    {
                        if (next is IAcSmSubset nested)
                        {
                            var name = nested.GetName() ?? "Subset";
                            int level = depth + 1;
                            result.Nodes.Add(new SheetSetNodeInfo
                            {
                                Kind = SheetSetRowKind.Subset,
                                SubsetName = name,
                                Title = name,
                                SubsetLevel = level,
                            });
                            ReadLevel(nested, name, result, level);
                            continue;
                        }

                        if (next is IAcSmSheet sheet)
                        {
                            string layoutName = null;
                            string dwgPath = null;
                            try
                            {
                                var layout = sheet.GetLayout();
                                layoutName = layout?.GetName();
                                try { dwgPath = layout?.ResolveFileName(); } catch { }
                                if (string.IsNullOrWhiteSpace(dwgPath))
                                {
                                    try { dwgPath = layout?.GetFileName(); } catch { }
                                }
                                ReleaseCom(layout);
                            }
                            catch { }

                            if (string.IsNullOrWhiteSpace(layoutName))
                                continue;

                            result.Nodes.Add(new SheetSetNodeInfo
                            {
                                Kind = SheetSetRowKind.Sheet,
                                SubsetName = subsetName,
                                SubsetLevel = depth,
                                LayoutName = layoutName,
                                DwgPath = dwgPath,
                                Number = sheet.GetNumber() ?? "",
                                Title = sheet.GetTitle() ?? "",
                                Revision = TryGetRevisionNumber(sheet),
                            });
                        }
                    }
                    finally
                    {
                        ReleaseCom(next);
                    }
                }
            }
            finally
            {
                ReleaseCom(enumerator);
            }
        }

        private static string TryGetRevisionNumber(IAcSmSheet sheet)
        {
            if (sheet == null) return "";
            try
            {
                if (sheet is IAcSmSheet2 sheet2)
                    return sheet2.GetRevisionNumber() ?? "";
            }
            catch { }
            return "";
        }

        private static void TrySetRevisionNumber(IAcSmSheet sheet, string revision)
        {
            if (sheet == null) return;
            try
            {
                if (sheet is IAcSmSheet2 sheet2)
                    sheet2.SetRevisionNumber(revision ?? "");
            }
            catch
            {
                // Older AcSm builds without IAcSmSheet2 — Number/Title still work.
            }
        }

        private static string ResolveNewSheetFolder(string dstPath, IList<SheetSetEntry> ordered)
        {
            foreach (var entry in ordered)
            {
                if (entry == null || entry.IsSubset) continue;
                if (string.IsNullOrWhiteSpace(entry.DwgPath)) continue;
                try
                {
                    var dir = Path.GetDirectoryName(entry.DwgPath);
                    if (!string.IsNullOrWhiteSpace(dir))
                        return dir;
                }
                catch { }
            }

            try
            {
                var dstDir = Path.GetDirectoryName(dstPath);
                if (!string.IsNullOrWhiteSpace(dstDir))
                    return dstDir;
            }
            catch { }

            return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        }

        private static void ApplyNewSheetLocation(
            IAcSmPersist owner,
            object target,
            string folder)
        {
            if (owner == null || target == null || string.IsNullOrWhiteSpace(folder))
                return;

            try
            {
                Directory.CreateDirectory(folder);
            }
            catch { }

            IAcSmFileReference fileRef = null;
            try
            {
                fileRef = (IAcSmFileReference)CreateComObject("AcSmFileReference");
                fileRef.InitNew(owner);
                fileRef.SetFileName(folder);

                if (target is IAcSmSheetSet sheetSet)
                    sheetSet.SetNewSheetLocation(fileRef);
                else if (target is IAcSmSubset subset)
                    subset.SetNewSheetLocation(fileRef);
            }
            catch
            {
                // Non-fatal: Import Layout as Sheet still works without this.
            }
            finally
            {
                ReleaseCom(fileRef);
            }
        }

        private static int RemoveSheetsByKeys(IAcSmSheetSet sheetSet, HashSet<string> keys)
        {
            if (sheetSet == null || keys == null || keys.Count == 0) return 0;
            int removed = 0;
            RemoveSheetsByKeysRecursive(sheetSet, keys, ref removed);
            return removed;
        }

        private static void RemoveSheetsByKeysRecursive(object host, HashSet<string> keys, ref int removed)
        {
            if (host == null) return;
            object[] owned = null;
            try
            {
                if (host is IAcSmSheetSet ss) ss.GetDirectlyOwnedObjects(out owned);
                else if (host is IAcSmSubset sub) sub.GetDirectlyOwnedObjects(out owned);
            }
            catch { return; }
            if (owned == null) return;

            // Snapshot first — mutating while enumerating is unsafe.
            var list = owned.Where(o => o != null).ToList();
            foreach (var item in list)
            {
                try
                {
                    if (item is IAcSmSubset nested)
                    {
                        RemoveSheetsByKeysRecursive(nested, keys, ref removed);
                        continue;
                    }

                    if (item is IAcSmSheet sheet)
                    {
                        if (!TryGetSheetKey(sheet, out var key) || !keys.Contains(key))
                            continue;
                        try
                        {
                            if (host is IAcSmSubset ownerSub)
                                ownerSub.RemoveSheet((AcSmSheet)sheet);
                            else if (host is IAcSmSheetSet ownerSs)
                                ownerSs.RemoveSheet((AcSmSheet)sheet);
                            removed++;
                        }
                        catch { }
                    }
                }
                finally
                {
                    ReleaseCom(item);
                }
            }
        }

        private static bool TryGetSheetKey(IAcSmSheet sheet, out string key)
        {
            key = null;
            if (sheet == null) return false;
            IAcSmAcDbLayoutReference layout = null;
            try
            {
                layout = sheet.GetLayout();
                if (layout == null) return false;
                string file = null;
                try { file = layout.ResolveFileName(); } catch { }
                if (string.IsNullOrWhiteSpace(file))
                {
                    try { file = layout.GetFileName(); } catch { }
                }
                var layoutName = layout.GetName() ?? "";
                if (string.IsNullOrWhiteSpace(file) || string.IsNullOrWhiteSpace(layoutName))
                    return false;
                key = SheetSetFolderImport.SheetKey(file, layoutName);
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(layout);
            }
        }

        /// <summary>
        /// parentSubsetPath like "Kientruc/test" (empty = sheet set root).
        /// Creates missing subsets along the path.
        /// </summary>
        private static object ResolveOrCreateSubsetPath(
            IAcSmSheetSet sheetSet,
            string parentSubsetPath,
            string newSheetFolder,
            FolderImportWriteResult stats)
        {
            object host = sheetSet;
            if (string.IsNullOrWhiteSpace(parentSubsetPath))
                return host;

            var parts = parentSubsetPath
                .Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(p => p.Trim())
                .Where(p => p.Length > 0)
                .ToArray();

            foreach (var part in parts)
            {
                var found = FindChildSubset(host, part);
                if (found == null)
                {
                    if (host is IAcSmSheetSet ss)
                        found = ss.CreateSubset(part, "");
                    else if (host is IAcSmSubset sub)
                        found = sub.CreateSubset(part, "");
                    if (found != null)
                    {
                        stats.SubsetsCreated++;
                        ApplyNewSheetLocation((IAcSmPersist)found, found, newSheetFolder);
                        try { found.SetPromptForDwt(true); } catch { }
                    }
                }
                if (found == null) return null;
                host = found;
            }

            return host;
        }

        private static IAcSmSubset FindChildSubset(object host, string name)
        {
            if (host == null || string.IsNullOrWhiteSpace(name)) return null;
            object[] owned = null;
            try
            {
                if (host is IAcSmSheetSet ss) ss.GetDirectlyOwnedObjects(out owned);
                else if (host is IAcSmSubset sub) sub.GetDirectlyOwnedObjects(out owned);
            }
            catch { return null; }
            if (owned == null) return null;

            foreach (var item in owned)
            {
                var subset = item as IAcSmSubset;
                if (subset != null)
                {
                    var n = subset.GetName() ?? "";
                    if (string.Equals(n, name, StringComparison.OrdinalIgnoreCase))
                        return subset; // caller keeps this COM ref for the merge pass
                }
                ReleaseCom(item);
            }
            return null;
        }

        private static void MergeFolderNode(
            object parentHost,
            ImportFolderNode node,
            string newSheetFolder,
            FolderImportWriteResult stats)
        {
            if (parentHost == null || node == null || !node.HasContent) return;

            var name = string.IsNullOrWhiteSpace(node.Name) ? "Subset" : node.Name.Trim();
            var subset = FindChildSubset(parentHost, name);
            if (subset == null)
            {
                if (parentHost is IAcSmSheetSet ss)
                    subset = ss.CreateSubset(name, "");
                else if (parentHost is IAcSmSubset sub)
                    subset = sub.CreateSubset(name, "");
                if (subset != null)
                {
                    stats.SubsetsCreated++;
                    ApplyNewSheetLocation(subset, subset, newSheetFolder);
                    try { subset.SetPromptForDwt(true); } catch { }
                }
            }
            if (subset == null) return;

            foreach (var drawing in node.Drawings)
            {
                if (drawing == null || string.IsNullOrWhiteSpace(drawing.DwgPath)) continue;
                if (!File.Exists(drawing.DwgPath)) continue;
                foreach (var layout in drawing.Layouts ?? Enumerable.Empty<ImportLayoutSheet>())
                {
                    if (layout == null || string.IsNullOrWhiteSpace(layout.LayoutName)) continue;
                    if (ImportOneSheet(subset, drawing.DwgPath, layout, stats))
                        stats.SheetsAdded++;
                }
            }

            foreach (var child in node.Children)
                MergeFolderNode(subset, child, newSheetFolder, stats);
        }

        private static bool ImportOneSheet(
            object host,
            string dwgPath,
            ImportLayoutSheet layout,
            FolderImportWriteResult stats)
        {
            if (host == null || layout == null) return false;
            var layoutName = layout.LayoutName;
            if (string.IsNullOrWhiteSpace(layoutName)) return false;
            if (SheetSetFolderImport.IsTemplateLayout(layoutName)) return false;
            if (string.IsNullOrWhiteSpace(layout.DrawingName)) return false;

            IAcSmPersist owner = host as IAcSmPersist;
            if (owner == null) return false;

            IAcSmAcDbLayoutReference layoutRef = null;
            IAcSmSheet sheet = null;
            try
            {
                layoutRef = (IAcSmAcDbLayoutReference)CreateComObject("AcSmAcDbLayoutReference");
                layoutRef.InitNew(owner);
                layoutRef.SetFileName(dwgPath);
                layoutRef.SetName(layoutName);

                if (host is IAcSmSubset subset)
                    sheet = subset.ImportSheet((AcSmAcDbLayoutReference)layoutRef);
                else if (host is IAcSmSheetSet sheetSet)
                    sheet = sheetSet.ImportSheet((AcSmAcDbLayoutReference)layoutRef);
                else
                    return false;

                if (sheet == null) return false;

                // SSM shows "Number - Title" → att01 - name_drawing1
                sheet.SetNumber(layoutName);
                sheet.SetTitle(SheetSetFolderImport.ResolveSheetTitle(
                    layoutName, layout.DrawingName, dwgPath));
                TrySetRevisionNumber(sheet, "");

                try
                {
                    if (host is IAcSmSubset sub)
                        sub.InsertComponent(sheet, null);
                    else if (host is IAcSmSheetSet ss)
                        ss.InsertComponent(sheet, null);
                }
                catch (COMException) { }

                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(sheet);
                ReleaseCom(layoutRef);
            }
        }

        /// <summary>
        /// Hands back the session's database for <paramref name="dstPath"/>. If AcSm already
        /// has it open — Sheet Set Manager's palette, or our own UI hold — that same object is
        /// reused instead of opening a rival handle on the file. <paramref name="weOpened"/>
        /// tells the caller whether it owns the matching Close.
        /// </summary>
        private static IAcSmDatabase AcquireDatabase(
            IAcSmSheetSetMgr manager, string dstPath, string tag, out bool weOpened)
        {
            weOpened = false;
            if (manager == null) return null;

            IAcSmDatabase database = null;
            try { database = manager.FindOpenDatabase(dstPath); } catch { database = null; }
            if (database != null)
            {
                SheetSetAutoLog.WriteStep(null, null, dstPath, $"FindOpenDatabase [{tag}]",
                    "reusing the database already open in this session");
                return database;
            }

            SheetSetAutoLog.WriteStep(null, null, dstPath, $"OpenDatabase [{tag}]", null);
            database = manager.OpenDatabase(dstPath, false);
            weOpened = true;
            return database;
        }

        /// <summary>
        /// Unlock, then close only a database this code opened. Closing one that Sheet Set
        /// Manager owns leaves its palette holding a dead object and breaks later AcSm calls.
        /// </summary>
        private static void ReleaseDatabase(
            IAcSmSheetSetMgr manager,
            IAcSmDatabase database,
            bool weOpened,
            bool locked,
            bool commit)
        {
            if (database != null && locked)
            {
                try { database.UnlockDb(database, commit); } catch { }
            }
            if (weOpened)
            {
                if (manager != null && database != null)
                {
                    try { manager.Close((AcSmDatabase)database); } catch { }
                }
                ReleaseCom(database);
            }
            ReleaseCom(manager);
        }

        private static object CreateComObject(string className)
        {
            Exception last = null;

            // 1) Prefer coclass from the referenced AcSm interop (correct CLSID for this build).
            try
            {
                var coclass = FindAcSmCoclass(className);
                if (coclass != null)
                {
                    var instance = Activator.CreateInstance(coclass);
                    if (instance != null) return instance;
                }
            }
            catch (Exception ex)
            {
                last = ex;
            }

            // 2) ProgID fallback — only versioned IDs exist (…​.24). Unversioned = Class not registered.
            var progIds = new List<string>();
            int major = GetAcadMajorVersion();
            var versions = new List<int>();
            if (major > 0) versions.Add(major);
            foreach (var v in new[] { 25, 24, 23, 22, 21 })
            {
                if (!versions.Contains(v)) versions.Add(v);
            }

            foreach (var v in versions)
                progIds.Add($"AcSmComponents.{className}.{v}");

            foreach (var progId in progIds)
            {
                try
                {
                    var type = Type.GetTypeFromProgID(progId, false);
                    if (type == null) continue;
                    var instance = Activator.CreateInstance(type);
                    if (instance != null) return instance;
                }
                catch (Exception ex)
                {
                    last = ex;
                }
            }

            throw new InvalidOperationException(
                $"AutoCAD Sheet Set component '{className}' is unavailable (Class not registered). "
                + "Run inside full AutoCAD; rebuild the add-in against matching AcSmComponents."
                + (last != null ? " " + last.Message : ""));
        }

        private static Type FindAcSmCoclass(string className)
        {
            string typeName = "Autodesk.AutoCAD.Interop." + className + "Class";
            foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
            {
                try
                {
                    var t = asm.GetType(typeName, throwOnError: false);
                    if (t != null && t.IsClass) return t;
                }
                catch { }
            }
            return null;
        }

        private static int GetAcadMajorVersion()
        {
            try
            {
                var raw = Convert.ToString(
                    AcadApp.GetSystemVariable("ACADVER"),
                    CultureInfo.InvariantCulture);
                if (string.IsNullOrWhiteSpace(raw)) return 0;
                var dot = raw.IndexOf('.');
                var first = dot >= 0 ? raw.Substring(0, dot) : raw;
                return int.TryParse(first, NumberStyles.Integer, CultureInfo.InvariantCulture, out var major)
                    ? major
                    : 0;
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// Drops exactly the one reference we took — never <c>FinalReleaseComObject</c>.
        /// AcSmSheetSetMgr is a per-session singleton that AutoCAD's own Sheet Set Manager
        /// shares, and databases can be co-owned by its palette; forcing an RCW's count to
        /// zero left AutoCAD holding released pointers, after which <c>CreateDatabase</c>
        /// failed for the rest of the session with a bogus "Class not registered".
        /// </summary>
        private static void ReleaseCom(object value)
        {
            if (value == null) return;
            try
            {
                if (Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
            }
            catch { }
        }
#endif

        private static bool PathsEqual(string a, string b)
        {
            if (string.IsNullOrWhiteSpace(a) || string.IsNullOrWhiteSpace(b))
                return string.Equals(a ?? "", b ?? "", StringComparison.OrdinalIgnoreCase);
            try
            {
                return string.Equals(
                    Path.GetFullPath(a),
                    Path.GetFullPath(b),
                    StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return string.Equals(a.Trim(), b.Trim(), StringComparison.OrdinalIgnoreCase);
            }
        }

        private static InvalidOperationException MissingInterop()
        {
            return new InvalidOperationException(
                "Sheet Set support was not compiled into this build (AcSmComponents.Interop.dll missing). "
                + "Rebuild with AutoCAD installed, or pass /p:AcSmInteropPath=... to the build.");
        }

        private static InvalidOperationException WrapCom(
            COMException ex, string operation = null, string dstPath = null, string dwgPath = null)
        {
            SheetSetAutoLog.WriteComFailure(dwgPath, dstPath, operation ?? "AcSm COM", ex);

            string msg = ex.Message ?? "";
            string hint;
            if (msg.IndexOf("Class not registered", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                hint = " (COM class not registered — AcSm component missing/mismatched for this AutoCAD. "
                    + "Not a simple DST file lock.)";
            }
            else if (unchecked((uint)ex.ErrorCode) == 0x80040211)
            {
                hint = " (HRESULT 0x80040211 — AcSm rejected the call. It is usually ImportSheet: "
                    + "the layout has to exist in the .dwg AS SAVED ON DISK, so save the drawing "
                    + "and try again. The log line 'ImportSheet REJECTED' names the layout and "
                    + "says whether it is in the saved file. See pl_sheetset_auto.log.)";
            }
            else
            {
                hint = "";
            }

            string op = string.IsNullOrWhiteSpace(operation) ? "" : " [" + operation + "]";
            return new InvalidOperationException(
                "AutoCAD Sheet Set API failed" + op + "." + hint + " " + msg, ex);
        }
    }
}
