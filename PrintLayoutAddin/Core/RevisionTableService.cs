using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.DatabaseServices;

namespace PrintLayoutAddin.Core
{
    public class RevisionItem
    {
        public string RevNo { get; set; } = "";
        public string Description { get; set; } = "";
        public string Date { get; set; } = "";

        public bool IsEmpty =>
            string.IsNullOrWhiteSpace(RevNo)
            && string.IsNullOrWhiteSpace(Description)
            && string.IsNullOrWhiteSpace(Date);

        public string Summary
        {
            get
            {
                var no = (RevNo ?? "").Trim();
                var desc = (Description ?? "").Trim();
                if (string.IsNullOrEmpty(no) && string.IsNullOrEmpty(desc)) return "";
                if (string.IsNullOrEmpty(desc)) return no;
                if (string.IsNullOrEmpty(no)) return desc;
                return no + " - " + desc;
            }
        }
    }

    public class RevisionTableFindResult
    {
        public bool Found { get; set; }
        public string Message { get; set; }
        public List<RevisionItem> Items { get; set; } = new List<RevisionItem>();
        /// <summary>Physical data rows on the CAD Table (header excluded).</summary>
        public int CadDataRowCount { get; set; }
    }

    /// <summary>
    /// Reads/writes revision rows on a Paper Space Table located on
    /// <see cref="Config.RevTableLayer"/> (default TITLE_BLOCK_REV_TABLE).
    /// Column 0 = Rev No, 1 = Description, 2 = Date. Row 0 is the header.
    /// When the Table lives inside a shared title-block definition, Save clones
    /// that block per layout so revisions stay independent.
    /// </summary>
    public static class RevisionTableService
    {
        public const int ColRevNo = 0;
        public const int ColDescription = 1;
        public const int ColDate = 2;
        public const int MinColumns = 3;
        public const string UniqueBlockMarker = "__plrev_";

        /// <summary>Active slot count (from config; fallback <see cref="Config.DefaultRevTableDataRows"/>).</summary>
        public static int DataRowSlots
        {
            get
            {
                int n = Config.Instance?.RevTableDataRows ?? Config.DefaultRevTableDataRows;
                return n > 0 ? n : Config.DefaultRevTableDataRows;
            }
        }

        /// <summary>Pad/truncate to the fixed revision slot count (empties kept).</summary>
        public static List<RevisionItem> NormalizeToDataSlots(IEnumerable<RevisionItem> items)
        {
            int slots = DataRowSlots;
            var list = new List<RevisionItem>(slots);
            if (items != null)
            {
                foreach (var r in items)
                {
                    if (list.Count >= slots) break;
                    list.Add(r == null
                        ? new RevisionItem()
                        : new RevisionItem
                        {
                            RevNo = r.RevNo ?? "",
                            Description = r.Description ?? "",
                            Date = r.Date ?? "",
                        });
                }
            }
            while (list.Count < slots)
                list.Add(new RevisionItem());
            return list;
        }

        public static string FormatSummary(IList<RevisionItem> items, int maxLen = 60)
        {
            if (items == null || items.Count == 0) return "";
            RevisionItem last = null;
            for (int i = items.Count - 1; i >= 0; i--)
            {
                var it = items[i];
                if (it == null || it.IsEmpty) continue;
                if (string.IsNullOrWhiteSpace(it.Summary)) continue;
                last = it;
                break;
            }
            if (last == null) return "";
            var text = last.Summary.Replace('\r', ' ').Replace('\n', ' ').Trim();
            if (maxLen > 0 && text.Length > maxLen)
                return text.Substring(0, maxLen - 1) + "…";
            return text;
        }

        public static RevisionTableFindResult ReadRevisionsFromLayout(Database db, string layoutName)
        {
            var result = new RevisionTableFindResult();
            if (db == null)
            {
                result.Message = "Database is null.";
                return result;
            }
            if (string.IsNullOrWhiteSpace(layoutName))
            {
                result.Message = "Layout name is empty.";
                return result;
            }

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    if (!TryGetPaperBtr(db, layoutName, tr, OpenMode.ForRead, out var paperBtr, out var err))
                    {
                        result.Message = err;
                        tr.Commit();
                        return result;
                    }

                    var hit = LocateRevisionTable(paperBtr, tr, OpenMode.ForRead);
                    if (hit?.Table == null)
                    {
                        result.Message =
                            $"No Table on layer '{Config.Instance.RevTableLayer}' in layout '{layoutName}'.";
                        tr.Commit();
                        return result;
                    }

                    result.Found = true;
                    result.CadDataRowCount = Math.Max(0, hit.Table.Rows.Count - 1);
                    result.Items = NormalizeToDataSlots(ReadRows(hit.Table));
                    int filled = result.Items.Count(i => !i.IsEmpty);
                    result.Message =
                        $"Loaded {filled}/{DataRowSlots} slot(s); CAD table has {result.CadDataRowCount} data row(s).";
                    if (result.CadDataRowCount < DataRowSlots)
                    {
                        result.Message +=
                            $" Missing rows will be restored from '{Config.Instance.TemplateLayout}' on Save if that template still has them.";
                    }
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                result.Found = false;
                result.Message = ex.Message;
            }

            if (result.Items == null || result.Items.Count != DataRowSlots)
                result.Items = NormalizeToDataSlots(result.Items);
            return result;
        }

        public static bool WriteRevisionsToLayout(
            Database db,
            string layoutName,
            IList<RevisionItem> revisions,
            out string message)
        {
            message = "";
            if (db == null)
            {
                message = "Database is null.";
                return false;
            }
            if (string.IsNullOrWhiteSpace(layoutName))
            {
                message = "Layout name is empty.";
                return false;
            }

            // Always write a fixed number of slots (including blanks) — never grow/shrink Table.
            var rows = NormalizeToDataSlots(revisions);

            try
            {
                bool uniquified = false;
                bool truncated = false;
                int writeCount = 0;
                int dataCapacity = 0;
                int trimmedRows = 0;
                string restoreNote = null;
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    if (!TryGetPaperBtr(db, layoutName, tr, OpenMode.ForWrite, out var paperBtr, out var err))
                    {
                        message = err;
                        return false;
                    }

                    var hit = LocateRevisionTable(paperBtr, tr, OpenMode.ForWrite);
                    if (hit?.Table == null)
                    {
                        message =
                            $"No Table on layer '{Config.Instance.RevTableLayer}' in layout '{layoutName}'. "
                            + "Put a 3-column revision Table on that layer in Paper Space, "
                            + "or inside the title-block definition inserted on this layout "
                            + "(not inside an XREF). Close Block Editor first, then Reload.";
                        return false;
                    }

                    // Shared title-block → clone definition for this layout before writing.
                    if (hit.PaperSpaceBlockRef != null)
                    {
                        if (!EnsureLayoutUniqueTitleBlock(
                                db, tr, hit.PaperSpaceBlockRef, layoutName, out var uniqMsg, out uniquified))
                        {
                            message = uniqMsg;
                            return false;
                        }

                        // Re-locate after possible retarget (table ObjectId may have changed).
                        hit = LocateRevisionTable(paperBtr, tr, OpenMode.ForWrite);
                        if (hit?.Table == null)
                        {
                            message = "Revision Table lost after uniquifying title block.";
                            return false;
                        }

                        // Nested shared block (table not in outer def): uniquify that too.
                        if (hit.NestedOwnerBlockRef != null)
                        {
                            if (!EnsureNestedUniqueBlock(
                                    db, tr, hit.NestedOwnerBlockRef, layoutName, out uniqMsg, out var nestedUniq))
                            {
                                message = uniqMsg;
                                return false;
                            }
                            uniquified |= nestedUniq;
                            hit = LocateRevisionTable(paperBtr, tr, OpenMode.ForWrite);
                            if (hit?.Table == null)
                            {
                                message = "Revision Table lost after uniquifying nested block.";
                                return false;
                            }
                        }
                    }

                    var table = hit.Table;
                    if (table.Columns.Count < MinColumns)
                    {
                        message = $"Revision Table must have at least {MinColumns} columns.";
                        return false;
                    }

                    if (!table.IsWriteEnabled)
                        table.UpgradeOpen();

                    // Root cause of "can't add row 3": older DeleteRows removed empty
                    // template rows, so Rows.Count-1 collapsed to the filled count.
                    // Restore ONLY up to min(slots, template Layout1 rows) using that
                    // template's row heights — not unbounded InsertRows.
                    if (!TryRestoreMissingDataRowsFromTemplate(
                            db, tr, table, layoutName, DataRowSlots,
                            out dataCapacity, out restoreNote))
                    {
                        // Still write what fits; message explains shortfall.
                    }

                    // Temporary: drop physical rows beyond DefaultRevTableDataRows
                    // (e.g. table ballooned to 8) before writing slot content.
                    trimmedRows = TrimExcessDataRows(table, Config.DefaultRevTableDataRows);
                    dataCapacity = Math.Max(0, table.Rows.Count - 1);

                    writeCount = Math.Min(rows.Count, dataCapacity);
                    truncated = rows.Count > dataCapacity;

                    for (int i = 0; i < writeCount; i++)
                    {
                        int r = i + 1;
                        SetCellText(table, r, ColRevNo, rows[i].RevNo);
                        SetCellText(table, r, ColDescription, rows[i].Description);
                        SetCellText(table, r, ColDate, rows[i].Date);
                        ApplyDataCellAlignment(table, r);
                    }

                    // Clear any leftover data rows beyond written slots (keep geometry
                    // only when still within the allowed cap — excess already deleted).
                    for (int r = writeCount + 1; r < table.Rows.Count; r++)
                    {
                        SetCellText(table, r, ColRevNo, "");
                        SetCellText(table, r, ColDescription, "");
                        SetCellText(table, r, ColDate, "");
                        ApplyDataCellAlignment(table, r);
                    }

                    tr.Commit();
                }

                message = uniquified
                    ? $"Saved {writeCount}/{DataRowSlots} revision slot(s) to layout '{layoutName}' "
                      + "(title block cloned so other layouts are not affected)."
                    : $"Saved {writeCount}/{DataRowSlots} revision slot(s) to layout '{layoutName}'.";
                if (!string.IsNullOrWhiteSpace(restoreNote))
                    message += " " + restoreNote;
                if (trimmedRows > 0)
                    message +=
                        $" Removed {trimmedRows} excess data row(s) (cap {Config.DefaultRevTableDataRows}).";
                if (truncated)
                    message +=
                        $" Table only has {dataCapacity} data row(s) after restore attempt; "
                        + $"fix '{Config.Instance.TemplateLayout}' or lower revTableDataRows.";
                return true;
            }
            catch (System.Exception ex)
            {
                message = ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Batch-read last-revision summaries for many layouts (one transaction).
        /// </summary>
        public static Dictionary<string, string> ReadLastRevisionSummaries(
            Database db,
            IEnumerable<string> layoutNames)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (db == null || layoutNames == null) return map;

            var names = layoutNames
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (names.Count == 0) return map;

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    foreach (var name in names)
                    {
                        if (!TryGetPaperBtr(db, name, tr, OpenMode.ForRead, out var paperBtr, out _))
                        {
                            map[name] = "";
                            continue;
                        }
                        var hit = LocateRevisionTable(paperBtr, tr, OpenMode.ForRead);
                        if (hit?.Table == null)
                        {
                            map[name] = "";
                            continue;
                        }
                        map[name] = FormatSummary(ReadRows(hit.Table));
                    }
                    tr.Commit();
                }
            }
            catch
            {
                // leave missing keys empty
            }
            return map;
        }

        private sealed class LocateHit
        {
            public Table Table;
            /// <summary>Title-block insert on this layout's Paper Space (null if Table is direct).</summary>
            public BlockReference PaperSpaceBlockRef;
            /// <summary>If Table is inside a nested block of the title block, that nested insert.</summary>
            public BlockReference NestedOwnerBlockRef;
        }

        /// <summary>
        /// Delete trailing data rows until count equals <paramref name="maxDataRows"/>
        /// (header row 0 kept). Returns how many rows were removed.
        /// </summary>
        private static int TrimExcessDataRows(Table table, int maxDataRows)
        {
            if (table == null || maxDataRows < 0) return 0;
            if (!table.IsWriteEnabled)
                table.UpgradeOpen();

            int removed = 0;
            // Keep row 0 (header) + maxDataRows data rows → total Rows.Count == maxDataRows + 1
            while (table.Rows.Count - 1 > maxDataRows)
            {
                int last = table.Rows.Count - 1;
                if (last <= 0) break;
                table.DeleteRows(last, 1);
                removed++;
            }
            return removed;
        }

        /// <summary>
        /// If this layout's revision Table lost empty rows (e.g. old DeleteRows),
        /// re-create them using row heights from <see cref="Config.TemplateLayout"/>.
        /// Never grows past min(neededSlots, template data rows) — avoids NOTE overflow.
        /// </summary>
        private static bool TryRestoreMissingDataRowsFromTemplate(
            Database db,
            Transaction tr,
            Table table,
            string currentLayoutName,
            int neededSlots,
            out int capacity,
            out string note)
        {
            note = null;
            capacity = table == null ? 0 : Math.Max(0, table.Rows.Count - 1);
            if (table == null || neededSlots <= 0 || capacity >= neededSlots)
                return true;

            string templateName = Config.Instance.TemplateLayout;
            if (string.IsNullOrWhiteSpace(templateName))
            {
                note = $"CAD table has {capacity} data row(s); templateLayout is empty.";
                return false;
            }

            // Prefer a different layout's healthy table. If we're already on the
            // template layout, there is no other source to copy heights from.
            Table templateTable = null;
            if (!string.Equals(templateName, currentLayoutName, StringComparison.OrdinalIgnoreCase)
                && TryGetPaperBtr(db, templateName, tr, OpenMode.ForRead, out var templatePaper, out _))
            {
                templateTable = LocateRevisionTable(templatePaper, tr, OpenMode.ForRead)?.Table;
            }

            int templateDataRows = templateTable == null
                ? 0
                : Math.Max(0, templateTable.Rows.Count - 1);

            if (templateDataRows <= capacity)
            {
                note = templateTable == null
                    ? $"CAD table has {capacity} data row(s); no revision Table on template '{templateName}'."
                    : $"CAD table has {capacity} data row(s); template '{templateName}' only has {templateDataRows}.";
                return false;
            }

            int target = Math.Min(neededSlots, templateDataRows);
            if (!table.IsWriteEnabled)
                table.UpgradeOpen();

            int added = 0;
            while (capacity < target)
            {
                int templateRow = Math.Min(capacity + 1, templateTable.Rows.Count - 1);
                double height = templateTable.Rows[templateRow].Height;
                if (height <= 1e-9 && table.Rows.Count > 0)
                    height = table.Rows[table.Rows.Count - 1].Height;
                if (height <= 1e-9)
                    height = 1.0;

                // AutoCAD signature: InsertRows(rowIndex, height, count)
                table.InsertRows(table.Rows.Count, height, 1);
                int newRow = table.Rows.Count - 1;
                SetCellText(table, newRow, ColRevNo, "");
                SetCellText(table, newRow, ColDescription, "");
                SetCellText(table, newRow, ColDate, "");
                ApplyDataCellAlignment(table, newRow);
                capacity = Math.Max(0, table.Rows.Count - 1);
                added++;
            }

            note = added > 0
                ? $"Restored {added} empty row(s) from template '{templateName}' (now {capacity} data row(s))."
                : null;
            return capacity >= neededSlots;
        }

        /// <summary>Read physical data rows including blanks (slot order = CAD row order).</summary>
        private static List<RevisionItem> ReadRows(Table table)
        {
            var list = new List<RevisionItem>();
            if (table == null || table.Rows.Count <= 1) return list;
            int cols = table.Columns.Count;
            int max = Math.Min(table.Rows.Count - 1, DataRowSlots);
            for (int r = 1; r <= max; r++)
            {
                list.Add(new RevisionItem
                {
                    RevNo = GetCellText(table, r, ColRevNo),
                    Description = cols > ColDescription ? GetCellText(table, r, ColDescription) : "",
                    Date = cols > ColDate ? GetCellText(table, r, ColDate) : "",
                });
            }
            return list;
        }

        private static bool TryGetPaperBtr(
            Database db,
            string layoutName,
            Transaction tr,
            OpenMode mode,
            out BlockTableRecord paperBtr,
            out string error)
        {
            paperBtr = null;
            error = null;
            var lm = LayoutManager.Current;
            if (!lm.LayoutExists(layoutName))
            {
                error = $"Layout '{layoutName}' not found.";
                return false;
            }
            var layoutId = lm.GetLayoutId(layoutName);
            var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
            paperBtr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, mode);
            return true;
        }

        private static LocateHit LocateRevisionTable(
            BlockTableRecord paperBtr,
            Transaction tr,
            OpenMode mode)
        {
            string layer = Config.Instance.RevTableLayer;

            var direct = FindTableOnLayer(paperBtr, tr, mode, layer);
            if (direct != null)
                return new LocateHit { Table = direct };

            foreach (ObjectId id in paperBtr)
            {
                if (id.IsNull || id.IsErased) continue;
                var br = tr.GetObject(id, OpenMode.ForRead, false) as BlockReference;
                if (br == null) continue;

                ObjectId defId = GetEffectiveBlockDefId(br);
                if (defId.IsNull) continue;

                var def = (BlockTableRecord)tr.GetObject(defId, OpenMode.ForRead);
                if (def.IsFromExternalReference || def.IsFromOverlayReference)
                    continue;

                var nested = FindTableOnLayer(def, tr, mode, layer);
                if (nested != null)
                {
                    return new LocateHit
                    {
                        Table = nested,
                        PaperSpaceBlockRef = br,
                    };
                }

                // Dynamic title blocks: table often lives on the dynamic definition,
                // while inserts reference an anonymous record.
                if (br.IsDynamicBlock)
                {
                    try
                    {
                        var dynDef = (BlockTableRecord)tr.GetObject(
                            br.DynamicBlockTableRecord, OpenMode.ForRead);
                        var dynTable = FindTableOnLayer(dynDef, tr, mode, layer);
                        if (dynTable != null)
                        {
                            return new LocateHit
                            {
                                Table = dynTable,
                                PaperSpaceBlockRef = br,
                            };
                        }
                    }
                    catch { }
                }

                foreach (ObjectId nid in def)
                {
                    if (nid.IsNull || nid.IsErased) continue;
                    var nestedBr = tr.GetObject(nid, OpenMode.ForRead, false) as BlockReference;
                    if (nestedBr == null) continue;
                    ObjectId nestedDefId = GetEffectiveBlockDefId(nestedBr);
                    if (nestedDefId.IsNull) continue;
                    var nestedDef = (BlockTableRecord)tr.GetObject(nestedDefId, OpenMode.ForRead);
                    if (nestedDef.IsFromExternalReference || nestedDef.IsFromOverlayReference)
                        continue;
                    var deep = FindTableOnLayer(nestedDef, tr, mode, layer);
                    if (deep != null)
                    {
                        return new LocateHit
                        {
                            Table = deep,
                            PaperSpaceBlockRef = br,
                            NestedOwnerBlockRef = nestedBr,
                        };
                    }
                }
            }

            return null;
        }

        private static ObjectId GetEffectiveBlockDefId(BlockReference br)
        {
            if (br == null) return ObjectId.Null;
            // Visible geometry for dynamic inserts lives on the anonymous record;
            // for normal blocks BlockTableRecord is the definition.
            return br.BlockTableRecord;
        }

        private static Table FindTableOnLayer(
            BlockTableRecord container,
            Transaction tr,
            OpenMode mode,
            string layer)
        {
            foreach (ObjectId id in container)
            {
                if (id.IsNull || id.IsErased) continue;
                var obj = tr.GetObject(id, OpenMode.ForRead, false) as Entity;
                if (obj == null) continue;
                if (!(obj is Table)) continue;
                if (!string.Equals(obj.Layer, layer, StringComparison.OrdinalIgnoreCase))
                    continue;
                if (mode == OpenMode.ForWrite && !obj.IsWriteEnabled)
                    obj.UpgradeOpen();
                return (Table)obj;
            }
            return null;
        }

        /// <summary>
        /// If this layout's title-block insert still points at a shared definition,
        /// clone it to a per-layout name and retarget the insert.
        /// </summary>
        private static bool EnsureLayoutUniqueTitleBlock(
            Database db,
            Transaction tr,
            BlockReference paperBr,
            string layoutName,
            out string message,
            out bool uniquified)
        {
            message = null;
            uniquified = false;

            if (paperBr.IsDynamicBlock)
            {
                message =
                    "Revision Table is inside a dynamic title block shared by layouts. "
                    + "Use a static title block, or place the Table directly on Paper Space.";
                return false;
            }

            ObjectId defId = GetEffectiveBlockDefId(paperBr);
            var def = (BlockTableRecord)tr.GetObject(defId, OpenMode.ForRead);
            string desired = MakeUniqueBlockName(StripUniqueSuffix(def.Name), layoutName);

            if (string.Equals(def.Name, desired, StringComparison.OrdinalIgnoreCase))
                return true; // already unique for this layout

            // Also skip if definition has only this one reference AND name already layout-specific.
            // Otherwise always clone when name doesn't match desired — covers shared original.
            if (!TryCloneBlockAndRetarget(db, tr, paperBr, def, desired, out message))
                return false;

            uniquified = true;
            return true;
        }

        private static bool EnsureNestedUniqueBlock(
            Database db,
            Transaction tr,
            BlockReference nestedBr,
            string layoutName,
            out string message,
            out bool uniquified)
        {
            message = null;
            uniquified = false;

            ObjectId defId = GetEffectiveBlockDefId(nestedBr);
            var def = (BlockTableRecord)tr.GetObject(defId, OpenMode.ForRead);
            string desired = MakeUniqueBlockName(StripUniqueSuffix(def.Name), layoutName);

            if (string.Equals(def.Name, desired, StringComparison.OrdinalIgnoreCase))
                return true;

            // Only uniquify nested if shared by more than one insert.
            var refs = def.GetBlockReferenceIds(true, true);
            if (refs == null || refs.Count <= 1)
                return true;

            if (!TryCloneBlockAndRetarget(db, tr, nestedBr, def, desired, out message))
                return false;

            uniquified = true;
            return true;
        }

        private static bool TryCloneBlockAndRetarget(
            Database db,
            Transaction tr,
            BlockReference br,
            BlockTableRecord sourceDef,
            string newName,
            out string message)
        {
            message = null;
            try
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
                ObjectId newDefId;

                if (bt.Has(newName))
                {
                    newDefId = bt[newName];
                }
                else
                {
                    var newBtr = new BlockTableRecord
                    {
                        Name = newName,
                        Origin = sourceDef.Origin,
                    };
                    bt.Add(newBtr);
                    tr.AddNewlyCreatedDBObject(newBtr, true);

                    var ids = new ObjectIdCollection();
                    foreach (ObjectId oid in sourceDef)
                    {
                        if (!oid.IsNull && !oid.IsErased)
                            ids.Add(oid);
                    }

                    if (ids.Count > 0)
                    {
                        var map = new IdMapping();
                        db.DeepCloneObjects(ids, newBtr.ObjectId, map, false);
                    }

                    newDefId = newBtr.ObjectId;
                }

                var brW = (BlockReference)tr.GetObject(br.ObjectId, OpenMode.ForWrite);
                brW.BlockTableRecord = newDefId;
                return true;
            }
            catch (System.Exception ex)
            {
                message = "Failed to clone title block for per-layout revisions: " + ex.Message;
                return false;
            }
        }

        private static string StripUniqueSuffix(string blockName)
        {
            if (string.IsNullOrWhiteSpace(blockName)) return "TitleBlock";
            int idx = blockName.IndexOf(UniqueBlockMarker, StringComparison.OrdinalIgnoreCase);
            if (idx > 0) return blockName.Substring(0, idx);
            return blockName;
        }

        private static string MakeUniqueBlockName(string baseName, string layoutName)
        {
            string safeBase = SanitizeBlockNamePart(baseName);
            if (string.IsNullOrEmpty(safeBase)) safeBase = "TitleBlock";
            string safeLayout = SanitizeBlockNamePart(layoutName);
            if (string.IsNullOrEmpty(safeLayout)) safeLayout = "Layout";

            string name = safeBase + UniqueBlockMarker + safeLayout;
            // AutoCAD block name max length is 255.
            if (name.Length > 250)
                name = name.Substring(0, 250);
            return name;
        }

        private static string SanitizeBlockNamePart(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";
            var sb = new StringBuilder(text.Length);
            foreach (char c in text.Trim())
            {
                if (char.IsLetterOrDigit(c) || c == '_' || c == '-')
                    sb.Append(c);
                else
                    sb.Append('_');
            }
            return sb.ToString();
        }

        private static void ApplyDataCellAlignment(Table table, int row)
        {
            if (table == null || row < 0 || row >= table.Rows.Count) return;
            int cols = Math.Min(table.Columns.Count, MinColumns);
            for (int c = 0; c < cols; c++)
            {
                try
                {
                    table.Cells[row, c].Alignment = c == ColDescription
                        ? CellAlignment.MiddleLeft
                        : CellAlignment.MiddleCenter;
                }
                catch { }
            }
        }

        private static string GetCellText(Table table, int row, int col)
        {
            try
            {
                if (row < 0 || row >= table.Rows.Count) return "";
                if (col < 0 || col >= table.Columns.Count) return "";
                return table.Cells[row, col].TextString ?? "";
            }
            catch
            {
                return "";
            }
        }

        private static void SetCellText(Table table, int row, int col, string value)
        {
            table.Cells[row, col].TextString = value ?? "";
        }
    }
}
