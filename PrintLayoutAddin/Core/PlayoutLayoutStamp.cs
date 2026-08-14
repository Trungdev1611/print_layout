using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.DatabaseServices;

namespace PrintLayoutAddin.Core
{
    /// <summary>
    /// Marks the layouts PLAYOUT created, by writing an Xrecord into each Layout's
    /// extension dictionary. The mark lives inside the DWG, so it survives
    /// save/close/open and layout renames.
    ///
    /// This replaces the old "does the paper space look like a title block?" scan,
    /// which flagged every hand-made layout that had a border or an xref and made
    /// <see cref="LayoutDstSyncWatcher"/> prompt on deletions it should ignore.
    /// </summary>
    public static class PlayoutLayoutStamp
    {
        /// <summary>Key inside the Layout's extension dictionary.</summary>
        public const string DictKey = "PL_PLAYOUT_LAYOUT";

        private const string Marker = "PRINTLAYOUT";

        /// <summary>Mark <paramref name="layoutId"/> as PLAYOUT-created.</summary>
        public static void Stamp(Database db, ObjectId layoutId, string dstPath = null)
        {
            if (db == null || layoutId.IsNull) return;
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var layout = tr.GetObject(layoutId, OpenMode.ForWrite) as Layout;
                    if (layout != null)
                        Write(tr, layout, dstPath);
                    tr.Commit();
                }
            }
            catch { }
        }

        /// <summary>
        /// Record which DST the stamped layouts belong to, so the delete prompt can point
        /// the file picker straight at it. Called once PLAYOUT has written/updated the DST.
        /// </summary>
        public static void SetDstPathForStamped(Database db, string dstPath)
        {
            if (db == null || string.IsNullOrWhiteSpace(dstPath)) return;
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var dict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                    foreach (DBDictionaryEntry entry in dict)
                    {
                        Layout layout;
                        try { layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout; }
                        catch { continue; }
                        if (layout == null) continue;
                        if (!TryRead(tr, layout, out _)) continue;

                        try { layout.UpgradeOpen(); }
                        catch { continue; }
                        Write(tr, layout, dstPath);
                    }
                    tr.Commit();
                }
            }
            catch { }
        }

        /// <summary>Stamped layout name -&gt; DST path held in its stamp (may be empty).</summary>
        public static Dictionary<string, string> Snapshot(Database db)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (db == null) return map;
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var dict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                    foreach (DBDictionaryEntry entry in dict)
                    {
                        Layout layout;
                        try { layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout; }
                        catch { continue; }
                        if (layout == null) continue;
                        if (TryRead(tr, layout, out var dst))
                            map[layout.LayoutName] = dst;
                    }
                    tr.Commit();
                }
            }
            catch { }
            return map;
        }

        /// <summary>Read the stamp of a single live layout.</summary>
        public static bool TryReadByName(Database db, string layoutName, out string dstPath)
        {
            dstPath = "";
            if (db == null || string.IsNullOrWhiteSpace(layoutName)) return false;
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var dict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                    if (!dict.Contains(layoutName.Trim())) { tr.Commit(); return false; }

                    var layout = tr.GetObject(dict.GetAt(layoutName.Trim()), OpenMode.ForRead) as Layout;
                    bool stamped = layout != null && TryRead(tr, layout, out dstPath);
                    tr.Commit();
                    return stamped;
                }
            }
            catch { return false; }
        }

        /// <summary>
        /// Read the stamp off a Layout that is being erased. Opens erased objects — the
        /// extension dictionary is torn down together with the layout.
        /// </summary>
        public static bool TryReadFromErased(Layout layout, out string dstPath)
        {
            dstPath = "";
            if (layout == null) return false;
            try
            {
                var db = layout.Database;
                if (db == null) return false;
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    bool stamped = TryRead(tr, layout, out dstPath);
                    tr.Commit();
                    return stamped;
                }
            }
            catch { return false; }
        }

        private static void Write(Transaction tr, Layout layout, string dstPath)
        {
            try
            {
                if (layout.ExtensionDictionary.IsNull)
                    layout.CreateExtensionDictionary();

                var xdict = tr.GetObject(layout.ExtensionDictionary, OpenMode.ForWrite) as DBDictionary;
                if (xdict == null) return;

                var xrec = new Xrecord
                {
                    Data = new ResultBuffer(
                        new TypedValue((int)DxfCode.Text, Marker),
                        new TypedValue((int)DxfCode.Text, dstPath ?? "")),
                };
                xdict.SetAt(DictKey, xrec);
                tr.AddNewlyCreatedDBObject(xrec, true);
            }
            catch { }
        }

        private static bool TryRead(Transaction tr, Layout layout, out string dstPath)
        {
            dstPath = "";
            try
            {
                var extId = layout.ExtensionDictionary;
                if (extId.IsNull) return false;

                var xdict = tr.GetObject(extId, OpenMode.ForRead, true) as DBDictionary;
                if (xdict == null || !xdict.Contains(DictKey)) return false;

                var xrec = tr.GetObject(xdict.GetAt(DictKey), OpenMode.ForRead, true) as Xrecord;
                if (xrec?.Data == null) return false;

                var values = xrec.Data.AsArray();
                if (values.Length == 0) return false;
                if (!string.Equals(values[0].Value as string, Marker, StringComparison.OrdinalIgnoreCase))
                    return false;

                if (values.Length > 1)
                    dstPath = (values[1].Value as string) ?? "";
                return true;
            }
            catch { return false; }
        }
    }
}
