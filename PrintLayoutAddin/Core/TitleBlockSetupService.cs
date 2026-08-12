using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace PrintLayoutAddin.Core
{
    /// <summary>
    /// Helpers for <c>PLFRAME_SETUP</c>: Sheet Set fields on title-block attributes
    /// and a revision Table on <see cref="Config.RevTableLayer"/>.
    /// </summary>
    public static class TitleBlockSetupService
    {
        public enum SheetSetFieldKind
        {
            SheetNumber,
            SheetTitle,
            Revision,
        }

        /// <summary>Contextual Sheet Set field keys (Field dialog names).</summary>
        public static string FieldKey(SheetSetFieldKind kind)
        {
            switch (kind)
            {
                case SheetSetFieldKind.SheetNumber:
                    return "CurrentSheetNumber";
                case SheetSetFieldKind.SheetTitle:
                    return "CurrentSheetTitle";
                case SheetSetFieldKind.Revision:
                    return "CurrentSheetRevisionNumber";
                default:
                    throw new ArgumentOutOfRangeException(nameof(kind));
            }
        }

        public static string FieldPromptLabel(SheetSetFieldKind kind)
        {
            switch (kind)
            {
                case SheetSetFieldKind.SheetNumber:
                    return "sheet number (CurrentSheetNumber)";
                case SheetSetFieldKind.SheetTitle:
                    return "sheet title (CurrentSheetTitle)";
                case SheetSetFieldKind.Revision:
                    return "revision (CurrentSheetRevisionNumber)";
                default:
                    return "field";
            }
        }

        public static string FieldPlacementPrompt(SheetSetFieldKind kind)
        {
            switch (kind)
            {
                case SheetSetFieldKind.SheetNumber:
                    return "Pick insertion point for sheet number field (DRAWING_NO):";
                case SheetSetFieldKind.SheetTitle:
                    return "Pick insertion point for sheet title field (DRAWING_NAME):";
                case SheetSetFieldKind.Revision:
                    return "Pick insertion point for revision field (REVISION):";
                default:
                    return "Pick insertion point for field:";
            }
        }

        private sealed class AttributeSpec
        {
            public string Tag;
            public string Prompt;
            /// <summary>Visible default text in BEDIT for alignment (e.g. DRAWING_NAME).</summary>
            public string Placeholder;
        }

        private static AttributeSpec DefaultAttributeSpec(SheetSetFieldKind kind)
        {
            switch (kind)
            {
                case SheetSetFieldKind.SheetNumber:
                    return new AttributeSpec
                    {
                        Tag = "DRAWING_NO",
                        Prompt = "Drawing No",
                        Placeholder = "DRAWING_NO",
                    };
                case SheetSetFieldKind.SheetTitle:
                    return new AttributeSpec
                    {
                        Tag = "DRAWING_NAME",
                        Prompt = "Drawing Name",
                        Placeholder = "DRAWING_NAME",
                    };
                case SheetSetFieldKind.Revision:
                    return new AttributeSpec
                    {
                        Tag = "REVISION",
                        Prompt = "Revision",
                        Placeholder = "REVISION",
                    };
                default:
                    throw new ArgumentOutOfRangeException(nameof(kind));
            }
        }

        /// <summary>Raw field expression for contextual Sheet Set fields.</summary>
        public static string BuildFieldCode(string fieldKey)
        {
            // Field dialog names (CurrentSheet*) map to AcSm Sheet.* — AcVar CurrentSheet*
            // is invalid and evaluates to ####.
            switch (fieldKey)
            {
                case "CurrentSheetNumber":
                    return "%<\\AcSm Sheet.Number \\f \"%tc1\">%";
                case "CurrentSheetTitle":
                    return "%<\\AcSm Sheet.Title \\f \"%tc1\">%";
                case "CurrentSheetRevisionNumber":
                    return "%<\\AcSm Sheet.RevisionNumber \\f \"%tc1\">%";
                default:
                    return "%<\\AcSm Sheet.Number \\f \"%tc1\">%";
            }
        }

        public static string BuildFieldCode(SheetSetFieldKind kind) =>
            BuildFieldCode(FieldKey(kind));

        /// <summary>Default attribute / placeholder text height (drawing units).</summary>
        public const double DefaultTextHeight = 3.0;

        /// <summary>Default total revision-table width (drawing units).</summary>
        public const double DefaultTableWidth = 75.0;

        /// <summary>Default revision-table header text height.</summary>
        public const double DefaultTableHeaderTextHeight = 3.0;

        /// <summary>Default revision-table data cell text height.</summary>
        public const double DefaultTableDataTextHeight = 3.0;

        /// <summary>Default height for the secondary title under the viewport.</summary>
        public const double DefaultCenterTitleHeight = 5.0;

        /// <summary>List text style names in the drawing (for UI dropdowns).</summary>
        public static System.Collections.Generic.List<string> ListTextStyleNames(Database db)
        {
            var names = new System.Collections.Generic.List<string>();
            if (db == null) return names;
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var tst = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);
                    foreach (ObjectId id in tst)
                    {
                        if (id.IsNull || id.IsErased) continue;
                        var rec = (TextStyleTableRecord)tr.GetObject(id, OpenMode.ForRead);
                        if (rec == null || string.IsNullOrWhiteSpace(rec.Name)) continue;
                        // Skip anonymous / shape styles with empty usable names.
                        if (rec.Name.StartsWith("*", StringComparison.Ordinal)) continue;
                        names.Add(rec.Name);
                    }
                    tr.Commit();
                }
                names.Sort(StringComparer.OrdinalIgnoreCase);
            }
            catch { }

            if (names.Count == 0)
                names.Add("Standard");
            return names;
        }

        public static ObjectId ResolveTextStyleId(Database db, Transaction tr, string styleName)
        {
            if (db == null || tr == null) return ObjectId.Null;
            try
            {
                var tst = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);
                if (!string.IsNullOrWhiteSpace(styleName) && tst.Has(styleName))
                    return tst[styleName];
                if (tst.Has("Standard"))
                    return tst["Standard"];
                foreach (ObjectId id in tst)
                {
                    if (!id.IsNull && !id.IsErased)
                        return id;
                }
            }
            catch { }
            return ObjectId.Null;
        }

        /// <summary>
        /// Allowed everywhere a drawing is open: Model Space (xref title-block source),
        /// Paper Space, or Block Editor.
        /// </summary>
        public static bool IsAllowedSpace(Database db) => db != null;

        /// <summary>True Model Space and not inside Block Editor.</summary>
        public static bool IsRealModelSpace(Database db) =>
            db != null && db.TileMode && !IsBlockEditorActive();

        public static bool IsBlockEditorActive()
        {
            try
            {
                object v = AcadApp.GetSystemVariable("BLOCKEDITOR");
                if (v == null) return false;
                return Convert.ToInt32(v) != 0;
            }
            catch
            {
                return false;
            }
        }

        public static string DescribeCurrentSpace(Database db)
        {
            if (db == null) return "Unknown";
            if (IsBlockEditorActive())
            {
                try
                {
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
                        string name = btr?.Name ?? "?";
                        tr.Commit();
                        return $"Block Editor ({name})";
                    }
                }
                catch
                {
                    return "Block Editor";
                }
            }

            if (db.TileMode)
                return "Model Space (OK for xref title-block source)";

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
                    bool layout = IsLayoutBlock(db, btr, tr);
                    tr.Commit();
                    if (layout)
                        return "Paper Space";
                    return $"Current space ({btr.Name})";
                }
            }
            catch
            {
                return "Paper Space";
            }
        }

        public static bool InsertFieldAtPoint(
            Database db,
            Point3d insertPoint,
            SheetSetFieldKind kind,
            out string message,
            double textHeight = 0,
            string textStyleName = null,
            bool centerAligned = false)
        {
            message = "";
            if (db == null)
            {
                message = "Database is null.";
                return false;
            }

            var spec = DefaultAttributeSpec(kind);
            double height = textHeight > 1e-9 ? textHeight : DefaultTextHeight;

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    // Only BEDIT creates AttributeDefinitions. Model + Paper use DBText
                    // (xref title-block sources are typically drawn in Model Space).
                    bool inBlockEditor = IsBlockEditorActive();
                    ObjectId styleId = ResolveTextStyleId(db, tr, textStyleName);

                    if (inBlockEditor)
                    {
                        string tag = MakeUniqueTag(space, tr, spec.Tag);
                        var attDef = new AttributeDefinition
                        {
                            Position = insertPoint,
                            Height = height,
                            Tag = tag,
                            Prompt = spec.Prompt,
                            TextString = spec.Placeholder,
                            LockPositionInBlock = true,
                        };
                        if (centerAligned)
                        {
                            attDef.HorizontalMode = TextHorizontalMode.TextCenter;
                            attDef.VerticalMode = TextVerticalMode.TextVerticalMid;
                            attDef.AlignmentPoint = insertPoint;
                            try { attDef.AdjustAlignment(db); } catch { }
                        }
                        if (!styleId.IsNull)
                            attDef.TextStyleId = styleId;
                        space.AppendEntity(attDef);
                        tr.AddNewlyCreatedDBObject(attDef, true);
                        tr.Commit();
                        message =
                            $"Created attribute '{tag}' showing '{spec.Placeholder}' at picked point "
                            + $"(height {height:F2}). Use button 5 to link Sheet Set fields when alignment is done.";
                        return true;
                    }

                    var text = new DBText();
                    text.SetDatabaseDefaults(db);
                    text.Height = height;
                    text.TextString = spec.Placeholder;
                    if (!styleId.IsNull)
                        text.TextStyleId = styleId;

                    if (centerAligned)
                    {
                        // AlignmentPoint is ignored unless Position is set and AdjustAlignment runs.
                        text.Position = insertPoint;
                        text.HorizontalMode = TextHorizontalMode.TextCenter;
                        text.VerticalMode = TextVerticalMode.TextVerticalMid;
                        text.AlignmentPoint = insertPoint;
                        try { text.AdjustAlignment(db); } catch { }
                    }
                    else
                    {
                        text.Position = insertPoint;
                    }

                    space.AppendEntity(text);
                    tr.AddNewlyCreatedDBObject(text, true);
                    tr.Commit();
                    string where = db.TileMode ? "Model Space" : "this layout";
                    message =
                        $"Created text '{spec.Placeholder}' in {where} @ ({insertPoint.X:F2},{insertPoint.Y:F2})"
                        + $" (height {height:F2}"
                        + (centerAligned ? ", middle-center" : "")
                        + "). "
                        + (db.TileMode
                            ? "Save this DWG and xref/reload it on host layouts."
                            : "For a shared title block, BEDIT the block or edit the xref source DWG.");
                    return true;
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Convert DRAWING_* placeholders to Sheet Set field codes.
        /// BEDIT → AttributeDefinitions in the block.
        /// Model / Paper Space → DBText / MText placeholders from buttons 1–3.
        /// </summary>
        public static bool ActivateSheetSetFieldsOnPlaceholders(Database db, out string message)
        {
            message = "";
            if (db == null)
            {
                message = "Database is null.";
                return false;
            }

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    bool inBlockEditor = IsBlockEditorActive();

                    int activated = 0;
                    if (inBlockEditor)
                    {
                        foreach (ObjectId id in space)
                        {
                            if (id.IsNull || id.IsErased) continue;
                            if (!(tr.GetObject(id, OpenMode.ForWrite, false) is AttributeDefinition ad))
                                continue;
                            if (!TryResolveFieldKind(ad.Tag, out var fieldKind)
                                && !TryResolveFieldKind(ad.TextString, out fieldKind))
                                continue;

                            ad.TextString = BuildFieldCode(fieldKind);
                            activated++;
                        }

                        tr.Commit();
                        message = activated > 0
                            ? $"Linked {activated} attribute(s) in block to Sheet Set fields. "
                              + "BCLOSE Save, then REGEN / UPDATEFIELD if values show ####."
                            : "No DRAWING_NO / DRAWING_NAME / REVISION attributes in this block. "
                              + "Place with buttons 1–3 first.";
                        return activated > 0;
                    }

                    // Model or Paper Space: DBText / MText placeholders.
                    foreach (ObjectId id in space)
                    {
                        if (id.IsNull || id.IsErased) continue;
                        var obj = tr.GetObject(id, OpenMode.ForWrite, false);

                        if (obj is DBText dbText)
                        {
                            if (!TryResolveFieldKind(dbText.TextString, out var fieldKind))
                                continue;
                            dbText.TextString = BuildFieldCode(fieldKind);
                            activated++;
                            continue;
                        }

                        if (obj is MText mText)
                        {
                            string plain = (mText.Contents ?? "").Trim();
                            if (plain.StartsWith("{", StringComparison.Ordinal)
                                && plain.Contains("\\"))
                                plain = mText.Text ?? plain;
                            if (!TryResolveFieldKind(plain, out var fieldKind)
                                && !TryResolveFieldKind(mText.Contents, out fieldKind))
                                continue;
                            mText.Contents = BuildFieldCode(fieldKind);
                            activated++;
                        }
                    }

                    tr.Commit();
                    string where = db.TileMode ? "Model Space" : "this layout";
                    message = activated > 0
                        ? $"Linked {activated} text field(s) in {where} to Sheet Set (AcSm). "
                          + "Run UPDATEFIELD then REGEN. "
                          + (db.TileMode
                              ? "Save & reload xref on host if this is a title-block source."
                              : "Still ####? Open that layout's sheet in SSM.")
                        : $"No DRAWING_NO / DRAWING_NAME / REVISION text in {where}. "
                          + "Place with buttons 1–3 first.";
                    return activated > 0;
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return false;
            }
        }

        /// <summary>Tooltip / hint text for button 5 based on current space.</summary>
        public static string ActivateButtonHint(Database db)
        {
            if (db == null) return "Activate Sheet Set fields on placeholders.";
            if (IsBlockEditorActive())
                return "BEDIT: convert DRAWING_* attributes in this block to Sheet Set fields.";
            if (db.TileMode)
                return "Model Space: convert DRAWING_* text to Sheet Set fields (xref title-block source).";
            return "This layout: convert DRAWING_* placeholders to Sheet Set fields.";
        }

        /// <summary>
        /// Place DRAWING_* placeholders at the center of each label's line-bounded cell (method B).
        /// Does not activate Sheet Set fields — use button 5 / Auto for that.
        /// </summary>
        public static bool PlaceFieldsNearScannedLabels(
            Database db,
            TitleBlockLabelScanner.ScanResult scan,
            out string message,
            double titleHeight = 0,
            string titleStyle = null,
            double numberHeight = 0,
            string numberStyle = null,
            double revisionHeight = 0,
            string revisionStyle = null)
        {
            message = "";
            if (db == null || scan == null)
            {
                message = "Nothing to place.";
                return false;
            }

            var lines = new List<string>();
            int ok = 0, fail = 0;
            var search = scan.FullStrip;

            void TryOne(TitleBlockLabelScanner.Hit hit, double ht, string style)
            {
                if (hit == null)
                {
                    lines.Add("skip (label not found)");
                    return;
                }

                if (!TitleBlockLabelScanner.TryResolveCellCenter(
                        db, hit, search, out var center, out var cell, out var cellDetail))
                {
                    fail++;
                    lines.Add($"FAIL {hit.Kind}: no cell — {cellDetail}");
                    return;
                }

                if (InsertFieldAtPoint(
                        db, center, hit.Kind, out var msg, ht, style, centerAligned: true))
                {
                    ok++;
                    lines.Add($"OK {hit.Kind}: {cellDetail} → {msg}");
                }
                else
                {
                    fail++;
                    lines.Add($"FAIL {hit.Kind}: {msg}");
                }
            }

            TryOne(scan.Title, titleHeight, titleStyle);
            TryOne(scan.Number, numberHeight, numberStyle);
            TryOne(scan.Revision, revisionHeight, revisionStyle);

            message = $"Placed {ok} field(s)"
                + (fail > 0 ? $", {fail} failed" : "")
                + ".\n" + string.Join("\n", lines);
            return ok > 0;
        }

        /// <summary>
        /// Secondary drawing title at mid-bottom of the PLAYOUT viewport (CurrentSheetTitle).
        /// </summary>
        public static bool PlaceCenterTitle(
            Database db,
            ViewportCornerGeometry.Bounds viewport,
            out string message,
            double textHeight = 0,
            string textStyleName = null)
        {
            double ht = textHeight > 1e-9 ? textHeight : DefaultCenterTitleHeight;
            var pt = ViewportCornerGeometry.CenterTitlePoint(viewport, ht);
            return InsertFieldAtPoint(
                db, pt, SheetSetFieldKind.SheetTitle, out message, ht, textStyleName,
                centerAligned: true);
        }

        /// <summary>
        /// Auto Frame Setup: scan+place title-block fields, rev table at viewport top-right,
        /// secondary center title, then activate Sheet Set field codes.
        /// </summary>
        public static bool RunAutoFrameSetup(
            Database db,
            ViewportCornerGeometry.Bounds viewport,
            out string message,
            double titleHeight = 0,
            string titleStyle = null,
            double centerTitleHeight = 0,
            string centerTitleStyle = null,
            double tableWidth = 0,
            double headerTextHeight = 0,
            double dataTextHeight = 0,
            string headerStyleName = null,
            string dataStyleName = null,
            double numberHeight = 0,
            string numberStyle = null,
            double revisionHeight = 0,
            string revisionStyle = null)
        {
            message = "";
            if (db == null)
            {
                message = "Database is null.";
                return false;
            }

            var lines = new List<string>();
            lines.Add(ViewportCornerGeometry.Describe(viewport));
            int okSteps = 0;
            int softFails = 0;

            double numHt = numberHeight > 1e-9 ? numberHeight : titleHeight;
            string numStyle = !string.IsNullOrWhiteSpace(numberStyle) ? numberStyle : titleStyle;
            double revHt = revisionHeight > 1e-9 ? revisionHeight : numHt;
            string revStyle = !string.IsNullOrWhiteSpace(revisionStyle) ? revisionStyle : numStyle;

            var scan = TitleBlockLabelScanner.Scan(db, viewport);
            lines.Add(scan.Summarize());

            bool fieldsOk = PlaceFieldsNearScannedLabels(
                db, scan, out var fieldsMsg,
                titleHeight, titleStyle,
                numHt, numStyle,
                revHt, revStyle);
            lines.Add("--- Fields ---");
            lines.Add(fieldsMsg ?? (fieldsOk ? "OK" : "Failed"));
            if (fieldsOk) okSteps++; else softFails++;

            var revAt = ViewportCornerGeometry.RevTableInsertPoint(viewport);
            bool tableOk = InsertRevisionTable(
                db, revAt, out var tableMsg,
                tableWidth, headerTextHeight, dataTextHeight,
                headerStyleName, dataStyleName,
                skipIfExists: true);
            lines.Add("--- Rev table @ (" + revAt.X.ToString("F1") + "," + revAt.Y.ToString("F1") + ") ---");
            lines.Add(tableMsg ?? (tableOk ? "OK" : "Failed"));
            if (tableOk) okSteps++; else softFails++;

            double cHt = centerTitleHeight > 1e-9
                ? centerTitleHeight
                : (titleHeight > 1e-9 ? Math.Max(titleHeight * 2.0, DefaultCenterTitleHeight) : DefaultCenterTitleHeight);
            bool centerOk = PlaceCenterTitle(
                db, viewport, out var centerMsg, cHt, centerTitleStyle ?? titleStyle);
            lines.Add("--- Center title ---");
            lines.Add(centerMsg ?? (centerOk ? "OK" : "Failed"));
            if (centerOk) okSteps++; else softFails++;

            bool actOk = ActivateSheetSetFieldsOnPlaceholders(db, out var actMsg);
            lines.Add("--- Activate ---");
            lines.Add(actMsg ?? (actOk ? "OK" : "Failed"));
            if (actOk) okSteps++; else softFails++;

            message = $"Auto Frame Setup: {okSteps} step(s) ok"
                + (softFails > 0 ? $", {softFails} issue(s)" : "")
                + ".\n" + string.Join("\n", lines);
            return okSteps > 0;
        }

        public static bool ApplySheetSetField(
            Database db,
            ObjectId attributeId,
            SheetSetFieldKind kind,
            out string message)
        {
            message = "";
            if (db == null)
            {
                message = "Database is null.";
                return false;
            }
            if (attributeId.IsNull || attributeId.IsErased)
            {
                message = "Invalid attribute.";
                return false;
            }

            string fieldKey = FieldKey(kind);
            string code = BuildFieldCode(kind);

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var obj = tr.GetObject(attributeId, OpenMode.ForWrite, false);
                    if (obj is AttributeDefinition attDef)
                    {
                        ApplyFieldToTextObject(attDef, code);
                        tr.Commit();
                        message =
                            $"Applied {fieldKey} to attribute definition '{attDef.Tag}' "
                            + "(saved in block definition).";
                        return true;
                    }

                    if (obj is AttributeReference attRef)
                    {
                        ApplyFieldToTextObject(attRef, code);
                        tr.Commit();
                        message =
                            $"Applied {fieldKey} to attribute '{attRef.Tag}' on this insert. "
                            + "Use BEDIT on the title-block definition to affect all layouts.";
                        return true;
                    }

                    message = "Selected object is not an attribute.";
                    return false;
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return false;
            }
        }

        public static bool EnsureRevTableLayer(Database db, out string message)
        {
            message = "";
            if (db == null)
            {
                message = "Database is null.";
                return false;
            }

            string layerName = Config.Instance?.RevTableLayer ?? Config.DefaultRevTableLayer;
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (lt.Has(layerName))
                    {
                        var existing = (LayerTableRecord)tr.GetObject(lt[layerName], OpenMode.ForWrite);
                        bool touched = false;
                        if (!existing.IsPlottable)
                        {
                            existing.IsPlottable = true;
                            touched = true;
                        }
                        if (existing.IsLocked)
                        {
                            existing.IsLocked = false;
                            touched = true;
                        }
                        if (existing.IsFrozen)
                        {
                            existing.IsFrozen = false;
                            touched = true;
                        }
                        tr.Commit();
                        message = touched
                            ? $"Layer '{layerName}' already exists (unlocked / unfrozen / plottable)."
                            : $"Layer '{layerName}' already exists.";
                        return true;
                    }

                    lt.UpgradeOpen();
                    var rec = new LayerTableRecord
                    {
                        Name = layerName,
                        Color = Color.FromColorIndex(ColorMethod.ByAci, 7),
                        IsPlottable = true,
                    };
                    lt.Add(rec);
                    tr.AddNewlyCreatedDBObject(rec, true);
                    tr.Commit();
                    message = $"Created layer '{layerName}' (color 7, plottable).";
                    return true;
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return false;
            }
        }

        public static bool InsertRevisionTable(
            Database db,
            Point3d insertPoint,
            out string message,
            double tableWidth = 0,
            double headerTextHeight = 0,
            double dataTextHeight = 0,
            string headerStyleName = null,
            string dataStyleName = null,
            bool skipIfExists = false)
        {
            message = "";
            if (db == null)
            {
                message = "Database is null.";
                return false;
            }

            if (!EnsureRevTableLayer(db, out var layerMsg))
            {
                message = layerMsg;
                return false;
            }

            string layerName = Config.Instance?.RevTableLayer ?? Config.DefaultRevTableLayer;
            int dataRows = RevisionTableService.DataRowSlots;
            double width = tableWidth > 1e-9 ? tableWidth : DefaultTableWidth;
            double headerHt = headerTextHeight > 1e-9
                ? headerTextHeight
                : DefaultTableHeaderTextHeight;
            double dataHt = dataTextHeight > 1e-9
                ? dataTextHeight
                : DefaultTableDataTextHeight;

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    if (TableExistsOnLayer(space, tr, layerName))
                    {
                        if (skipIfExists)
                        {
                            message =
                                $"Skipped — Table on layer '{layerName}' already exists.";
                            return true;
                        }
                        message =
                            $"A Table on layer '{layerName}' already exists in this space. "
                            + "Delete it first if you want a new one.";
                        return false;
                    }

                    ObjectId headerStyleId = ResolveTextStyleId(db, tr, headerStyleName);
                    ObjectId dataStyleId = ResolveTextStyleId(db, tr, dataStyleName);

                    var table = BuildRevisionTable(
                        insertPoint, layerName, dataRows, width, headerHt, dataHt,
                        headerStyleId, dataStyleId);
                    space.AppendEntity(table);
                    tr.AddNewlyCreatedDBObject(table, true);
                    tr.Commit();
                    message =
                        $"Inserted revision Table ({dataRows} data rows, width {width:F1}, "
                        + $"header ht {headerHt:F2}, data ht {dataHt:F2}) on layer '{layerName}'. "
                        + layerMsg;
                    return true;
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return false;
            }
        }

        private static Table BuildRevisionTable(
            Point3d insertPoint,
            string layerName,
            int dataRows,
            double totalWidth,
            double headerTextHeight,
            double dataTextHeight,
            ObjectId headerStyleId,
            ObjectId dataStyleId)
        {
            int rows = 1 + Math.Max(1, dataRows);
            const int cols = RevisionTableService.MinColumns;

            // Proportional widths: Rev 15% / Description 60% / Date 25%.
            double w = totalWidth > 1e-9 ? totalWidth : DefaultTableWidth;
            double colRev = w * 0.15;
            double colDesc = w * 0.60;
            double colDate = w * 0.25;
            double headerHt = headerTextHeight > 1e-9
                ? headerTextHeight
                : DefaultTableHeaderTextHeight;
            double dataHt = dataTextHeight > 1e-9
                ? dataTextHeight
                : DefaultTableDataTextHeight;
            double headerRowH = Math.Max(headerHt * 1.6, 3.0);
            double dataRowH = Math.Max(dataHt * 1.6, 3.0);

            var table = new Table
            {
                Position = insertPoint,
                Layer = layerName,
            };
            table.SetSize(rows, cols);
            table.Columns[RevisionTableService.ColRevNo].Width = colRev;
            table.Columns[RevisionTableService.ColDescription].Width = colDesc;
            table.Columns[RevisionTableService.ColDate].Width = colDate;

            table.Rows[0].Height = headerRowH;
            for (int r = 1; r < rows; r++)
                table.Rows[r].Height = dataRowH;

            SetHeaderCell(table, 0, RevisionTableService.ColRevNo, "Rev", headerHt, headerStyleId);
            SetHeaderCell(table, 0, RevisionTableService.ColDescription, "Description", headerHt, headerStyleId);
            SetHeaderCell(table, 0, RevisionTableService.ColDate, "Date", headerHt, headerStyleId);

            for (int r = 1; r < rows; r++)
            {
                SetDataCell(table, r, RevisionTableService.ColRevNo, "", dataHt, dataStyleId);
                SetDataCell(table, r, RevisionTableService.ColDescription, "", dataHt, dataStyleId);
                SetDataCell(table, r, RevisionTableService.ColDate, "", dataHt, dataStyleId);
            }

            return table;
        }

        private static void SetHeaderCell(
            Table table, int row, int col, string text, double textHeight, ObjectId textStyleId)
        {
            var cell = table.Cells[row, col];
            cell.TextString = text ?? "";
            try
            {
                cell.Alignment = CellAlignment.MiddleCenter;
                if (textHeight > 1e-9)
                    cell.TextHeight = textHeight;
                if (!textStyleId.IsNull)
                    cell.TextStyleId = textStyleId;
            }
            catch { }
        }

        private static void SetDataCell(
            Table table, int row, int col, string text, double textHeight, ObjectId textStyleId)
        {
            var cell = table.Cells[row, col];
            cell.TextString = text ?? "";
            try
            {
                cell.Alignment = col == RevisionTableService.ColDescription
                    ? CellAlignment.MiddleLeft
                    : CellAlignment.MiddleCenter;
                if (textHeight > 1e-9)
                    cell.TextHeight = textHeight;
                if (!textStyleId.IsNull)
                    cell.TextStyleId = textStyleId;
            }
            catch { }
        }

        private static bool TableExistsOnLayer(BlockTableRecord space, Transaction tr, string layerName)
        {
            foreach (ObjectId id in space)
            {
                if (id.IsNull || id.IsErased) continue;
                if (!(tr.GetObject(id, OpenMode.ForRead, false) is Table tbl)) continue;
                if (string.Equals(tbl.Layer, layerName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }

        private static void ApplyFieldToTextObject(DBText dbText, string fieldCode)
        {
            if (dbText == null) throw new ArgumentNullException(nameof(dbText));
            if (!dbText.IsWriteEnabled)
                dbText.UpgradeOpen();
            dbText.TextString = fieldCode ?? "";
        }

        private static bool TryResolveFieldKind(string key, out SheetSetFieldKind kind)
        {
            kind = default;
            if (string.IsNullOrWhiteSpace(key)) return false;
            string t = key.Trim();

            if (TryMapTagToFieldKind(t, out kind))
                return true;

            // Rewrite previous AcVar / AcSm field codes (including #### sources).
            if (t.IndexOf("CurrentSheetNumber", StringComparison.OrdinalIgnoreCase) >= 0
                || t.IndexOf("Sheet.Number", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                kind = SheetSetFieldKind.SheetNumber;
                return true;
            }
            if (t.IndexOf("CurrentSheetTitle", StringComparison.OrdinalIgnoreCase) >= 0
                || t.IndexOf("Sheet.Title", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                kind = SheetSetFieldKind.SheetTitle;
                return true;
            }
            if (t.IndexOf("CurrentSheetRevision", StringComparison.OrdinalIgnoreCase) >= 0
                || t.IndexOf("Sheet.Revision", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                kind = SheetSetFieldKind.Revision;
                return true;
            }

            return false;
        }

        private static bool TryMapTagToFieldKind(string tag, out SheetSetFieldKind kind)
        {
            kind = default;
            if (string.IsNullOrWhiteSpace(tag)) return false;

            if (tag.StartsWith("DRAWING_NO", StringComparison.OrdinalIgnoreCase))
            {
                kind = SheetSetFieldKind.SheetNumber;
                return true;
            }
            if (tag.StartsWith("DRAWING_NAME", StringComparison.OrdinalIgnoreCase))
            {
                kind = SheetSetFieldKind.SheetTitle;
                return true;
            }
            if (tag.StartsWith("REVISION", StringComparison.OrdinalIgnoreCase)
                && tag.IndexOf("%<", StringComparison.Ordinal) < 0)
            {
                kind = SheetSetFieldKind.Revision;
                return true;
            }
            return false;
        }

        private static string MakeUniqueTag(BlockTableRecord space, Transaction tr, string baseTag)
        {
            if (space == null || string.IsNullOrWhiteSpace(baseTag))
                return baseTag ?? "FIELD";

            if (!TagExists(space, tr, baseTag))
                return baseTag;

            for (int i = 2; i < 100; i++)
            {
                string candidate = baseTag + i;
                if (!TagExists(space, tr, candidate))
                    return candidate;
            }

            return baseTag + Guid.NewGuid().ToString("N").Substring(0, 4).ToUpperInvariant();
        }

        private static bool TagExists(BlockTableRecord space, Transaction tr, string tag)
        {
            foreach (ObjectId id in space)
            {
                if (id.IsNull || id.IsErased) continue;
                if (!(tr.GetObject(id, OpenMode.ForRead, false) is AttributeDefinition ad)) continue;
                if (string.Equals(ad.Tag, tag, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }

        private static bool IsLayoutBlock(Database db, BlockTableRecord btr, Transaction tr)
        {
            if (btr == null || db == null) return false;
            var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
            foreach (DBDictionaryEntry entry in layoutDict)
            {
                var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                if (layout.BlockTableRecordId == btr.ObjectId)
                    return true;
            }
            return false;
        }
    }
}
