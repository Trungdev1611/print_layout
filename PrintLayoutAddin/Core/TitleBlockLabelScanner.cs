using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace PrintLayoutAddin.Core
{
    /// <summary>
    /// Finds title-block label texts (Sheet Title / Number / Revision) near the viewport
    /// using labels from <see cref="Config"/>.
    /// </summary>
    public static class TitleBlockLabelScanner
    {
        public sealed class Hit
        {
            public TitleBlockSetupService.SheetSetFieldKind Kind { get; set; }
            public string ConfigLabel { get; set; }
            public string FoundText { get; set; }
            public Point3d Position { get; set; }
            public ObjectId EntityId { get; set; }
            /// <summary><c>lower</c> = title/number strip; <c>mid</c> = revision band.</summary>
            public string Band { get; set; }
            public string Source { get; set; }
        }

        public sealed class ScanResult
        {
            public ViewportCornerGeometry.Bounds Viewport { get; set; }
            public ViewportCornerGeometry.Rect LowerBox { get; set; }
            public ViewportCornerGeometry.Rect MidBox { get; set; }
            public ViewportCornerGeometry.Rect FullStrip { get; set; }
            /// <summary>Texts in the strip that looked similar when a label was MISSING.</summary>
            public List<string> NearMissHints { get; } = new List<string>();
            public List<Hit> Hits { get; } = new List<Hit>();
            public Hit Title => Find(TitleBlockSetupService.SheetSetFieldKind.SheetTitle);
            public Hit Number => Find(TitleBlockSetupService.SheetSetFieldKind.SheetNumber);
            public Hit Revision => Find(TitleBlockSetupService.SheetSetFieldKind.Revision);

            Hit Find(TitleBlockSetupService.SheetSetFieldKind kind)
            {
                foreach (var h in Hits)
                    if (h.Kind == kind) return h;
                return null;
            }

            public string Summarize()
            {
                var sb = new StringBuilder();
                sb.AppendLine(ViewportCornerGeometry.Describe(Viewport));
                sb.AppendLine("  lower-scan " + LowerBox);
                sb.AppendLine("  mid-scan   " + MidBox);
                sb.AppendLine("  full-strip " + FullStrip);
                AppendHit(sb, "TITLE ", Title, Config.Instance.SheetTitleLabel);
                AppendHit(sb, "NUMBER", Number, Config.Instance.SheetNumberLabel);
                AppendHit(sb, "REV   ", Revision, Config.Instance.SheetRevisionLabel);
                if (NearMissHints.Count > 0)
                {
                    sb.AppendLine("  near-miss texts in strip (for debug):");
                    foreach (var h in NearMissHints)
                        sb.AppendLine("    · " + h);
                }
                return sb.ToString().TrimEnd();
            }

            static void AppendHit(StringBuilder sb, string tag, Hit hit, string expected)
            {
                if (hit == null)
                {
                    sb.AppendLine($"  [{tag}] MISSING  expected≈'{expected}'");
                    return;
                }
                sb.AppendLine(string.Format(
                    CultureInfo.InvariantCulture,
                    "  [{0}] OK  '{1}' @ ({2:F2},{3:F2}) band={4} src={5}",
                    tag, hit.FoundText, hit.Position.X, hit.Position.Y, hit.Band, hit.Source));
            }
        }

        /// <summary>
        /// Scan current space (and nested block refs) for the three config labels.
        /// </summary>
        public static ScanResult Scan(
            Database db,
            ViewportCornerGeometry.Bounds viewport,
            double? stripWidth = null)
        {
            double w = stripWidth ?? Config.Instance.TitleStripScanWidth;
            if (w <= 1e-9) w = ViewportCornerGeometry.DefaultTitleStripScanWidth;

            var result = new ScanResult
            {
                Viewport = viewport,
                LowerBox = ViewportCornerGeometry.TitleStripScanBox(viewport, w),
                MidBox = ViewportCornerGeometry.TitleStripMidScanBox(viewport, w),
                FullStrip = ViewportCornerGeometry.TitleStripFullScanBox(viewport, w),
            };

            if (db == null) return result;

            string titleLabel = Config.Instance.SheetTitleLabel ?? Config.DefaultSheetTitleLabel;
            string numberLabel = Config.Instance.SheetNumberLabel ?? Config.DefaultSheetNumberLabel;
            string revLabel = Config.Instance.SheetRevisionLabel ?? Config.DefaultSheetRevisionLabel;

            var stripSamples = new List<(string Text, Point3d Pos)>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
                var visited = new HashSet<ObjectId>();
                Walk(space, Matrix3d.Identity, "paper", tr, result, titleLabel, numberLabel, revLabel, visited, stripSamples);
                tr.Commit();
            }

            Deduplicate(result);

            if (result.Revision == null)
                CollectNearMisses(result, stripSamples, revLabel);

            return result;
        }

        /// <summary>
        /// Resolve placement at the geometric center of the line-bounded cell (method B).
        /// MiddleCenter text → no vertical bias (bias was pulling values to the cell floor).
        /// </summary>
        public static bool TryResolveCellCenter(
            Database db,
            Hit hit,
            ViewportCornerGeometry.Rect searchArea,
            out Point3d placeAt,
            out TitleBlockCellFinder.Cell cell,
            out string detail)
        {
            placeAt = default;
            cell = default;
            detail = "";
            if (hit == null)
            {
                detail = "no hit";
                return false;
            }

            if (!TitleBlockCellFinder.TryFindCell(
                    db, hit.Position, searchArea, out cell, out detail))
                return false;

            // Title/Number: geometric center. Revision: slightly below center so it
            // clears the pink label that sits near the top of the short cell.
            string placeMode = "center";
            if (hit.Kind == TitleBlockSetupService.SheetSetFieldKind.Revision)
            {
                const double fromBottom = 0.28;
                placeAt = new Point3d(
                    cell.Center.X,
                    cell.YMin + cell.Height * fromBottom,
                    0);
                placeMode = "revBias=" + fromBottom.ToString("F2", System.Globalization.CultureInfo.InvariantCulture);
            }
            else
            {
                placeAt = cell.Center;
            }

            detail = string.Format(
                System.Globalization.CultureInfo.InvariantCulture,
                "{0} | {1} | place=({2:F2},{3:F2}) {4}",
                cell, cell.FormatCorners(), placeAt.X, placeAt.Y, placeMode);
            return true;
        }

        /// <summary>
        /// Legacy offset below the label — only used if cell resolution fails and caller opts in.
        /// </summary>
        public static Point3d PlacementBelow(Hit hit, double textHeight)
        {
            double h = textHeight > 1e-9 ? textHeight : TitleBlockSetupService.DefaultTextHeight;
            double dy = Math.Max(h * 1.6, 2.0);
            return new Point3d(hit.Position.X, hit.Position.Y - dy, 0);
        }

        static void Deduplicate(ScanResult result)
        {
            Hit bestTitle = null, bestNumber = null, bestRev = null;
            foreach (var h in result.Hits)
            {
                switch (h.Kind)
                {
                    case TitleBlockSetupService.SheetSetFieldKind.SheetTitle:
                        if (bestTitle == null) bestTitle = h;
                        break;
                    case TitleBlockSetupService.SheetSetFieldKind.SheetNumber:
                        if (bestNumber == null) bestNumber = h;
                        break;
                    case TitleBlockSetupService.SheetSetFieldKind.Revision:
                        if (bestRev == null) bestRev = h;
                        break;
                }
            }
            result.Hits.Clear();
            if (bestTitle != null) result.Hits.Add(bestTitle);
            if (bestNumber != null) result.Hits.Add(bestNumber);
            if (bestRev != null) result.Hits.Add(bestRev);
        }

        static void CollectNearMisses(
            ScanResult result,
            List<(string Text, Point3d Pos)> samples,
            string revLabel)
        {
            string expect = NormalizeLabel(revLabel ?? "");
            string expectCompact = Compact(expect);
            foreach (var s in samples)
            {
                if (!result.FullStrip.Contains(s.Pos, margin: 2.0)) continue;
                string n = NormalizeLabel(s.Text);
                string c = Compact(n);
                bool interesting =
                    n.IndexOf("PHIÊN", StringComparison.OrdinalIgnoreCase) >= 0
                    || n.IndexOf("PHIEN", StringComparison.OrdinalIgnoreCase) >= 0
                    || n.IndexOf("BẢN", StringComparison.OrdinalIgnoreCase) >= 0
                    || n.IndexOf("BAN", StringComparison.OrdinalIgnoreCase) >= 0
                    || (expectCompact.Length > 0 && c.IndexOf(expectCompact, StringComparison.OrdinalIgnoreCase) >= 0)
                    || LabelMatches(n, revLabel);
                if (!interesting) continue;
                string line = string.Format(
                    CultureInfo.InvariantCulture,
                    "'{0}' @ ({1:F2},{2:F2}) inStrip={3} match={4}",
                    s.Text.Trim(),
                    s.Pos.X, s.Pos.Y,
                    result.FullStrip.Contains(s.Pos, 1.0),
                    LabelMatches(n, revLabel));
                if (!result.NearMissHints.Contains(line))
                    result.NearMissHints.Add(line);
                if (result.NearMissHints.Count >= 12) break;
            }
        }

        static string Compact(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var sb = new StringBuilder(s.Length);
            foreach (char ch in s)
            {
                if (!char.IsWhiteSpace(ch))
                    sb.Append(ch);
            }
            return sb.ToString();
        }

        static void Walk(
            BlockTableRecord space,
            Matrix3d toPaper,
            string source,
            Transaction tr,
            ScanResult result,
            string titleLabel,
            string numberLabel,
            string revLabel,
            HashSet<ObjectId> visitedBlocks,
            List<(string Text, Point3d Pos)> stripSamples)
        {
            if (space == null || space.ObjectId.IsNull) return;
            if (!visitedBlocks.Add(space.ObjectId)) return;

            foreach (ObjectId id in space)
            {
                if (id.IsNull || id.IsErased) continue;
                var ent = tr.GetObject(id, OpenMode.ForRead, false) as Entity;
                if (ent == null) continue;

                if (ent is DBText dbText)
                {
                    ConsiderText(
                        dbText.TextString, dbText.Position, id, toPaper, source,
                        result, titleLabel, numberLabel, revLabel, stripSamples);
                    continue;
                }

                if (ent is MText mText)
                {
                    string plain = mText.Text;
                    if (string.IsNullOrWhiteSpace(plain))
                        plain = mText.Contents ?? "";
                    ConsiderText(
                        plain, mText.Location, id, toPaper, source,
                        result, titleLabel, numberLabel, revLabel, stripSamples);
                    continue;
                }

                if (ent is AttributeDefinition attDef)
                {
                    ConsiderText(
                        attDef.Tag, attDef.Position, id, toPaper, source + "/attdef",
                        result, titleLabel, numberLabel, revLabel, stripSamples);
                    ConsiderText(
                        attDef.TextString, attDef.Position, id, toPaper, source + "/attdef",
                        result, titleLabel, numberLabel, revLabel, stripSamples);
                    continue;
                }

                if (ent is BlockReference br)
                {
                    try
                    {
                        foreach (ObjectId aid in br.AttributeCollection)
                        {
                            if (aid.IsNull || aid.IsErased) continue;
                            if (!(tr.GetObject(aid, OpenMode.ForRead, false) is AttributeReference ar))
                                continue;
                            ConsiderText(
                                ar.TextString, ar.Position, aid, Matrix3d.Identity,
                                source + "/attr:" + (ar.Tag ?? ""),
                                result, titleLabel, numberLabel, revLabel, stripSamples);
                        }

                        var btr = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                        Matrix3d nested = toPaper * br.BlockTransform;
                        string childSrc = source + "/" + (br.Name ?? "block");
                        Walk(btr, nested, childSrc, tr, result, titleLabel, numberLabel, revLabel, visitedBlocks, stripSamples);
                    }
                    catch
                    {
                        // Skip broken refs.
                    }
                }
            }
        }

        static void ConsiderText(
            string raw,
            Point3d localPos,
            ObjectId id,
            Matrix3d toPaper,
            string source,
            ScanResult result,
            string titleLabel,
            string numberLabel,
            string revLabel,
            List<(string Text, Point3d Pos)> stripSamples)
        {
            if (string.IsNullOrWhiteSpace(raw)) return;
            string plain = NormalizeLabel(raw);
            if (plain.Length == 0) return;

            Point3d paperPos;
            try { paperPos = localPos.TransformBy(toPaper); }
            catch { paperPos = localPos; }

            if (stripSamples != null && result.FullStrip.Contains(paperPos, margin: 5.0))
                stripSamples.Add((raw, paperPos));

            if (LabelMatches(plain, titleLabel) && result.LowerBox.Contains(paperPos, margin: 1.0))
            {
                result.Hits.Add(new Hit
                {
                    Kind = TitleBlockSetupService.SheetSetFieldKind.SheetTitle,
                    ConfigLabel = titleLabel,
                    FoundText = raw.Trim(),
                    Position = paperPos,
                    EntityId = id,
                    Band = "lower",
                    Source = source,
                });
            }

            if (LabelMatches(plain, numberLabel) && result.LowerBox.Contains(paperPos, margin: 1.0))
            {
                result.Hits.Add(new Hit
                {
                    Kind = TitleBlockSetupService.SheetSetFieldKind.SheetNumber,
                    ConfigLabel = numberLabel,
                    FoundText = raw.Trim(),
                    Position = paperPos,
                    EntityId = id,
                    Band = "lower",
                    Source = source,
                });
            }

            // PHIÊN BẢN often sits just above the big title/number cells — still below Y_mid.
            // Accept mid band OR full strip (unique string; will not match Rev table headers).
            bool revInBand = result.MidBox.Contains(paperPos, margin: 2.0)
                || result.FullStrip.Contains(paperPos, margin: 2.0);
            if (LabelMatches(plain, revLabel) && revInBand)
            {
                string band = result.MidBox.Contains(paperPos, margin: 2.0) ? "mid" : "strip";
                result.Hits.Add(new Hit
                {
                    Kind = TitleBlockSetupService.SheetSetFieldKind.Revision,
                    ConfigLabel = revLabel,
                    FoundText = raw.Trim(),
                    Position = paperPos,
                    EntityId = id,
                    Band = band,
                    Source = source,
                });
            }
        }

        /// <summary>Strip field codes / MText markup lightly, trim, drop trailing ':'.</summary>
        public static string NormalizeLabel(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return "";
            string t = raw.Trim();
            try { t = t.Normalize(NormalizationForm.FormC); } catch { }
            // Drop simple MText wrappers if Contents leaked through.
            if (t.IndexOf('\\') >= 0)
            {
                t = t.Replace("\\P", " ").Replace("\\p", " ");
                var sb = new StringBuilder(t.Length);
                bool inCmd = false;
                for (int i = 0; i < t.Length; i++)
                {
                    char c = t[i];
                    if (c == '\\') { inCmd = true; continue; }
                    if (inCmd)
                    {
                        if (c == ';' || char.IsWhiteSpace(c)) inCmd = false;
                        continue;
                    }
                    if (c == '{' || c == '}') continue;
                    sb.Append(c);
                }
                t = sb.ToString().Trim();
            }
            while (t.EndsWith(":", StringComparison.Ordinal) || t.EndsWith("：", StringComparison.Ordinal))
                t = t.Substring(0, t.Length - 1).TrimEnd();
            // Collapse all whitespace runs (incl. NBSP) to a single space.
            var parts = t.Split((char[])null, StringSplitOptions.RemoveEmptyEntries);
            t = string.Join(" ", parts);
            return t;
        }

        public static bool LabelMatches(string normalizedFound, string configLabel)
        {
            string expect = NormalizeLabel(configLabel ?? "");
            string found = NormalizeLabel(normalizedFound ?? "");
            if (expect.Length == 0 || found.Length == 0) return false;
            if (string.Equals(found, expect, StringComparison.OrdinalIgnoreCase)
                || found.StartsWith(expect, StringComparison.OrdinalIgnoreCase))
                return true;
            // Ignore spacing differences: "PHIÊN  BẢN" vs "PHIÊN BẢN"
            string fc = Compact(found);
            string ec = Compact(expect);
            return string.Equals(fc, ec, StringComparison.OrdinalIgnoreCase)
                || fc.StartsWith(ec, StringComparison.OrdinalIgnoreCase);
        }
    }
}
