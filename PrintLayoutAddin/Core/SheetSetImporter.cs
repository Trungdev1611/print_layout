using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace PrintLayoutAddin.Core
{
    public class SheetSetImportRow
    {
        public string Layout { get; set; } = "";
        public string Dwg { get; set; } = "";
        public string SheetNumber { get; set; }
        public string Title { get; set; }
        public string Revision { get; set; }
        public bool IsSubset { get; set; }
        public int SourceLine { get; set; }
    }

    public class SheetSetImportResult
    {
        public List<SheetSetImportRow> Rows { get; } = new List<SheetSetImportRow>();
        public List<string> Errors { get; } = new List<string>();
        public List<string> Warnings { get; } = new List<string>();
        public bool Ok => Errors.Count == 0 && Rows.Count > 0;
    }

    public class SheetSetImportApplyResult
    {
        public int Updated { get; set; }
        public int Unmatched { get; set; }
        public int SkippedSubset { get; set; }
        public List<string> Warnings { get; } = new List<string>();
        public List<string> Errors { get; } = new List<string>();
        public bool Ok => Errors.Count == 0;
        public string Summary =>
            $"Updated {Updated} sheet(s), unmatched {Unmatched}, skipped subset {SkippedSubset}.";
    }

    /// <summary>
    /// Import Sheet Number / Title / Revision into existing dialog rows.
    /// Match key = Layout (+ DWG file name when present). Never adds layouts/rows.
    /// </summary>
    public static class SheetSetImporter
    {
        public const string ExportNote =
            "NOTE: Only edit SheetNumber, Title, Revision. Do NOT change Order, Kind, Subset, Layout, DWG "
            + "(those columns are match/reference only). Import updates matching sheets; unmatched rows stay unchanged.";

        private static readonly string[] LayoutAliases =
            { "layout", "layoutname", "layout_name" };
        private static readonly string[] DwgAliases =
            { "dwg", "dwgname", "dwg_name", "file", "filename", "drawingfile" };
        private static readonly string[] SheetNumberAliases =
            { "sheetnumber", "sheet_number", "sheetno", "number", "no" };
        private static readonly string[] TitleAliases =
            { "title", "sheettitle", "sheet_title", "drawingname", "drawing_name", "namedrawing" };
        private static readonly string[] RevisionAliases =
            { "revision", "rev", "revno", "revisionnumber" };
        private static readonly string[] KindAliases =
            { "kind", "type", "rowkind" };

        public static SheetSetImportResult ReadFile(string path)
        {
            var result = new SheetSetImportResult();
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                result.Errors.Add("File not found: " + path);
                return result;
            }

            try
            {
                var ext = Path.GetExtension(path)?.ToLowerInvariant();
                List<string[]> table;
                if (ext == ".xlsx") table = ReadXlsx(path);
                else if (ext == ".csv" || ext == ".txt") table = ReadCsv(path);
                else
                {
                    result.Errors.Add("Unsupported file type. Use .xlsx or .csv.");
                    return result;
                }
                return ParseTable(table);
            }
            catch (Exception ex)
            {
                result.Errors.Add("Failed to read file: " + ex.Message);
                return result;
            }
        }

        /// <summary>
        /// Merge import into existing entries. Subset rows ignored. No new entries created.
        /// </summary>
        public static SheetSetImportApplyResult ApplyToEntries(
            IList<SheetSetEntry> entries,
            SheetSetImportResult import)
        {
            var apply = new SheetSetImportApplyResult();
            if (import == null)
            {
                apply.Errors.Add("Nothing to import.");
                return apply;
            }
            foreach (var e in import.Errors)
                apply.Errors.Add(e);
            if (apply.Errors.Count > 0) return apply;
            if (import.Rows.Count == 0)
            {
                apply.Errors.Add("No data rows found.");
                return apply;
            }

            var sheets = (entries ?? Array.Empty<SheetSetEntry>())
                .Where(e => e != null && !e.IsSubset)
                .ToList();

            foreach (var warn in import.Warnings)
                apply.Warnings.Add(warn);

            foreach (var row in import.Rows)
            {
                if (row == null) continue;
                if (row.IsSubset)
                {
                    apply.SkippedSubset++;
                    continue;
                }

                string layout = (row.Layout ?? "").Trim();
                if (string.IsNullOrWhiteSpace(layout))
                {
                    apply.Warnings.Add($"Row {row.SourceLine}: empty Layout — skipped.");
                    apply.Unmatched++;
                    continue;
                }

                string dwg = NormalizeDwgKey(row.Dwg);
                var candidates = sheets
                    .Where(e => string.Equals(e.LayoutName ?? "", layout, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (candidates.Count == 0)
                {
                    apply.Unmatched++;
                    apply.Warnings.Add($"Row {row.SourceLine}: layout '{layout}' not in table — left unchanged.");
                    continue;
                }

                SheetSetEntry target = null;
                if (!string.IsNullOrEmpty(dwg))
                {
                    target = candidates.FirstOrDefault(e =>
                        string.Equals(NormalizeDwgKey(e.DwgName), dwg, StringComparison.OrdinalIgnoreCase));
                    if (target == null)
                    {
                        apply.Unmatched++;
                        apply.Warnings.Add(
                            $"Row {row.SourceLine}: layout '{layout}' + DWG '{row.Dwg}' not matched — left unchanged.");
                        continue;
                    }
                }
                else if (candidates.Count == 1)
                {
                    target = candidates[0];
                }
                else
                {
                    apply.Unmatched++;
                    apply.Warnings.Add(
                        $"Row {row.SourceLine}: layout '{layout}' matches {candidates.Count} sheets — "
                        + "set DWG column to disambiguate. Left unchanged.");
                    continue;
                }

                if (row.SheetNumber != null) target.SheetNumber = row.SheetNumber;
                if (row.Title != null) target.Title = row.Title;
                if (row.Revision != null) target.Revision = row.Revision;
                apply.Updated++;
            }

            return apply;
        }

        private static string NormalizeDwgKey(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return "";
            try { return Path.GetFileName(value.Trim()); }
            catch { return value.Trim(); }
        }

        private static SheetSetImportResult ParseTable(List<string[]> rows)
        {
            var result = new SheetSetImportResult();
            if (rows == null || rows.Count == 0)
            {
                result.Errors.Add("File is empty.");
                return result;
            }

            int headerIdx = -1;
            int colLayout = -1, colDwg = -1, colNumber = -1, colTitle = -1, colRev = -1, colKind = -1;

            for (int i = 0; i < rows.Count; i++)
            {
                var header = rows[i];
                if (header == null || header.All(c => string.IsNullOrWhiteSpace(c))) continue;

                // Skip instruction / NOTE row(s).
                string first = (header.Length > 0 ? header[0] : "") ?? "";
                if (LooksLikeNoteRow(first, header)) continue;

                int layout = -1, dwg = -1, number = -1, title = -1, rev = -1, kind = -1;
                for (int c = 0; c < header.Length; c++)
                {
                    var norm = NormalizeHeader(header[c]);
                    if (layout < 0 && MatchesAny(norm, LayoutAliases)) layout = c;
                    else if (dwg < 0 && MatchesAny(norm, DwgAliases)) dwg = c;
                    else if (number < 0 && MatchesAny(norm, SheetNumberAliases)) number = c;
                    else if (title < 0 && MatchesAny(norm, TitleAliases)) title = c;
                    else if (rev < 0 && MatchesAny(norm, RevisionAliases)) rev = c;
                    else if (kind < 0 && MatchesAny(norm, KindAliases)) kind = c;
                }

                if (layout >= 0 && (number >= 0 || title >= 0 || rev >= 0))
                {
                    headerIdx = i;
                    colLayout = layout;
                    colDwg = dwg;
                    colNumber = number;
                    colTitle = title;
                    colRev = rev;
                    colKind = kind;
                    break;
                }
            }

            if (headerIdx < 0)
            {
                result.Errors.Add(
                    "Header row not found. Need a Layout column plus SheetNumber and/or Title and/or Revision.");
                return result;
            }

            if (colNumber < 0 && colTitle < 0 && colRev < 0)
            {
                result.Errors.Add("No editable columns found (SheetNumber / Title / Revision).");
                return result;
            }

            int line = 0;
            for (int i = headerIdx + 1; i < rows.Count; i++)
            {
                line++;
                var row = rows[i];
                if (row == null || row.All(c => string.IsNullOrWhiteSpace(c))) continue;

                string kind = Cell(row, colKind);
                bool isSubset = kind.Equals("Subset", StringComparison.OrdinalIgnoreCase);

                // Null means "column absent — do not touch"; empty string clears the field.
                string number = colNumber >= 0 ? Cell(row, colNumber) : null;
                string title = colTitle >= 0 ? Cell(row, colTitle) : null;
                string rev = colRev >= 0 ? Cell(row, colRev) : null;

                result.Rows.Add(new SheetSetImportRow
                {
                    Layout = Cell(row, colLayout),
                    Dwg = Cell(row, colDwg),
                    SheetNumber = number,
                    Title = title,
                    Revision = rev,
                    IsSubset = isSubset,
                    SourceLine = line,
                });
            }

            if (result.Rows.Count == 0)
                result.Errors.Add("No data rows found under the header.");
            return result;
        }

        private static bool LooksLikeNoteRow(string first, string[] header)
        {
            string f = (first ?? "").Trim();
            if (f.StartsWith("NOTE", StringComparison.OrdinalIgnoreCase)) return true;
            if (f.StartsWith("#")) return true;
            if (f.IndexOf("Only edit", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            // Single long cell, no Layout header-like tokens.
            int nonEmpty = header.Count(c => !string.IsNullOrWhiteSpace(c));
            if (nonEmpty <= 1 && f.Length > 40) return true;
            return false;
        }

        private static string Cell(string[] row, int col)
        {
            if (col < 0 || row == null || col >= row.Length) return "";
            return (row[col] ?? "").Trim();
        }

        private static string NormalizeHeader(string h)
        {
            if (h == null) return "";
            var sb = new StringBuilder();
            foreach (var ch in h)
            {
                if (char.IsLetterOrDigit(ch)) sb.Append(char.ToLowerInvariant(ch));
            }
            return sb.ToString();
        }

        private static bool MatchesAny(string normalizedHeader, string[] aliases)
        {
            foreach (var a in aliases)
                if (normalizedHeader == NormalizeHeader(a)) return true;
            return false;
        }

        private static List<string[]> ReadCsv(string path)
        {
            var rows = new List<string[]>();
            using (var reader = new StreamReader(path, Encoding.UTF8, detectEncodingFromByteOrderMarks: true))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                    rows.Add(SplitCsvLine(line));
            }
            return rows;
        }

        private static string[] SplitCsvLine(string line)
        {
            if (string.IsNullOrEmpty(line)) return Array.Empty<string>();
            var result = new List<string>();
            var sb = new StringBuilder();
            bool inQuotes = false;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (inQuotes)
                {
                    if (c == '"')
                    {
                        if (i + 1 < line.Length && line[i + 1] == '"') { sb.Append('"'); i++; }
                        else inQuotes = false;
                    }
                    else sb.Append(c);
                }
                else
                {
                    if (c == ',') { result.Add(sb.ToString()); sb.Clear(); }
                    else if (c == '"' && sb.Length == 0) inQuotes = true;
                    else sb.Append(c);
                }
            }
            result.Add(sb.ToString());
            return result.ToArray();
        }

        private static List<string[]> ReadXlsx(string path)
        {
            using (var archive = ZipFile.OpenRead(path))
            {
                var shared = new List<string>();
                var ssEntry = archive.GetEntry("xl/sharedStrings.xml");
                if (ssEntry != null)
                {
                    using (var s = ssEntry.Open())
                    {
                        var doc = XDocument.Load(s);
                        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                        foreach (var si in doc.Root.Elements(ns + "si"))
                        {
                            var val = string.Concat(si.Descendants(ns + "t").Select(t => t.Value));
                            shared.Add(val);
                        }
                    }
                }

                var sheetEntry = archive.GetEntry("xl/worksheets/sheet1.xml")
                    ?? archive.Entries.FirstOrDefault(e =>
                        e.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)
                        && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));
                if (sheetEntry == null)
                    throw new InvalidOperationException("No worksheet found inside the XLSX file.");

                var rows = new List<string[]>();
                using (var s = sheetEntry.Open())
                {
                    var doc = XDocument.Load(s);
                    XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                    foreach (var rowEl in doc.Descendants(ns + "row"))
                    {
                        var cells = new List<(int Col, string Val)>();
                        foreach (var c in rowEl.Elements(ns + "c"))
                        {
                            var cellRef = (string)c.Attribute("r") ?? "A";
                            int col = ParseColumnIndex(cellRef);
                            var t = (string)c.Attribute("t");
                            string val = "";
                            if (t == "s")
                            {
                                var v = c.Element(ns + "v")?.Value;
                                if (int.TryParse(v, out int idx) && idx >= 0 && idx < shared.Count)
                                    val = shared[idx];
                            }
                            else if (t == "inlineStr")
                            {
                                val = string.Concat(
                                    c.Element(ns + "is")?.Descendants(ns + "t").Select(x => x.Value)
                                    ?? Enumerable.Empty<string>());
                            }
                            else
                            {
                                val = c.Element(ns + "v")?.Value ?? "";
                            }
                            cells.Add((col, val));
                        }
                        int maxCol = cells.Count == 0 ? 0 : cells.Max(x => x.Col);
                        var arr = new string[maxCol];
                        foreach (var (col, val) in cells) arr[col - 1] = val;
                        rows.Add(arr);
                    }
                }
                return rows;
            }
        }

        private static int ParseColumnIndex(string cellRef)
        {
            int col = 0;
            foreach (var ch in cellRef)
            {
                if (ch >= 'A' && ch <= 'Z') col = col * 26 + (ch - 'A' + 1);
                else if (ch >= 'a' && ch <= 'z') col = col * 26 + (ch - 'a' + 1);
                else break;
            }
            return col < 1 ? 1 : col;
        }
    }
}
