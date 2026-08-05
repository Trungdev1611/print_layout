using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace PrintLayoutAddin.Core
{
    // ============================================================
    // NumberImporter — reads frame codes from CSV or XLSX files and
    // exports a blank template that users can fill in.
    //
    // Zero external dependencies: XLSX parsing uses System.IO.Compression
    // + System.Xml.Linq (built into .NET Framework 4.8).
    // ============================================================

    public class ImportedRow
    {
        public int? Order;            // nullable because the Order column is optional
        public string FrameNumber = "";
        public string DrawingName = "";
        public string Note = "";
        public int SourceLine;        // 1-based for error messages. Excludes header.
    }

    public class ImportResult
    {
        public List<ImportedRow> Rows = new List<ImportedRow>();
        public List<string> Errors = new List<string>();
        public List<string> Warnings = new List<string>();
        public bool HasOrderColumn;
        public bool HasDrawingNameColumn;
        public bool HasNoteColumn;
        public bool Ok => Errors.Count == 0 && Rows.Count > 0;
    }

    public static class NumberImporter
    {
        // -- Header matching --------------------------------------------------
        // Accept common variants for each column name so users don't have to
        // copy the exact header. Matching is case- and separator-insensitive.
        private static readonly string[] FrameNumberAliases =
            { "framenumber", "frameno", "frame_no", "frame#", "frame", "code", "number", "stt", "inno-stt" };
        private static readonly string[] OrderAliases =
            { "order", "stt#", "index", "idx", "no" };
        private static readonly string[] DrawingNameAliases =
            { "drawingname", "drawing_name", "namedrawing", "name_drawing", "inno_name_drawing", "tênbảnvẽ", "tenbanve" };
        private static readonly string[] NoteAliases =
            { "note", "notes", "comment", "description", "desc" };

        /// <summary>Alias matching the upgrade-spec naming (public wrapper).</summary>
        public static string NormalizeImportHeaders(string h) => NormalizeHeader(h);

        private static string NormalizeHeader(string h)
        {
            if (h == null) return "";
            // strip whitespace, underscore, dash, hash; lower-case.
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

        // -- Spec-name aliases --------------------------------------------------
        /// <summary>Alias of <see cref="ImportFromXlsx"/> matching the upgrade-spec naming.</summary>
        public static ImportResult ImportNumbersFromExcel(string path) => ImportFromXlsx(path);
        /// <summary>Alias of <see cref="ImportFromCsv"/> matching the upgrade-spec naming.</summary>
        public static ImportResult ImportNumbersFromCsv(string path) => ImportFromCsv(path);
        /// <summary>Alias of <see cref="ValidateImported"/> matching the upgrade-spec naming.</summary>
        public static ValidationResult ValidateImportedNumberList(
            ImportResult import, int? expectedFrameCount = null, bool allowDuplicates = false)
            => ValidateImported(import, expectedFrameCount, allowDuplicates);
        /// <summary>
        /// Alias matching the upgrade-spec naming. Dispatches to CSV or XLSX exporter
        /// based on the extension of <paramref name="path"/>.
        /// </summary>
        public static void ExportNumberingTemplate(string path, int rowCount)
        {
            var ext = (System.IO.Path.GetExtension(path) ?? "").ToLowerInvariant();
            if (ext == ".xlsx") ExportTemplateXlsx(path, rowCount);
            else ExportTemplateCsv(path, rowCount);
        }

        // -- Public entry point: auto-detect by extension --------------------
        public static ImportResult ImportFromFile(string path)
        {
            var r = new ImportResult();
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                r.Errors.Add($"File not found: {path}");
                return r;
            }
            var ext = Path.GetExtension(path).ToLowerInvariant();
            try
            {
                if (ext == ".csv" || ext == ".txt") return ImportFromCsv(path);
                if (ext == ".xlsx") return ImportFromXlsx(path);
                r.Errors.Add($"Unsupported file type '{ext}'. Use .csv or .xlsx.");
            }
            catch (Exception ex)
            {
                r.Errors.Add($"Failed to read '{Path.GetFileName(path)}': {ex.Message}");
            }
            return r;
        }

        // -- CSV reader -------------------------------------------------------
        // UTF-8 with BOM auto-detect. Handles quoted fields, embedded commas,
        // escaped quotes ("") inside quoted fields.
        public static ImportResult ImportFromCsv(string path)
        {
            var rows = new List<string[]>();
            using (var reader = new StreamReader(path, Encoding.UTF8, detectEncodingFromByteOrderMarks: true))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    rows.Add(SplitCsvLine(line));
                }
            }
            return ParseTabular(rows);
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

        // -- XLSX reader ------------------------------------------------------
        // Minimal: reads the first worksheet, resolves sharedStrings, handles
        // inlineStr cells. Good enough for flat data tables.
        public static ImportResult ImportFromXlsx(string path)
        {
            using (var archive = ZipFile.OpenRead(path))
            {
                // 1. Read shared strings table (optional — files without any strings skip this)
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
                            // An <si> may contain one <t> or multiple runs with <r><t>...</t></r>.
                            // Concatenating all descendant <t> values is the standard simplification.
                            var val = string.Concat(si.Descendants(ns + "t").Select(t => t.Value));
                            shared.Add(val);
                        }
                    }
                }

                // 2. Pick the first worksheet. Some producers don't name it sheet1.xml.
                var sheetEntry = archive.GetEntry("xl/worksheets/sheet1.xml")
                                 ?? archive.Entries.FirstOrDefault(e =>
                                        e.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)
                                        && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));

                if (sheetEntry == null)
                {
                    var r = new ImportResult();
                    r.Errors.Add("No worksheet found inside the XLSX file.");
                    return r;
                }

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
                                // shared-string reference
                                var v = c.Element(ns + "v")?.Value;
                                if (int.TryParse(v, out int idx) && idx >= 0 && idx < shared.Count)
                                    val = shared[idx];
                            }
                            else if (t == "inlineStr")
                            {
                                val = string.Concat(c.Element(ns + "is")?.Descendants(ns + "t").Select(x => x.Value) ?? Enumerable.Empty<string>());
                            }
                            else
                            {
                                // numeric, boolean, date — take the raw <v> text as string.
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
                return ParseTabular(rows);
            }
        }

        // "A1" -> 1, "B1" -> 2, "AA1" -> 27, ...
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

        // -- Common: row array -> ImportResult --------------------------------
        // Expects the first non-empty row to be a header. Locates the
        // FrameNumber column (required) plus optional Order/Note columns.
        private static ImportResult ParseTabular(List<string[]> rows)
        {
            var result = new ImportResult();
            if (rows == null || rows.Count == 0)
            {
                result.Errors.Add("File is empty.");
                return result;
            }

            // Find header row (first row with at least one non-empty cell).
            int headerIdx = -1;
            for (int i = 0; i < rows.Count; i++)
            {
                if (rows[i] != null && rows[i].Any(c => !string.IsNullOrWhiteSpace(c)))
                {
                    headerIdx = i;
                    break;
                }
            }
            if (headerIdx < 0)
            {
                result.Errors.Add("No header row found.");
                return result;
            }

            var header = rows[headerIdx];
            int colFrame = -1, colOrder = -1, colDrawingName = -1, colNote = -1;
            for (int c = 0; c < header.Length; c++)
            {
                var norm = NormalizeHeader(header[c]);
                if (colFrame < 0 && MatchesAny(norm, FrameNumberAliases)) colFrame = c;
                else if (colOrder < 0 && MatchesAny(norm, OrderAliases)) colOrder = c;
                else if (colDrawingName < 0 && MatchesAny(norm, DrawingNameAliases)) colDrawingName = c;
                else if (colNote < 0 && MatchesAny(norm, NoteAliases)) colNote = c;
            }

            if (colFrame < 0)
            {
                result.Errors.Add(
                    "FrameNumber column not found. Accepted headers (case-insensitive): "
                    + string.Join(", ", FrameNumberAliases));
                return result;
            }

            result.HasOrderColumn = colOrder >= 0;
            result.HasDrawingNameColumn = colDrawingName >= 0;
            result.HasNoteColumn = colNote >= 0;

            int dataLine = 0;
            for (int i = headerIdx + 1; i < rows.Count; i++)
            {
                dataLine++;
                var row = rows[i];
                if (row == null || row.All(c => string.IsNullOrWhiteSpace(c))) continue; // skip blank rows

                string frame = colFrame < row.Length ? (row[colFrame] ?? "").Trim() : "";
                string drawingName = (colDrawingName >= 0 && colDrawingName < row.Length)
                    ? (row[colDrawingName] ?? "").Trim()
                    : "";
                string note = (colNote >= 0 && colNote < row.Length) ? (row[colNote] ?? "").Trim() : "";
                int? order = null;
                if (colOrder >= 0 && colOrder < row.Length)
                {
                    var rawOrder = (row[colOrder] ?? "").Trim();
                    if (rawOrder.Length > 0)
                    {
                        if (int.TryParse(rawOrder, out int ord)) order = ord;
                        else result.Warnings.Add($"Row {dataLine}: Order '{rawOrder}' is not an integer, ignored.");
                    }
                }

                result.Rows.Add(new ImportedRow
                {
                    Order = order,
                    FrameNumber = frame,
                    DrawingName = drawingName,
                    Note = note,
                    SourceLine = dataLine
                });
            }

            if (result.Rows.Count == 0)
            {
                result.Errors.Add("No data rows found (file contains only a header).");
                return result;
            }

            // If Order column is present, sort by it. Missing Order values go to the end, keeping original order.
            if (result.HasOrderColumn)
            {
                var ordered = result.Rows
                    .Select((r, idx) => new { r, idx })
                    .OrderBy(x => x.r.Order.HasValue ? 0 : 1)    // rows with Order first
                    .ThenBy(x => x.r.Order ?? int.MaxValue)
                    .ThenBy(x => x.idx)
                    .Select(x => x.r)
                    .ToList();

                // Detect duplicate Order values
                var dupOrders = result.Rows
                    .Where(r => r.Order.HasValue)
                    .GroupBy(r => r.Order.Value)
                    .Where(g => g.Count() > 1)
                    .Select(g => g.Key)
                    .ToList();
                foreach (var d in dupOrders)
                    result.Errors.Add($"Duplicate Order value: {d}.");

                // Warn about missing Order values (some rows have, some don't)
                int withOrder = result.Rows.Count(r => r.Order.HasValue);
                if (withOrder > 0 && withOrder < result.Rows.Count)
                    result.Warnings.Add(
                        $"{result.Rows.Count - withOrder} row(s) have a blank Order value; they are appended after numbered rows.");

                result.Rows = ordered;
            }

            return result;
        }

        // -- Validate imported rows against the selected-frame count ----------
        public static ValidationResult ValidateImported(
            ImportResult import,
            int? expectedFrameCount = null,
            bool allowDuplicates = false)
        {
            var r = new ValidationResult();
            if (import == null || import.Rows.Count == 0)
            {
                r.Errors.Add("Nothing to validate.");
                return r;
            }

            // Forward any errors/warnings already discovered during parsing.
            foreach (var e in import.Errors) r.Errors.Add(e);
            foreach (var w in import.Warnings) r.Warnings.Add(w);

            // Empty FrameNumber check — per row.
            for (int i = 0; i < import.Rows.Count; i++)
            {
                if (string.IsNullOrWhiteSpace(import.Rows[i].FrameNumber))
                    r.Errors.Add($"Row {import.Rows[i].SourceLine}: FrameNumber is empty.");
            }

            // Duplicate check.
            if (!allowDuplicates)
            {
                var dups = import.Rows
                    .GroupBy(x => (x.FrameNumber ?? "").Trim(), StringComparer.OrdinalIgnoreCase)
                    .Where(g => !string.IsNullOrEmpty(g.Key) && g.Count() > 1)
                    .Select(g => g.Key)
                    .ToList();
                foreach (var d in dups) r.Errors.Add($"Duplicate FrameNumber: '{d}'.");
            }

            // Count match.
            if (expectedFrameCount.HasValue && expectedFrameCount.Value != import.Rows.Count)
            {
                r.Errors.Add(
                    $"Count mismatch: {expectedFrameCount.Value} frame(s) targeted but {import.Rows.Count} code(s) in file.");
            }
            return r;
        }

        // -- Export blank (or pre-seeded) CSV template ------------------------
        public static void ExportTemplateCsv(string path, int rowCount)
        {
            var sb = new StringBuilder();
            // UTF-8 BOM so Excel opens it with correct encoding
            sb.Append('\uFEFF');
            sb.AppendLine("Order,FrameNumber,DrawingName,Note");
            int n = Math.Max(0, rowCount);
            for (int i = 1; i <= n; i++) sb.AppendLine($"{i},,,");
            File.WriteAllText(path, sb.ToString(), new UTF8Encoding(true));
        }

        // -- Export blank (or pre-seeded) XLSX template -----------------------
        // Hand-rolled minimal OOXML writer so we don't pull in ClosedXML/EPPlus
        // for a 3-column template. Uses inlineStr for text cells to avoid the
        // shared-strings table entirely.
        public static void ExportTemplateXlsx(string path, int rowCount)
        {
            if (File.Exists(path)) File.Delete(path);
            using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
            {
                AddEntry(archive, "[Content_Types].xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
                        "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
                    "</Types>");

                AddEntry(archive, "_rels/.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                    "</Relationships>");

                AddEntry(archive, "xl/workbook.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                        "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>" +
                    "</workbook>");

                AddEntry(archive, "xl/_rels/workbook.xml.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
                    "</Relationships>");

                // Build sheet1.xml
                var sb = new StringBuilder();
                sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>");
                sb.Append("<row r=\"1\">");
                sb.Append(InlineStrCell("A1", "Order"));
                sb.Append(InlineStrCell("B1", "FrameNumber"));
                sb.Append(InlineStrCell("C1", "DrawingName"));
                sb.Append(InlineStrCell("D1", "Note"));
                sb.Append("</row>");

                int n = Math.Max(0, rowCount);
                for (int i = 1; i <= n; i++)
                {
                    int r = i + 1;
                    sb.Append($"<row r=\"{r}\">");
                    sb.Append($"<c r=\"A{r}\"><v>{i}</v></c>");
                    sb.Append(InlineStrCell($"B{r}", ""));
                    sb.Append(InlineStrCell($"C{r}", ""));
                    sb.Append(InlineStrCell($"D{r}", ""));
                    sb.Append("</row>");
                }
                sb.Append("</sheetData></worksheet>");
                AddEntry(archive, "xl/worksheets/sheet1.xml", sb.ToString());
            }
        }

        private static void AddEntry(ZipArchive archive, string fullName, string content)
        {
            var entry = archive.CreateEntry(fullName);
            using (var s = entry.Open())
            using (var w = new StreamWriter(s, new UTF8Encoding(false)))
            {
                w.Write(content);
            }
        }

        private static string InlineStrCell(string cellRef, string text)
        {
            // XML-escape the value so commas/ampersands/quotes don't break the file.
            var escaped = System.Security.SecurityElement.Escape(text ?? "") ?? "";
            return $"<c r=\"{cellRef}\" t=\"inlineStr\"><is><t xml:space=\"preserve\">{escaped}</t></is></c>";
        }
    }
}
