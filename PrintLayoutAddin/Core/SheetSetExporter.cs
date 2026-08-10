using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

namespace PrintLayoutAddin.Core
{
    /// <summary>
    /// Exports the Sheet Set dialog table to .xlsx / .csv (same columns as the UI).
    /// First row is an instruction note for Import Excel users.
    /// </summary>
    public static class SheetSetExporter
    {
        // Cell style indexes in xl/styles.xml (0 = default).
        private const int StyleNote = 1;    // red italic
        private const int StyleHeader = 2;  // bold

        public static void Export(string path, IEnumerable<SheetSetEntry> entries)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("Output path is required.", nameof(path));

            var rows = (entries ?? Enumerable.Empty<SheetSetEntry>())
                .Where(e => e != null)
                .OrderBy(e => e.Order)
                .ToList();

            var ext = Path.GetExtension(path)?.ToLowerInvariant();
            if (ext == ".xlsx") ExportXlsx(path, rows);
            else if (ext == ".csv" || ext == ".txt") ExportCsv(path, rows);
            else
                throw new InvalidOperationException("Unsupported file type. Use .xlsx or .csv.");
        }

        public static void ExportCsv(string path, IList<SheetSetEntry> rows)
        {
            var sb = new StringBuilder();
            sb.Append('\uFEFF');
            sb.AppendLine(Csv(SheetSetImporter.ExportNote));
            sb.AppendLine("Order,Kind,Subset,SheetNumber,Title,Revision,Layout,DWG");
            foreach (var e in rows)
            {
                sb.Append(e.Order).Append(',');
                sb.Append(Csv(e.IsSubset ? "Subset" : "Sheet")).Append(',');
                sb.Append(Csv(e.IsSubset ? (e.Title ?? e.SubsetName) : e.SubsetName)).Append(',');
                sb.Append(Csv(e.IsSubset ? "" : e.SheetNumber)).Append(',');
                sb.Append(Csv(e.IsSubset ? "" : e.Title)).Append(',');
                sb.Append(Csv(e.IsSubset ? "" : e.Revision)).Append(',');
                sb.Append(Csv(e.IsSubset ? (e.Title ?? e.SubsetName) : e.LayoutName)).Append(',');
                sb.Append(Csv(e.DwgName));
                sb.AppendLine();
            }
            File.WriteAllText(path, sb.ToString(), new UTF8Encoding(true));
        }

        public static void ExportXlsx(string path, IList<SheetSetEntry> rows)
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
                        "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>" +
                    "</Types>");

                AddEntry(archive, "_rels/.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                    "</Relationships>");

                AddEntry(archive, "xl/workbook.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                        "<sheets><sheet name=\"SheetSet\" sheetId=\"1\" r:id=\"rId1\"/></sheets>" +
                    "</workbook>");

                AddEntry(archive, "xl/_rels/workbook.xml.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
                        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>" +
                    "</Relationships>");

                AddEntry(archive, "xl/styles.xml", BuildStylesXml());

                var sb = new StringBuilder();
                sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
                sb.Append("<cols>");
                sb.Append("<col min=\"1\" max=\"1\" width=\"8\" customWidth=\"1\"/>");   // Order
                sb.Append("<col min=\"2\" max=\"2\" width=\"10\" customWidth=\"1\"/>");  // Kind
                sb.Append("<col min=\"3\" max=\"3\" width=\"12\" customWidth=\"1\"/>");  // Subset
                sb.Append("<col min=\"4\" max=\"4\" width=\"14\" customWidth=\"1\"/>");  // SheetNumber
                sb.Append("<col min=\"5\" max=\"5\" width=\"22\" customWidth=\"1\"/>");  // Title
                sb.Append("<col min=\"6\" max=\"6\" width=\"10\" customWidth=\"1\"/>");  // Revision
                sb.Append("<col min=\"7\" max=\"7\" width=\"12\" customWidth=\"1\"/>");  // Layout
                sb.Append("<col min=\"8\" max=\"8\" width=\"36\" customWidth=\"1\"/>");  // DWG
                sb.Append("</cols>");
                sb.Append("<sheetData>");

                // Row 1 — instruction (red italic, vertical center). No merge / wrap.
                sb.Append("<row r=\"1\" ht=\"30\" customHeight=\"1\">");
                sb.Append(InlineStrCell("A1", SheetSetImporter.ExportNote, StyleNote));
                sb.Append("</row>");

                sb.Append("<row r=\"2\">");
                sb.Append(InlineStrCell("A2", "Order", StyleHeader));
                sb.Append(InlineStrCell("B2", "Kind", StyleHeader));
                sb.Append(InlineStrCell("C2", "Subset", StyleHeader));
                sb.Append(InlineStrCell("D2", "SheetNumber", StyleHeader));
                sb.Append(InlineStrCell("E2", "Title", StyleHeader));
                sb.Append(InlineStrCell("F2", "Revision", StyleHeader));
                sb.Append(InlineStrCell("G2", "Layout", StyleHeader));
                sb.Append(InlineStrCell("H2", "DWG", StyleHeader));
                sb.Append("</row>");

                for (int i = 0; i < rows.Count; i++)
                {
                    var e = rows[i];
                    int r = i + 3;
                    sb.Append($"<row r=\"{r}\">");
                    sb.Append($"<c r=\"A{r}\"><v>{e.Order}</v></c>");
                    sb.Append(InlineStrCell($"B{r}", e.IsSubset ? "Subset" : "Sheet"));
                    sb.Append(InlineStrCell($"C{r}",
                        e.IsSubset ? (e.Title ?? e.SubsetName ?? "") : (e.SubsetName ?? "")));
                    sb.Append(InlineStrCell($"D{r}", e.IsSubset ? "" : (e.SheetNumber ?? "")));
                    sb.Append(InlineStrCell($"E{r}", e.IsSubset ? "" : (e.Title ?? "")));
                    sb.Append(InlineStrCell($"F{r}", e.IsSubset ? "" : (e.Revision ?? "")));
                    sb.Append(InlineStrCell($"G{r}",
                        e.IsSubset ? (e.Title ?? e.SubsetName ?? "") : (e.LayoutName ?? "")));
                    sb.Append(InlineStrCell($"H{r}", e.DwgName ?? ""));
                    sb.Append("</row>");
                }

                sb.Append("</sheetData>");
                sb.Append("</worksheet>");
                AddEntry(archive, "xl/worksheets/sheet1.xml", sb.ToString());
            }
        }

        private static string BuildStylesXml()
        {
            // font 0: Calibri 11 default
            // font 1: Calibri 11 italic red (#FF0000) + vertical center
            // font 2: Calibri 11 bold
            return
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
                "<fonts count=\"3\">" +
                    "<font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/></font>" +
                    "<font><i/><sz val=\"11\"/><color rgb=\"FFFF0000\"/><name val=\"Calibri\"/><family val=\"2\"/></font>" +
                    "<font><b/><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/></font>" +
                "</fonts>" +
                "<fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>" +
                "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>" +
                "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
                "<cellXfs count=\"3\">" +
                    "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>" +
                    "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyAlignment=\"1\">" +
                        "<alignment vertical=\"center\" />" +
                    "</xf>" +
                    "<xf numFmtId=\"0\" fontId=\"2\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>" +
                "</cellXfs>" +
                "</styleSheet>";
        }

        private static string Csv(string value)
        {
            var v = value ?? "";
            if (v.IndexOfAny(new[] { ',', '"', '\r', '\n' }) >= 0)
                return "\"" + v.Replace("\"", "\"\"") + "\"";
            return v;
        }

        private static void AddEntry(ZipArchive archive, string fullName, string content)
        {
            var entry = archive.CreateEntry(fullName);
            using (var s = entry.Open())
            using (var w = new StreamWriter(s, new UTF8Encoding(false)))
                w.Write(content);
        }

        private static string InlineStrCell(string cellRef, string text, int styleIndex = 0)
        {
            var escaped = System.Security.SecurityElement.Escape(text ?? "") ?? "";
            string styleAttr = styleIndex > 0 ? $" s=\"{styleIndex}\"" : "";
            return $"<c r=\"{cellRef}\"{styleAttr} t=\"inlineStr\"><is><t xml:space=\"preserve\">{escaped}</t></is></c>";
        }
    }
}
