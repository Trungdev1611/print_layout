using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace PrintLayoutAddin.Core
{
    public class ImportLayoutSheet
    {
        public string LayoutName { get; set; }
        /// <summary>Title-block drawing name (INNO_NAME_DRAWING); may be empty.</summary>
        public string DrawingName { get; set; }
    }

    public class ImportDrawingFile
    {
        public string DwgPath { get; set; }
        public List<ImportLayoutSheet> Layouts { get; set; } = new List<ImportLayoutSheet>();
    }

    /// <summary>One folder in the import tree (only kept when it has drawings or non-empty children).</summary>
    public class ImportFolderNode
    {
        public string Name { get; set; }
        public string FullPath { get; set; }
        public List<ImportFolderNode> Children { get; } = new List<ImportFolderNode>();
        public List<ImportDrawingFile> Drawings { get; } = new List<ImportDrawingFile>();

        public bool HasContent =>
            Drawings.Count > 0 || Children.Any(c => c != null && c.HasContent);

        public int CountDrawings() =>
            Drawings.Count + Children.Sum(c => c.CountDrawings());

        public int CountLayouts() =>
            Drawings.Sum(d => d.Layouts?.Count ?? 0)
            + Children.Sum(c => c.CountLayouts());

        public int CountSubsets() =>
            1 + Children.Sum(c => c.CountSubsets());

        public IEnumerable<(string DwgPath, ImportLayoutSheet Layout)> EnumerateLayouts()
        {
            foreach (var d in Drawings)
            {
                if (d == null || string.IsNullOrWhiteSpace(d.DwgPath) || d.Layouts == null)
                    continue;
                foreach (var layout in d.Layouts)
                {
                    if (layout == null || string.IsNullOrWhiteSpace(layout.LayoutName)) continue;
                    yield return (d.DwgPath, layout);
                }
            }
            foreach (var child in Children)
            {
                if (child == null) continue;
                foreach (var pair in child.EnumerateLayouts())
                    yield return pair;
            }
        }
    }

    public class FolderImportScanResult
    {
        public ImportFolderNode Root { get; set; }
        public int SkippedFiles { get; set; }
        public int FailedFiles { get; set; }
        public string Message { get; set; }
    }

    /// <summary>
    /// Recursive folder → subset/sheet tree for Sheet Set import.
    /// Only .dwg/.dwt; paper-space layouts except template; empty folders pruned.
    /// </summary>
    public static class SheetSetFolderImport
    {
        private static readonly string[] DrawingExtensions = { ".dwg", ".dwt" };

        public static FolderImportScanResult Scan(string rootFolder)
        {
            var result = new FolderImportScanResult();
            if (string.IsNullOrWhiteSpace(rootFolder) || !Directory.Exists(rootFolder))
            {
                result.Message = "Folder not found.";
                return result;
            }

            try
            {
                result.Root = BuildNode(rootFolder, result);
                if (result.Root == null || !result.Root.HasContent)
                {
                    result.Root = null;
                    if (string.IsNullOrWhiteSpace(result.Message))
                        result.Message = "No .dwg/.dwt with titled layouts (DrawingName) found under this folder.";
                    return result;
                }

                result.Message =
                    $"Found {result.Root.CountSubsets()} subset folder(s), "
                    + $"{result.Root.CountDrawings()} drawing(s), "
                    + $"{result.Root.CountLayouts()} layout(s)"
                    + (result.SkippedFiles > 0 ? $"; skipped {result.SkippedFiles} other file(s)" : "")
                    + (result.FailedFiles > 0 ? $"; {result.FailedFiles} file(s) failed to read" : "")
                    + ".";
                return result;
            }
            catch (Exception ex)
            {
                result.Root = null;
                result.Message = ex.Message;
                return result;
            }
        }

        public static string SheetKey(string dwgPath, string layoutName)
        {
            return NormalizePath(dwgPath) + "\n" + (layoutName ?? "").Trim().ToUpperInvariant();
        }

        public static string NormalizePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return "";
            try
            {
                return Path.GetFullPath(path).TrimEnd('\\', '/').ToUpperInvariant();
            }
            catch
            {
                return path.Trim().ToUpperInvariant();
            }
        }

        /// <summary>
        /// Sheet Set title (SSM shows "Number - Title"). Number = layout name;
        /// Title = drawing-name attribute (required for import).
        /// </summary>
        public static string ResolveSheetTitle(string layoutName, string drawingName, string dwgPath)
        {
            return (drawingName ?? "").Trim();
        }

        public static bool IsTemplateLayout(string layoutName)
        {
            var template = Config.Instance?.TemplateLayout;
            if (string.IsNullOrWhiteSpace(template))
                template = "Layout1";
            return string.Equals(layoutName?.Trim(), template.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Importable layouts: not template, and has a non-empty DrawingName
        /// (title-block attribute / STT map). No DrawingName ⇒ skip (no title block).
        /// </summary>
        public static List<ImportLayoutSheet> ReadImportLayouts(string dwgPath)
        {
            if (string.IsNullOrWhiteSpace(dwgPath) || !File.Exists(dwgPath))
                return new List<ImportLayoutSheet>();

            // If this DWG is open in any AutoCAD tab, read that document's DB
            // (not only the active tab — avoids eFileSharingViolation and wrong-file fallback).
            var openDb = TryFindOpenDocumentDatabase(dwgPath);
            if (openDb != null)
                return ReadImportLayouts(openDb);

            Database db = null;
            try
            {
                db = new Database(false, true);
                db.ReadDwgFile(dwgPath, FileOpenMode.OpenForReadAndReadShare, true, "");
                db.CloseInput(true);
                return ReadImportLayouts(db);
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
                when (ex.ErrorStatus == Autodesk.AutoCAD.Runtime.ErrorStatus.FileSharingViolation)
            {
                throw new InvalidOperationException(
                    "File is locked by another application and could not be read:\n"
                    + dwgPath
                    + "\n\nClose the file in the other app, or open it in this AutoCAD session "
                    + "so layouts can be read from memory.",
                    ex);
            }
            finally
            {
                try { db?.Dispose(); } catch { }
            }
        }

        /// <summary>
        /// Find an open document whose path matches <paramref name="dwgPath"/>
        /// (any tab, not only MdiActiveDocument).
        /// </summary>
        private static Database TryFindOpenDocumentDatabase(string dwgPath)
        {
            try
            {
                var docs = AcadApp.DocumentManager;
                if (docs == null) return null;
                foreach (Document doc in docs)
                {
                    if (doc?.Database == null) continue;
                    string name;
                    try { name = doc.Name; }
                    catch { continue; }
                    if (PathsEqual(name, dwgPath))
                        return doc.Database;
                }
            }
            catch { }
            return null;
        }

        /// <summary>
        /// Same rules as <see cref="ReadImportLayouts(string)"/> using an already-open Database
        /// (avoids eFileSharingViolation on drawings open in the editor).
        /// </summary>
        public static List<ImportLayoutSheet> ReadImportLayouts(Database db)
        {
            var list = new List<ImportLayoutSheet>();
            if (db == null) return list;

            Dictionary<string, string> byStt = null;
            try { byStt = FrameScanner.CollectDrawingNamesByStt(db); }
            catch { byStt = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase); }

            var printable = LayoutPlotter.GetPrintableLayouts(db);
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var layout in printable)
                {
                    if (layout == null || string.IsNullOrWhiteSpace(layout.Name)) continue;
                    if (IsTemplateLayout(layout.Name)) continue;

                    string drawingName = ReadDrawingNameOnLayout(db, tr, layout.Name);
                    if (string.IsNullOrWhiteSpace(drawingName)
                        && byStt != null
                        && byStt.TryGetValue(layout.Name.Trim(), out var fromStt))
                    {
                        drawingName = fromStt;
                    }

                    if (string.IsNullOrWhiteSpace(drawingName))
                        continue;

                    list.Add(new ImportLayoutSheet
                    {
                        LayoutName = layout.Name,
                        DrawingName = drawingName.Trim(),
                    });
                }
                tr.Commit();
            }

            return list;
        }

        /// <summary>
        /// Layout names that qualify for Sheet Set UI / DST (same rules as import).
        /// </summary>
        public static HashSet<string> GetImportableLayoutNames(string dwgPath)
        {
            return new HashSet<string>(
                ReadImportLayouts(dwgPath)
                    .Where(l => l != null && !string.IsNullOrWhiteSpace(l.LayoutName))
                    .Select(l => l.LayoutName),
                StringComparer.OrdinalIgnoreCase);
        }

        public static HashSet<string> GetImportableLayoutNames(Database db)
        {
            return new HashSet<string>(
                ReadImportLayouts(db)
                    .Where(l => l != null && !string.IsNullOrWhiteSpace(l.LayoutName))
                    .Select(l => l.LayoutName),
                StringComparer.OrdinalIgnoreCase);
        }

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

        /// <summary>Scan selected .dwg/.dwt files → drawings with paper-space layouts.</summary>
        public static FolderImportScanResult ScanDrawingFiles(IEnumerable<string> filePaths)
        {
            var result = new FolderImportScanResult
            {
                Root = new ImportFolderNode { Name = "(files)", FullPath = "" },
            };

            foreach (var path in (filePaths ?? Enumerable.Empty<string>())
                         .Where(p => !string.IsNullOrWhiteSpace(p))
                         .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var ext = Path.GetExtension(path) ?? "";
                if (!DrawingExtensions.Any(e => e.Equals(ext, StringComparison.OrdinalIgnoreCase)))
                {
                    result.SkippedFiles++;
                    continue;
                }
                if (!File.Exists(path))
                {
                    result.FailedFiles++;
                    continue;
                }

                try
                {
                    var layouts = ReadImportLayouts(path);
                    if (layouts.Count == 0)
                    {
                        result.SkippedFiles++;
                        continue;
                    }

                    result.Root.Drawings.Add(new ImportDrawingFile
                    {
                        DwgPath = Path.GetFullPath(path),
                        Layouts = layouts,
                    });
                }
                catch (Exception ex)
                {
                    result.FailedFiles++;
                    if (string.IsNullOrWhiteSpace(result.Message))
                        result.Message = ex.Message;
                }
            }

            if (!result.Root.HasContent)
            {
                result.Root = null;
                if (string.IsNullOrWhiteSpace(result.Message))
                    result.Message = "No layouts with DrawingName found in the selected drawing(s).";
                return result;
            }

            result.Message =
                $"Found {result.Root.CountDrawings()} drawing(s), {result.Root.CountLayouts()} layout(s)"
                + (result.SkippedFiles > 0 ? $"; skipped {result.SkippedFiles}" : "")
                + (result.FailedFiles > 0 ? $"; {result.FailedFiles} failed" : "")
                + ".";
            return result;
        }

        private static string ReadDrawingNameOnLayout(Database db, Transaction tr, string layoutName)
        {
            if (db == null || tr == null || string.IsNullOrWhiteSpace(layoutName))
                return "";

            try
            {
                var dict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                if (!dict.Contains(layoutName)) return "";
                var layout = (Layout)tr.GetObject(dict.GetAt(layoutName), OpenMode.ForRead);
                var paperBtr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                var cfg = Config.Instance;

                foreach (ObjectId id in paperBtr)
                {
                    if (id.IsNull || id.IsErased) continue;
                    var br = tr.GetObject(id, OpenMode.ForRead, false) as BlockReference;
                    if (br == null) continue;
                    var name = FrameScanner.ReadDrawingName(br, cfg);
                    if (!string.IsNullOrWhiteSpace(name))
                        return name.Trim();
                }
            }
            catch { }

            return "";
        }

        private static ImportFolderNode BuildNode(string folderPath, FolderImportScanResult stats)
        {
            var node = new ImportFolderNode
            {
                Name = Path.GetFileName(folderPath.TrimEnd('\\', '/')),
                FullPath = folderPath,
            };
            if (string.IsNullOrWhiteSpace(node.Name))
                node.Name = folderPath;

            try
            {
                foreach (var file in Directory.EnumerateFiles(folderPath))
                {
                    var ext = Path.GetExtension(file) ?? "";
                    if (!DrawingExtensions.Any(e => e.Equals(ext, StringComparison.OrdinalIgnoreCase)))
                    {
                        stats.SkippedFiles++;
                        continue;
                    }

                    try
                    {
                        var layouts = ReadImportLayouts(file);
                        if (layouts.Count == 0)
                        {
                            stats.SkippedFiles++;
                            continue;
                        }

                        node.Drawings.Add(new ImportDrawingFile
                        {
                            DwgPath = Path.GetFullPath(file),
                            Layouts = layouts,
                        });
                    }
                    catch (Exception ex)
                    {
                        stats.FailedFiles++;
                        if (string.IsNullOrWhiteSpace(stats.Message))
                            stats.Message = ex.Message;
                    }
                }

                foreach (var sub in Directory.EnumerateDirectories(folderPath)
                             .OrderBy(d => d, StringComparer.OrdinalIgnoreCase))
                {
                    var name = Path.GetFileName(sub);
                    if (string.IsNullOrWhiteSpace(name) || name.StartsWith(".", StringComparison.Ordinal))
                        continue;
                    if (name.Equals("sheetset_manager", StringComparison.OrdinalIgnoreCase))
                        continue;

                    var child = BuildNode(sub, stats);
                    if (child != null && child.HasContent)
                        node.Children.Add(child);
                }
            }
            catch
            {
                // Unreadable folder → treat as empty.
            }

            return node.HasContent ? node : null;
        }
    }
}
