using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

#if ACSM_INTEROP
using Autodesk.AutoCAD.Interop;
#endif

namespace PrintLayoutAddin.Core
{
    public class SheetSetEntry
    {
        public bool Include { get; set; } = true;
        public int Order { get; set; }
        public string SheetNumber { get; set; }
        public string Title { get; set; }
        public string DwgPath { get; set; }
        public PrintableLayout Layout { get; set; }

        public string LayoutName => Layout?.Name ?? "";
        public string DwgName => Path.GetFileName(DwgPath ?? "");
    }

    /// <summary>One sheet row read back from an existing .dst.</summary>
    public class SheetSetSheetInfo
    {
        public string LayoutName { get; set; }
        public string Number { get; set; }
        public string Title { get; set; }
    }

    public class SheetSetReadResult
    {
        public string SheetSetName { get; set; }
        public List<SheetSetSheetInfo> Sheets { get; } = new List<SheetSetSheetInfo>();
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
#if !ACSM_INTEROP
            throw new InvalidOperationException(
                "Sheet Set support was not compiled into this build (AcSmComponents.Interop.dll missing). "
                + "Rebuild with AutoCAD installed, or pass /p:AcSmInteropPath=... to the build.");
#else
            if (string.IsNullOrWhiteSpace(dstPath))
                throw new ArgumentException("A .dst output path is required.", nameof(dstPath));
            if (!dstPath.EndsWith(".dst", StringComparison.OrdinalIgnoreCase))
                dstPath += ".dst";
            if (entries == null || entries.Count == 0)
                throw new InvalidOperationException("No sheets are selected.");

            var dir = Path.GetDirectoryName(dstPath);
            if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);

            // Prefer the versioned ProgID so the RCW matches the running AutoCAD,
            // then cast to the typed vtable interfaces from AcSmComponents.Interop.
            IAcSmSheetSetMgr manager = null;
            IAcSmDatabase database = null;
            bool locked = false;
            try
            {
                manager = (IAcSmSheetSetMgr)CreateComObject("AcSmSheetSetMgr");
                database = manager.CreateDatabase(dstPath, "", true);
                if (database == null)
                    throw new InvalidOperationException("AutoCAD did not create the sheet set database.");

                database.LockDb(database);
                locked = true;

                IAcSmSheetSet sheetSet = database.GetSheetSet();
                sheetSet.SetName(
                    string.IsNullOrWhiteSpace(sheetSetName)
                        ? Path.GetFileNameWithoutExtension(dstPath)
                        : sheetSetName.Trim());

                foreach (var entry in entries)
                {
                    if (entry == null || !entry.Include || entry.Layout == null) continue;
                    if (string.IsNullOrWhiteSpace(entry.DwgPath) || !File.Exists(entry.DwgPath))
                        throw new FileNotFoundException(
                            "Save the source DWG before creating a sheet set.", entry.DwgPath);

                    var layoutRef = (IAcSmAcDbLayoutReference)CreateComObject("AcSmAcDbLayoutReference");
                    layoutRef.InitNew(sheetSet);
                    layoutRef.SetFileName(entry.DwgPath);
                    layoutRef.SetName(entry.Layout.Name);

                    // ImportSheet expects the coclass/RCW type from the interop assembly.
                    IAcSmSheet sheet = sheetSet.ImportSheet((AcSmAcDbLayoutReference)layoutRef);
                    if (sheet == null)
                        throw new InvalidOperationException(
                            $"Could not import layout '{entry.Layout.Name}'.");

                    sheet.SetNumber(entry.SheetNumber ?? "");
                    sheet.SetTitle(
                        string.IsNullOrWhiteSpace(entry.Title)
                            ? entry.Layout.Name
                            : entry.Title.Trim());
                    sheetSet.InsertComponent(sheet, null);

                    ReleaseCom(sheet);
                    ReleaseCom(layoutRef);
                }

                // Attach/refresh the sheet-set hint in the currently open DWG.
                // CurrentSheetNumber / CurrentSheetTitle fields rely on this
                // context to resolve the active layout to its sheet metadata.
                try { database.UpdateInMemoryDwgHints(); } catch { }
                try { sheetSet.UpdateInMemoryDwgHints(); } catch { }

                ReleaseCom(sheetSet);
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException(
                    "AutoCAD Sheet Set API failed. Close the DST in Sheet Set Manager if it is open, "
                    + "then try again. " + ex.Message, ex);
            }
            finally
            {
                if (database != null && locked)
                {
                    try { database.UnlockDb(database, true); } catch { }
                }
                if (manager != null && database != null)
                {
                    try { manager.Close((AcSmDatabase)database); } catch { }
                }
                ReleaseCom(database);
                ReleaseCom(manager);
            }
#endif
        }

        /// <summary>
        /// Reads sheet number/title from an existing DST, keyed by layout name.
        /// Returns null when the file is missing or cannot be opened.
        /// </summary>
        public static SheetSetReadResult TryRead(string dstPath)
        {
#if !ACSM_INTEROP
            return null;
#else
            if (string.IsNullOrWhiteSpace(dstPath) || !File.Exists(dstPath))
                return null;

            IAcSmSheetSetMgr manager = null;
            IAcSmDatabase database = null;
            bool locked = false;
            try
            {
                manager = (IAcSmSheetSetMgr)CreateComObject("AcSmSheetSetMgr");
                database = manager.OpenDatabase(dstPath, false);
                if (database == null) return null;

                database.LockDb(database);
                locked = true;

                var sheetSet = database.GetSheetSet();
                var result = new SheetSetReadResult
                {
                    SheetSetName = sheetSet?.GetName()
                };

                var enumerator = sheetSet?.GetSheetEnumerator();
                if (enumerator != null)
                {
                    object next;
                    while ((next = enumerator.Next()) != null)
                    {
                        var sheet = next as IAcSmSheet;
                        if (sheet == null)
                        {
                            ReleaseCom(next);
                            continue;
                        }

                        string layoutName = null;
                        try
                        {
                            var layout = sheet.GetLayout();
                            layoutName = layout?.GetName();
                            ReleaseCom(layout);
                        }
                        catch { }

                        if (string.IsNullOrWhiteSpace(layoutName))
                        {
                            ReleaseCom(sheet);
                            continue;
                        }

                        result.Sheets.Add(new SheetSetSheetInfo
                        {
                            LayoutName = layoutName,
                            Number = sheet.GetNumber() ?? "",
                            Title = sheet.GetTitle() ?? "",
                        });
                        ReleaseCom(sheet);
                    }
                    ReleaseCom(enumerator);
                }

                ReleaseCom(sheetSet);
                return result;
            }
            catch
            {
                return null;
            }
            finally
            {
                if (database != null && locked)
                {
                    try { database.UnlockDb(database, true); } catch { }
                }
                if (manager != null && database != null)
                {
                    try { manager.Close((AcSmDatabase)database); } catch { }
                }
                ReleaseCom(database);
                ReleaseCom(manager);
            }
#endif
        }

#if ACSM_INTEROP
        private static object CreateComObject(string className)
        {
            var progIds = new List<string>();
            int major = GetAcadMajorVersion();
            if (major > 0)
                progIds.Add($"AcSmComponents.{className}.{major}");
            progIds.Add($"AcSmComponents.{className}");

            Exception last = null;
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
                $"AutoCAD Sheet Set component '{className}' is unavailable. "
                + "Run this command inside a supported full AutoCAD installation."
                + (last != null ? " " + last.Message : ""));
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

        private static void ReleaseCom(object value)
        {
            if (value == null) return;
            try
            {
                if (Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
            }
            catch { }
        }
#endif
    }
}
