using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

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

    /// <summary>
    /// Creates native AutoCAD .dst files through the in-process Sheet Set COM API.
    /// Late binding avoids a version-specific AcSmComponents interop reference,
    /// which would otherwise prevent one add-in build from spanning AutoCAD releases.
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
            if (entries == null || entries.Count == 0)
                throw new InvalidOperationException("No sheets are selected.");

            var dir = Path.GetDirectoryName(dstPath);
            if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);

            object manager = null;
            object database = null;
            bool locked = false;
            try
            {
                manager = CreateComObject("AcSmSheetSetMgr");
                database = Invoke(manager, "CreateDatabase", dstPath, "", true);
                if (database == null)
                    throw new InvalidOperationException("AutoCAD did not create the sheet set database.");

                Invoke(database, "LockDb", database);
                locked = true;

                var sheetSet = Invoke(database, "GetSheetSet");
                Invoke(sheetSet, "SetName",
                    string.IsNullOrWhiteSpace(sheetSetName)
                        ? Path.GetFileNameWithoutExtension(dstPath)
                        : sheetSetName.Trim());

                foreach (var entry in entries)
                {
                    if (entry == null || !entry.Include || entry.Layout == null) continue;
                    if (string.IsNullOrWhiteSpace(entry.DwgPath) || !File.Exists(entry.DwgPath))
                        throw new FileNotFoundException("Save the source DWG before creating a sheet set.", entry.DwgPath);

                    var layoutRef = CreateComObject("AcSmAcDbLayoutReference");
                    Invoke(layoutRef, "InitNew", sheetSet);
                    Invoke(layoutRef, "SetFileName", entry.DwgPath);
                    Invoke(layoutRef, "SetName", entry.Layout.Name);

                    var sheet = Invoke(sheetSet, "ImportSheet", layoutRef);
                    if (sheet == null)
                        throw new InvalidOperationException($"Could not import layout '{entry.Layout.Name}'.");
                    Invoke(sheet, "SetNumber", entry.SheetNumber ?? "");
                    Invoke(sheet, "SetTitle",
                        string.IsNullOrWhiteSpace(entry.Title) ? entry.Layout.Name : entry.Title.Trim());
                    Invoke(sheetSet, "InsertComponent", sheet, null);

                    ReleaseCom(sheet);
                    ReleaseCom(layoutRef);
                }

                // Attach/refresh the sheet-set hint in the currently open DWG.
                // CurrentSheetNumber / CurrentSheetTitle fields rely on this
                // context to resolve the active layout to its sheet metadata.
                try { Invoke(database, "UpdateInMemoryDwgHints"); } catch { }

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
                    try { Invoke(database, "UnlockDb", database, true); } catch { }
                }
                if (manager != null && database != null)
                {
                    try { Invoke(manager, "Close", database); } catch { }
                }
                ReleaseCom(database);
                ReleaseCom(manager);
            }
        }

        private static object CreateComObject(string className)
        {
            var progIds = new List<string>();
            int major = GetAcadMajorVersion();
            if (major > 0)
                progIds.Add($"AcSmComponents.{className}.{major}");
            progIds.Add($"AcSmComponents.{className}");

            foreach (var progId in progIds)
            {
                try
                {
                    var type = Type.GetTypeFromProgID(progId, false);
                    if (type != null) return Activator.CreateInstance(type);
                }
                catch { }
            }

            throw new InvalidOperationException(
                $"AutoCAD Sheet Set component '{className}' is unavailable. "
                + "Run this command inside a supported full AutoCAD installation.");
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

        private static object Invoke(object target, string method, params object[] args)
        {
            if (target == null) throw new ArgumentNullException(nameof(target));
            return target.GetType().InvokeMember(
                method,
                System.Reflection.BindingFlags.InvokeMethod,
                null,
                target,
                args,
                CultureInfo.InvariantCulture);
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
    }
}
