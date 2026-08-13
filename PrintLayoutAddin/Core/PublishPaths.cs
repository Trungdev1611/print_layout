using System;
using System.IO;

namespace PrintLayoutAddin.Core
{
    /// <summary>
    /// Shared default locations for Sheet Set (.dst) and PDF export.
    /// Both land in <c>{dwgDir}/sheetset_manager/</c> (name configurable).
    /// Path helpers do <b>not</b> create the folder — callers that write a file
    /// must call <see cref="EnsureFolder"/> (or create the parent of the file path).
    /// </summary>
    public static class PublishPaths
    {
        /// <summary>
        /// Returns the default publish folder path. Does not create it unless
        /// <paramref name="create"/> is true (prefer false for path-only / dialog open).
        /// </summary>
        public static string GetFolder(string dwgPath, bool create = false)
        {
            string dir = null;
            try
            {
                if (!string.IsNullOrWhiteSpace(dwgPath))
                    dir = Path.GetDirectoryName(dwgPath);
            }
            catch { }

            if (string.IsNullOrWhiteSpace(dir) || !Directory.Exists(dir))
                dir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            var folderName = Config.Instance.SheetSetFolderName;
            if (string.IsNullOrWhiteSpace(folderName))
                folderName = Config.DefaultSheetSetFolderName;

            var folder = Path.Combine(dir, folderName.Trim());
            if (create)
                EnsureFolder(folder);
            return folder;
        }

        /// <summary>Create <paramref name="folder"/> when missing (write sites only).</summary>
        public static void EnsureFolder(string folder)
        {
            if (string.IsNullOrWhiteSpace(folder)) return;
            try { Directory.CreateDirectory(folder); }
            catch { /* caller may surface IO errors later */ }
        }

        public static string DefaultDstPath(string dwgPath)
        {
            return Path.Combine(GetFolder(dwgPath, create: false), BaseName(dwgPath) + ".dst");
        }

        public static string DefaultPdfPath(string dwgPath)
        {
            return Path.Combine(GetFolder(dwgPath, create: false), BaseName(dwgPath) + "_layout.pdf");
        }

        /// <summary>
        /// Keep a remembered path only when it still targets this drawing's
        /// sheetset_manager folder; otherwise fall back to the fresh default.
        /// </summary>
        public static string ResolveRememberedPath(string dwgPath, string rememberedPath, string freshDefault)
        {
            if (string.IsNullOrWhiteSpace(rememberedPath)) return freshDefault;
            try
            {
                var rememberedDir = Path.GetDirectoryName(rememberedPath);
                var expectedDir = GetFolder(dwgPath, create: false);
                if (string.Equals(rememberedDir, expectedDir, StringComparison.OrdinalIgnoreCase))
                    return rememberedPath;
            }
            catch { }
            return freshDefault;
        }

        private static string BaseName(string dwgPath)
        {
            try
            {
                var name = Path.GetFileNameWithoutExtension(dwgPath ?? "");
                if (!string.IsNullOrWhiteSpace(name)) return name;
            }
            catch { }
            return "PrintLayout";
        }
    }
}
