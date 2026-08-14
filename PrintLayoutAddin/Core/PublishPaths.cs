using System;
using System.IO;

namespace PrintLayoutAddin.Core
{
    /// <summary>
    /// Default PDF path is next to the DWG. Callers that write must create parent folders.
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
            string dir = null;
            try
            {
                if (!string.IsNullOrWhiteSpace(dwgPath))
                    dir = Path.GetDirectoryName(dwgPath);
            }
            catch { }
            if (string.IsNullOrWhiteSpace(dir) || !Directory.Exists(dir))
                dir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            return Path.Combine(dir, BaseName(dwgPath) + "_layout.pdf");
        }

        /// <summary>
        /// Unused by PDF now; kept for any caller that still keys off the configurable subfolder.
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
