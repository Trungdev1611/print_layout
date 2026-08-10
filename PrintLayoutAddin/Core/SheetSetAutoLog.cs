using System;
using System.IO;
using Autodesk.AutoCAD.EditorInput;

namespace PrintLayoutAddin.Core
{
    /// <summary>
    /// Command-line + file log for silent sheet-set sync (PLAYOUT / layout-delete).
    /// File: <c>{sheetset_manager}/pl_sheetset_auto.log</c> next to the DWG.
    /// </summary>
    public static class SheetSetAutoLog
    {
        public const string FileName = "pl_sheetset_auto.log";

        public static void Write(Editor ed, string dwgPath, string message)
        {
            string msg = message ?? "";
            try { ed?.WriteMessage("\n[PLSHEETSET-AUTO] " + msg); }
            catch { }

            try
            {
                string folder = !string.IsNullOrWhiteSpace(dwgPath)
                    ? PublishPaths.GetFolder(dwgPath, create: true)
                    : Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                        Config.Instance.SheetSetFolderName ?? Config.DefaultSheetSetFolderName);
                Directory.CreateDirectory(folder);
                string line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + msg;
                File.AppendAllText(Path.Combine(folder, FileName), line + Environment.NewLine);
            }
            catch { }
        }
    }
}
