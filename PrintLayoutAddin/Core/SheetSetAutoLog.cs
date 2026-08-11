using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Autodesk.AutoCAD.EditorInput;

namespace PrintLayoutAddin.Core
{
    /// <summary>
    /// Command-line + file log for sheet-set operations (auto after PLAYOUT, PLSHEETSET dialog, COM failures).
    /// File: <c>{sheetset_manager}/pl_sheetset_auto.log</c> next to the DWG (or next to the .dst).
    /// </summary>
    public static class SheetSetAutoLog
    {
        public const string FileName = "pl_sheetset_auto.log";

        public static void Write(Editor ed, string dwgPath, string message) =>
            WriteTagged(ed, dwgPath, null, message);

        public static void WriteTagged(Editor ed, string dwgPath, string tag, string message)
        {
            string prefix = string.IsNullOrWhiteSpace(tag) ? "[PLSHEETSET]" : "[" + tag + "]";
            string msg = message ?? "";
            try { ed?.WriteMessage("\n" + prefix + " " + msg); }
            catch { }

            AppendFile(dwgPath, null, prefix + " " + msg);
        }

        /// <summary>Log a pipeline step with optional .dst path context (file existence, temp path, etc.).</summary>
        public static void WriteStep(Editor ed, string dwgPath, string dstPath, string step, string detail = null)
        {
            var sb = new StringBuilder(step ?? "");
            if (!string.IsNullOrWhiteSpace(dstPath))
                sb.Append(" | dst=").Append(dstPath);
            if (!string.IsNullOrWhiteSpace(detail))
                sb.Append(" | ").Append(detail);
            WriteTagged(ed, dwgPath, "PLSHEETSET-STEP", sb.ToString());
        }

        public static void WriteException(
            Editor ed, string dwgPath, string tag, string context, Exception ex, string dstPath = null)
        {
            if (ex == null)
            {
                WriteTagged(ed, dwgPath, tag, context ?? "exception: (null)");
                return;
            }

            WriteTagged(ed, dwgPath, tag, context ?? "exception");
            WriteTagged(ed, dwgPath, tag, "  Type: " + ex.GetType().FullName);
            WriteTagged(ed, dwgPath, tag, "  Message: " + (ex.Message ?? ""));

            if (ex is COMException com)
            {
                WriteTagged(ed, dwgPath, tag,
                    "  COM ErrorCode: 0x" + unchecked((uint)com.ErrorCode).ToString("X8")
                    + "  HResult: 0x" + unchecked((uint)com.HResult).ToString("X8"));
            }

            if (!string.IsNullOrWhiteSpace(dstPath))
            {
                WriteTagged(ed, dwgPath, tag,
                    "  dst exists=" + File.Exists(dstPath)
                    + "  dst$ exists=" + File.Exists(dstPath + "$"));
            }

            var inner = ex.InnerException;
            while (inner != null)
            {
                WriteTagged(ed, dwgPath, tag, "  Inner: " + inner.GetType().Name + ": " + inner.Message);
                if (inner is COMException innerCom)
                {
                    WriteTagged(ed, dwgPath, tag,
                        "    Inner COM: 0x" + unchecked((uint)innerCom.ErrorCode).ToString("X8"));
                }
                inner = inner.InnerException;
            }

            if (!string.IsNullOrWhiteSpace(ex.StackTrace))
            {
                var lines = ex.StackTrace.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                int max = Math.Min(lines.Length, 6);
                for (int i = 0; i < max; i++)
                    WriteTagged(ed, dwgPath, tag, "  at " + lines[i].Trim());
            }
        }

        public static void WriteComFailure(string dwgPath, string dstPath, string operation, COMException ex) =>
            WriteException(null, dwgPath, "PLSHEETSET-COM", operation ?? "COM call failed", ex, dstPath);

        public static string GetLogFilePath(string dwgPath, string dstPath = null) =>
            Path.Combine(ResolveLogFolder(dwgPath, dstPath), FileName);

        private static void AppendFile(string dwgPath, string dstPath, string line)
        {
            try
            {
                string folder = ResolveLogFolder(dwgPath, dstPath);
                Directory.CreateDirectory(folder);
                string text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + (line ?? "") + Environment.NewLine;
                File.AppendAllText(Path.Combine(folder, FileName), text);
            }
            catch { }
        }

        private static string ResolveLogFolder(string dwgPath, string dstPath)
        {
            if (!string.IsNullOrWhiteSpace(dwgPath))
            {
                try
                {
                    return PublishPaths.GetFolder(dwgPath, create: true);
                }
                catch { }
            }

            if (!string.IsNullOrWhiteSpace(dstPath))
            {
                try
                {
                    var dir = Path.GetDirectoryName(dstPath);
                    if (!string.IsNullOrWhiteSpace(dir))
                        return dir;
                }
                catch { }
            }

            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                Config.Instance.SheetSetFolderName ?? Config.DefaultSheetSetFolderName);
        }
    }
}
