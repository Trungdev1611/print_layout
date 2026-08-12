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

        /// <summary>
        /// Pass <paramref name="dstPath"/> whenever the caller has no dwgPath: without it the
        /// line lands in the fallback log under Documents instead of the one next to the .dst,
        /// which is how a stack trace naming the real failing AcSm call went unnoticed.
        /// </summary>
        public static void WriteTagged(
            Editor ed, string dwgPath, string tag, string message, string dstPath = null)
        {
            string prefix = string.IsNullOrWhiteSpace(tag) ? "[PLSHEETSET]" : "[" + tag + "]";
            string msg = message ?? "";
            try { ed?.WriteMessage("\n" + prefix + " " + msg); }
            catch { }

            AppendFile(dwgPath, dstPath, prefix + " " + msg);
        }

        private static void WriteTaggedDst(
            Editor ed, string dwgPath, string tag, string dstPath, string message) =>
            WriteTagged(ed, dwgPath, tag, message, dstPath);

        /// <summary>Log a pipeline step with optional .dst path context (file existence, temp path, etc.).</summary>
        public static void WriteStep(Editor ed, string dwgPath, string dstPath, string step, string detail = null)
        {
            var sb = new StringBuilder(step ?? "");
            if (!string.IsNullOrWhiteSpace(dstPath))
                sb.Append(" | dst=").Append(dstPath);
            if (!string.IsNullOrWhiteSpace(detail))
                sb.Append(" | ").Append(detail);
            WriteTagged(ed, dwgPath, "PLSHEETSET-STEP", sb.ToString(), dstPath);
        }

        public static void WriteException(
            Editor ed, string dwgPath, string tag, string context, Exception ex, string dstPath = null)
        {
            if (ex == null)
            {
                WriteTaggedDst(ed, dwgPath, tag, dstPath, context ?? "exception: (null)");
                return;
            }

            WriteTaggedDst(ed, dwgPath, tag, dstPath, context ?? "exception");
            WriteTaggedDst(ed, dwgPath, tag, dstPath, "  Type: " + ex.GetType().FullName);
            WriteTaggedDst(ed, dwgPath, tag, dstPath, "  Message: " + (ex.Message ?? ""));

            if (ex is COMException com)
            {
                WriteTaggedDst(ed, dwgPath, tag, dstPath,
                    "  COM ErrorCode: 0x" + unchecked((uint)com.ErrorCode).ToString("X8")
                    + "  HResult: 0x" + unchecked((uint)com.HResult).ToString("X8"));
            }

            if (!string.IsNullOrWhiteSpace(dstPath))
            {
                WriteTaggedDst(ed, dwgPath, tag, dstPath,
                    "  dst exists=" + File.Exists(dstPath)
                    + "  dst$ exists=" + File.Exists(dstPath + "$"));
            }

            var inner = ex.InnerException;
            while (inner != null)
            {
                WriteTaggedDst(ed, dwgPath, tag, dstPath, "  Inner: " + inner.GetType().Name + ": " + inner.Message);
                if (inner is COMException innerCom)
                {
                    WriteTaggedDst(ed, dwgPath, tag, dstPath,
                        "    Inner COM: 0x" + unchecked((uint)innerCom.ErrorCode).ToString("X8"));
                }
                inner = inner.InnerException;
            }

            if (!string.IsNullOrWhiteSpace(ex.StackTrace))
            {
                var lines = ex.StackTrace.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                int max = Math.Min(lines.Length, 6);
                for (int i = 0; i < max; i++)
                    WriteTaggedDst(ed, dwgPath, tag, dstPath, "  at " + lines[i].Trim());
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
