using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.ApplicationServices;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

#if ACSM_INTEROP
using Autodesk.AutoCAD.Interop;
#endif

namespace PrintLayoutAddin.Core
{
    /// <summary>
    /// Opens the Sheet Set Manager UI and tries to surface a .dst for the user
    /// so they can continue (subsets, publish, etc.) in AutoCAD's SSM.
    /// </summary>
    public static class SheetSetLauncher
    {
#if ACSM_INTEROP
        // Keep one DST open in the Sheet Set Manager session so the palette lists it.
        private static IAcSmSheetSetMgr _uiManager;
        private static IAcSmDatabase _uiDatabase;
        private static string _uiPath;
#endif

        /// <summary>
        /// Shows SSM and attempts to open <paramref name="dstPath"/> in it.
        /// Returns a short status message for the command line / dialog.
        /// </summary>
        public static string OpenForUser(string dstPath)
        {
            if (string.IsNullOrWhiteSpace(dstPath) || !File.Exists(dstPath))
                return "DST file not found — open it manually in Sheet Set Manager.";

            string comNote = TryKeepOpenInManager(dstPath);
            ShowSheetSetPalette();
            TryShellOpen(dstPath);

            return "Sheet Set Manager opened for: " + dstPath
                + (string.IsNullOrEmpty(comNote) ? "" : " (" + comNote + ")");
        }

        /// <summary>
        /// Soft-close any open handle on this DST (our UI hold + AcSm FindOpenDatabase),
        /// then open it again so SSM shows the updated tree. Does not delete the file.
        /// </summary>
        public static string ReloadForUser(string dstPath)
        {
            if (string.IsNullOrWhiteSpace(dstPath) || !File.Exists(dstPath))
                return "DST file not found — open it manually in Sheet Set Manager.";

            ReleaseUiOpen();
            SoftCloseOpenDatabase(dstPath);
            return OpenForUser(dstPath);
        }

        /// <summary>
        /// Close this DST if AcSm has it open — does not delete the file.
        /// </summary>
        public static void SoftCloseOpenDatabase(string dstPath)
        {
#if ACSM_INTEROP
            if (string.IsNullOrWhiteSpace(dstPath)) return;
            IAcSmSheetSetMgr manager = null;
            try
            {
                manager = (IAcSmSheetSetMgr)CreateMgr();
                try
                {
                    var open = manager.FindOpenDatabase(dstPath);
                    if (open != null)
                    {
                        try { manager.Close((AcSmDatabase)open); } catch { }
                        ReleaseCom(open);
                    }
                }
                catch { }
            }
            catch { }
            finally
            {
                ReleaseCom(manager);
            }
#endif
        }

        /// <summary>
        /// Close any DST we held open for the UI so Create/Update can rewrite the file.
        /// </summary>
        public static void ReleaseUiOpen()
        {
#if ACSM_INTEROP
            try
            {
                if (_uiManager != null && _uiDatabase != null)
                {
                    try { _uiManager.Close((AcSmDatabase)_uiDatabase); } catch { }
                }
            }
            finally
            {
                ReleaseCom(_uiDatabase);
                ReleaseCom(_uiManager);
                _uiDatabase = null;
                _uiManager = null;
                _uiPath = null;
            }
#endif
        }

        // ForceCloseForRewrite / TryDeleteDstFiles used to live here. They called
        // manager.CloseAll() — closing every sheet set AcSm had open in the session,
        // including ones the SSM palette owned — and then File.Delete'd the user's .dst so a
        // temp file could be moved over it. SheetSetService now edits an existing .dst in
        // place through AcSm, so nothing needs to force a close or delete the file.

        private static string TryKeepOpenInManager(string dstPath)
        {
#if !ACSM_INTEROP
            return "COM interop unavailable";
#else
            try
            {
                ReleaseUiOpen();

                var manager = (IAcSmSheetSetMgr)CreateMgr();
                // false = do not fail if already open (typical OpenDatabase signature).
                var database = manager.OpenDatabase(dstPath, false);
                if (database == null)
                    return "OpenDatabase returned null";

                _uiManager = manager;
                _uiDatabase = database;
                _uiPath = dstPath;
                return "DST registered with Sheet Set Manager";
            }
            catch (Exception ex)
            {
                ReleaseUiOpen();
                return "could not register DST in SSM: " + ex.Message;
            }
#endif
        }

        private static void ShowSheetSetPalette()
        {
            try
            {
                var doc = AcadApp.DocumentManager.MdiActiveDocument;
                if (doc == null) return;
                // Show the Sheet Set Manager palette; runs after the modal dialog yields.
                doc.SendStringToExecute("_.SHEETSET ", true, false, false);
            }
            catch { }
        }

        private static void TryShellOpen(string dstPath)
        {
            try
            {
                Process.Start(new ProcessStartInfo(dstPath) { UseShellExecute = true });
            }
            catch { }
        }

#if ACSM_INTEROP
        private static object CreateMgr()
        {
            int major = 0;
            try
            {
                var raw = Convert.ToString(AcadApp.GetSystemVariable("ACADVER"));
                if (!string.IsNullOrWhiteSpace(raw))
                {
                    var dot = raw.IndexOf('.');
                    var first = dot >= 0 ? raw.Substring(0, dot) : raw;
                    int.TryParse(first, out major);
                }
            }
            catch { }

            string[] progIds = major > 0
                ? new[] { $"AcSmComponents.AcSmSheetSetMgr.{major}", "AcSmComponents.AcSmSheetSetMgr" }
                : new[] { "AcSmComponents.AcSmSheetSetMgr" };

            foreach (var progId in progIds)
            {
                try
                {
                    var type = Type.GetTypeFromProgID(progId, false);
                    if (type == null) continue;
                    var instance = Activator.CreateInstance(type);
                    if (instance != null) return instance;
                }
                catch { }
            }

            throw new InvalidOperationException("AcSmSheetSetMgr is unavailable.");
        }

        /// <summary>
        /// Drops only the reference we took. <c>FinalReleaseComObject</c> here forced the
        /// shared AcSmSheetSetMgr singleton's RCW to zero while AutoCAD's own Sheet Set
        /// Manager still used it — see the note on SheetSetService.ReleaseCom.
        /// </summary>
        private static void ReleaseCom(object value)
        {
            if (value == null) return;
            try
            {
                if (Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
            }
            catch { }
        }
#endif
    }
}
