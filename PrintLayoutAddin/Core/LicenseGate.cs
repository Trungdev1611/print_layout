using Autodesk.AutoCAD.EditorInput;
using PrintLayoutAddin.UI;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace PrintLayoutAddin.Core
{
    public static class LicenseGate
    {
        // Returns true if the addin is licensed and the calling command may proceed.
        // On invalid state, opens the LicenseDialog modally so the user can activate
        // without restarting AutoCAD. Re-checks after the dialog closes.
        public static bool Allow(Editor ed)
        {
            var info = LicenseManager.Current;
            if (info.IsValid)
            {
                if (info.DaysRemaining >= 0 && info.DaysRemaining <= LicenseManager.WarnDaysThreshold)
                    ed?.WriteMessage(
                        $"\n[License] Expires in {info.DaysRemaining} day(s) (on {info.Expiration:yyyy-MM-dd}). Contact your supplier to renew.");
                return true;
            }

            ed?.WriteMessage("\n[License] " + (info.Message ?? "Not valid.") + " Opening activation dialog...");
            try
            {
                using (var dlg = new LicenseDialog())
                    AcadApp.ShowModalDialog(dlg);
            }
            catch (System.Exception ex)
            {
                ed?.WriteMessage("\n[License] Could not open dialog: " + ex.Message);
                return false;
            }

            info = LicenseManager.Current;
            if (info.IsValid) return true;

            ed?.WriteMessage("\n[License] Command rejected. Run 'PLLICENSE' to activate.");
            return false;
        }
    }
}
