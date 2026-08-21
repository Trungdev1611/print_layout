using System.Reflection;
using Autodesk.AutoCAD.Runtime;
using FlexiLicense;
using PrintLayoutAddin.Core;
using PrintLayoutAddin.UI;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: ExtensionApplication(typeof(PrintLayoutAddin.ExtensionApp))]

namespace PrintLayoutAddin
{
    public class ExtensionApp : IExtensionApplication
    {
        public void Initialize()
        {
            // Prompts for a key if needed (brand-new install, no key saved yet) and
            // warms the cache CheckCached() reads per-command below. Deliberately
            // does NOT call HandleResult() here — no ForceUpdate/blocking popup at
            // AutoCAD startup. Each command's own CheckCached()+HandleResult() call
            // gates enforcement at the moment it's actually invoked instead.
            var licInitResult = LicenseClient.VerifyWithPrompt("Print_layout", Assembly.GetExecutingAssembly());
            LicenseLog.Write("Initialize", licInitResult);

            try { RibbonBuilder.Build(); } catch { }
            try { ShortcutManager.Install(); } catch { }
            try { LayoutDstSyncWatcher.Start(); } catch { }
            try
            {
                var location = Assembly.GetExecutingAssembly().Location;
                var name = Assembly.GetExecutingAssembly().GetName();
                var dll = System.IO.Path.GetFileName(location);
                var version = name.Version?.ToString() ?? "?";

                AcadApp.DocumentManager?.MdiActiveDocument?.Editor?.WriteMessage(
                    "\nPrint Layout loaded: v" + version + "  (" + dll + ")");

                // Plain-text snapshot next to the DLL so the installed version can be
                // checked (support, scripts) without opening AutoCAD. Overwritten on
                // every load — this is "what's installed now", not a history log.
                var logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(location), "version.log");
                System.IO.File.WriteAllText(logPath,
                    "PrintLayoutAddin v" + version + "\r\n" +
                    "DLL: " + dll + "\r\n" +
                    "Last loaded: " + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "\r\n");
            }
            catch { }
        }

        public void Terminate()
        {
            try { LayoutDstSyncWatcher.Stop(); } catch { }
            try { RibbonBuilder.Remove(); } catch { }
            try { ShortcutManager.Uninstall(); } catch { }
        }
    }
}
