using System.Reflection;
using Autodesk.AutoCAD.Runtime;
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
            try { RibbonBuilder.Build(); } catch { }
            try { ShortcutManager.Install(); } catch { }
            try { LayoutDstSyncWatcher.Start(); } catch { }
            try
            {
                var name = Assembly.GetExecutingAssembly().GetName();
                var dll = System.IO.Path.GetFileName(Assembly.GetExecutingAssembly().Location);
                AcadApp.DocumentManager?.MdiActiveDocument?.Editor?.WriteMessage(
                    "\nPrint Layout loaded: v" + (name.Version?.ToString() ?? "?")
                    + "  (" + dll + ")");
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
