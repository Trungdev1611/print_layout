using Autodesk.AutoCAD.Runtime;
using PrintLayoutAddin.Core;
using PrintLayoutAddin.UI;

[assembly: ExtensionApplication(typeof(PrintLayoutAddin.ExtensionApp))]

namespace PrintLayoutAddin
{
    public class ExtensionApp : IExtensionApplication
    {
        public void Initialize()
        {
            try { RibbonBuilder.Build(); } catch { }
            try { ShortcutManager.Install(); } catch { }
        }

        public void Terminate()
        {
            try { RibbonBuilder.Remove(); } catch { }
            try { ShortcutManager.Uninstall(); } catch { }
        }
    }
}
