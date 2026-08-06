using System;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;

namespace PrintLayoutAddin.Core
{
    public class Config
    {
        // Central defaults for frame attributes. Deployments can override either
        // value in config.json without recompiling the add-in.
        public const string DefaultFrameNumberTag = "INNO-STT";
        public const string DefaultDrawingNameTag = "INNO_NAME_DRAWING";
        public const string DefaultSheetSetFolderName = "sheetset_manager";

        public string AttributeTag { get; set; } = DefaultFrameNumberTag;
        public string DrawingNameTag { get; set; } = DefaultDrawingNameTag;
        public string VpLayer { get; set; } = "360D-Mview";
        public string XdataAppName { get; set; } = "PLADDIN_STT";
        public string TemplateLayout { get; set; } = "Layout1";
        /// <summary>Subfolder next to the DWG for .dst + PDF defaults.</summary>
        public string SheetSetFolderName { get; set; } = DefaultSheetSetFolderName;

        private static Config _instance;

        public static Config Instance => _instance ?? (_instance = Load());

        private static Config Load()
        {
            var cfg = new Config();
            try
            {
                var dllDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                var path = Path.Combine(dllDir ?? "", "config.json");
                if (!File.Exists(path)) return cfg;

                var json = File.ReadAllText(path);
                cfg.AttributeTag = ExtractString(json, "attributeTag") ?? cfg.AttributeTag;
                cfg.DrawingNameTag = ExtractString(json, "drawingNameTag") ?? cfg.DrawingNameTag;
                cfg.VpLayer = ExtractString(json, "vpLayer") ?? cfg.VpLayer;
                cfg.XdataAppName = ExtractString(json, "xdataAppName") ?? cfg.XdataAppName;
                cfg.TemplateLayout = ExtractString(json, "templateLayout") ?? cfg.TemplateLayout;
                cfg.SheetSetFolderName = ExtractString(json, "sheetSetFolderName") ?? cfg.SheetSetFolderName;
            }
            catch
            {
                // fall back to defaults
            }
            return cfg;
        }

        private static string ExtractString(string json, string key)
        {
            var m = Regex.Match(json, "\"" + Regex.Escape(key) + "\"\\s*:\\s*\"([^\"]*)\"");
            return m.Success ? m.Groups[1].Value : null;
        }
    }
}
