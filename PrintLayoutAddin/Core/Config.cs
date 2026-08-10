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
        public const string DefaultRevTableLayer = "TITLE_BLOCK_REV_TABLE";
        /// <summary>Fixed revision history slots (UI + write). Override via config.json.</summary>
        public const int DefaultRevTableDataRows = 4;

        public string AttributeTag { get; set; } = DefaultFrameNumberTag;
        public string DrawingNameTag { get; set; } = DefaultDrawingNameTag;
        public string VpLayer { get; set; } = "360D-Mview";
        public string XdataAppName { get; set; } = "PLADDIN_STT";
        public string TemplateLayout { get; set; } = "Layout1";
        /// <summary>Subfolder next to the DWG for .dst + PDF defaults.</summary>
        public string SheetSetFolderName { get; set; } = DefaultSheetSetFolderName;
        /// <summary>Paper-space layer of the revision history Table.</summary>
        public string RevTableLayer { get; set; } = DefaultRevTableLayer;
        /// <summary>Fixed number of revision data rows (header excluded).</summary>
        public int RevTableDataRows { get; set; } = DefaultRevTableDataRows;
        /// <summary>
        /// After PLAYOUT, silently Create/Update the default DST so Sheet Set
        /// title-block fields resolve (no more ####) without opening PLSHEETSET.
        /// </summary>
        public bool AutoSheetSetAfterLayout { get; set; } = true;

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
                cfg.RevTableLayer = ExtractString(json, "revTableLayer") ?? cfg.RevTableLayer;
                var rows = ExtractInt(json, "revTableDataRows");
                if (rows.HasValue && rows.Value > 0)
                    cfg.RevTableDataRows = rows.Value;
                var autoDst = ExtractBool(json, "autoSheetSetAfterLayout");
                if (autoDst.HasValue)
                    cfg.AutoSheetSetAfterLayout = autoDst.Value;
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

        private static int? ExtractInt(string json, string key)
        {
            var m = Regex.Match(json, "\"" + Regex.Escape(key) + "\"\\s*:\\s*(-?\\d+)");
            if (!m.Success) return null;
            if (int.TryParse(m.Groups[1].Value, out var n)) return n;
            return null;
        }

        private static bool? ExtractBool(string json, string key)
        {
            var m = Regex.Match(json, "\"" + Regex.Escape(key) + "\"\\s*:\\s*(true|false)",
                RegexOptions.IgnoreCase);
            if (!m.Success) return null;
            return string.Equals(m.Groups[1].Value, "true", StringComparison.OrdinalIgnoreCase);
        }
    }
}
