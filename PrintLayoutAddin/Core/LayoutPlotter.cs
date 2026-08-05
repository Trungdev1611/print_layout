using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Publishing;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace PrintLayoutAddin.Core
{
    public class PrintableLayout
    {
        public string Name { get; set; }
        public ObjectId LayoutId { get; set; }
        public int TabOrder { get; set; }

        public override string ToString() => Name;
    }

    public class PlotMediaOption
    {
        public string CanonicalName { get; set; }
        public string DisplayName { get; set; }

        public override string ToString() => DisplayName ?? CanonicalName;
    }

    public enum PlotOrientationChoice
    {
        LayoutSetting,
        Portrait,
        Landscape
    }

    public enum PdfOutputMode
    {
        Combined,
        SeparatePerLayout
    }

    public enum PlotAreaChoice
    {
        Layout,
        Extents
    }

    public class PrintJobOptions
    {
        public string DeviceName { get; set; }
        public string MediaName { get; set; }
        public string StyleSheet { get; set; }
        public int Copies { get; set; } = 1;
        public bool PlotToFile { get; set; } = true;
        public string OutputPath { get; set; }
        public PdfOutputMode PdfOutputMode { get; set; } = PdfOutputMode.Combined;
        public PlotAreaChoice PlotArea { get; set; } = PlotAreaChoice.Layout;
        public bool ScaleToFit { get; set; }
        public PlotOrientationChoice Orientation { get; set; } = PlotOrientationChoice.LayoutSetting;
        public bool OpenAfterExport { get; set; }
        public List<PrintableLayout> Layouts { get; set; } = new List<PrintableLayout>();
    }

    public static class LayoutPlotter
    {
        public static List<PrintableLayout> GetPrintableLayouts(Database db)
        {
            var layouts = new List<PrintableLayout>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var dict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                foreach (DBDictionaryEntry entry in dict)
                {
                    var layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                    if (layout == null || layout.ModelType) continue;
                    layouts.Add(new PrintableLayout
                    {
                        Name = layout.LayoutName,
                        LayoutId = layout.ObjectId,
                        TabOrder = layout.TabOrder
                    });
                }
                tr.Commit();
            }
            return layouts.OrderBy(l => l.TabOrder).ThenBy(l => l.Name).ToList();
        }

        public static List<string> GetPlotDevices()
        {
            try
            {
                return ToList(PlotSettingsValidator.Current.GetPlotDeviceList())
                    .OrderBy(PreferredDeviceRank)
                    .ThenBy(s => s, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
            catch
            {
                // A broken/offline plotter driver can throw while enumerating; don't take AutoCAD down with it.
                return new List<string>();
            }
        }

        public static List<PlotMediaOption> GetMediaOptions(string deviceName)
        {
            var result = new List<PlotMediaOption>();
            if (string.IsNullOrWhiteSpace(deviceName)) return result;

            try
            {
                var psv = PlotSettingsValidator.Current;
                using (var ps = new PlotSettings(false))
                {
                    psv.SetPlotConfigurationName(ps, deviceName, null);
                    psv.RefreshLists(ps);
                    foreach (var canonical in ToList(psv.GetCanonicalMediaNameList(ps)))
                    {
                        string display = canonical;
                        try { display = psv.GetLocaleMediaName(ps, canonical); }
                        catch { }
                        result.Add(new PlotMediaOption
                        {
                            CanonicalName = canonical,
                            DisplayName = string.IsNullOrWhiteSpace(display) ? canonical : display
                        });
                    }
                }
            }
            catch
            {
                // SetPlotConfigurationName/RefreshLists can throw for a broken or offline device.
                // Return whatever was collected instead of letting the exception crash the dialog.
            }

            return result
                .OrderBy(m => MediaRank(m.DisplayName ?? m.CanonicalName))
                .ThenBy(m => m.DisplayName, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        public static List<string> GetPlotStyleSheets()
        {
            try
            {
                return ToList(PlotSettingsValidator.Current.GetPlotStyleSheetList())
                    .OrderBy(PreferredStyleRank)
                    .ThenBy(s => s, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
            catch
            {
                return new List<string>();
            }
        }

        public static void Plot(Document doc, PrintJobOptions options, Action<string> log = null)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (options.Layouts == null || options.Layouts.Count == 0)
                throw new InvalidOperationException("No layouts selected for plotting.");
            if (string.IsNullOrWhiteSpace(options.DeviceName))
                throw new InvalidOperationException("No plot device selected.");
            if (string.IsNullOrWhiteSpace(options.MediaName))
                throw new InvalidOperationException("No paper size selected.");
            if (options.Copies < 1)
                throw new InvalidOperationException("Copies must be at least 1.");
            if (options.PlotToFile && string.IsNullOrWhiteSpace(options.OutputPath))
                throw new InvalidOperationException("Output PDF path is required.");

            if (PlotFactory.ProcessPlotState != ProcessPlotState.NotPlotting)
                throw new InvalidOperationException("AutoCAD is already plotting. Wait for the current plot job to finish.");

            if (options.PlotToFile)
            {
                var dir = options.PdfOutputMode == PdfOutputMode.SeparatePerLayout
                    ? options.OutputPath
                    : Path.GetDirectoryName(options.OutputPath);
                if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);
            }

            var lm = LayoutManager.Current;
            string originalLayout = null;
            try { originalLayout = lm.CurrentLayout; } catch { }

            try
            {
                if (options.PlotToFile)
                {
                    PublishPdf(doc, options, lm, log);
                    return;
                }

                var plotInfos = BuildPlotInfos(doc.Database, options, lm);
                if (plotInfos.Count == 0) return;
                foreach (var pi in plotInfos)
                {
                    PlotDocument(doc, lm, new List<PreparedPlotInfo> { pi }, options.Copies, false, null, log);
                }
            }
            finally
            {
                if (!string.IsNullOrWhiteSpace(originalLayout))
                {
                    try { lm.CurrentLayout = originalLayout; } catch { }
                }
            }
        }

        private static void PublishPdf(
            Document doc,
            PrintJobOptions options,
            LayoutManager lm,
            Action<string> log)
        {
            if (string.IsNullOrWhiteSpace(doc.Name) || !File.Exists(doc.Name))
                throw new InvalidOperationException("Save the DWG before publishing layouts to PDF.");

            string setupName = "PLPRINT_PAGESETUP";
            try
            {
                CreateNamedPageSetup(doc.Database, setupName, options);

                object oldBackgroundPlot = null;
                try
                {
                    try { oldBackgroundPlot = AcadApp.GetSystemVariable("BACKGROUNDPLOT"); } catch { }
                    try { AcadApp.SetSystemVariable("BACKGROUNDPLOT", 0); } catch { }
                    // Publishing MUST run in the foreground. With background publishing on, PublishDsd
                    // returns immediately while the job keeps running; the temp .dsd is then deleted and
                    // the progress dialog disposed out from under it -> access violation / AutoCAD crash.
                    try
                    {
                        var bgp = AcadApp.GetSystemVariable("BACKGROUNDPLOT");
                        if (bgp != null && Convert.ToInt32(bgp) != 0)
                            log?.Invoke("Warning: BACKGROUNDPLOT could not be set to 0; publishing may be unstable.");
                    }
                    catch { }

                    if (options.PdfOutputMode == PdfOutputMode.SeparatePerLayout)
                    {
                        DeleteExistingSeparatePdfs(options);
                        log?.Invoke($"Publishing {options.Layouts.Count} layouts into separate PDFs...");
                        PublishDsd(doc, options.Layouts, setupName, options.OutputPath, SheetType.SinglePdf);
                    }
                    else
                    {
                        if (File.Exists(options.OutputPath)) File.Delete(options.OutputPath);
                        log?.Invoke($"Publishing {options.Layouts.Count} layouts into one PDF...");
                        PublishDsd(doc, options.Layouts, setupName, options.OutputPath, SheetType.MultiPdf);
                    }
                }
                finally
                {
                    if (oldBackgroundPlot != null)
                    {
                        try { AcadApp.SetSystemVariable("BACKGROUNDPLOT", oldBackgroundPlot); } catch { }
                    }
                }
            }
            finally
            {
                if (options.Layouts.Count > 0)
                {
                    try { lm.CurrentLayout = options.Layouts[0].Name; } catch { }
                }
            }
        }

        private static void DeleteExistingSeparatePdfs(PrintJobOptions options)
        {
            try
            {
                foreach (var layout in options.Layouts)
                {
                    var path = Path.Combine(options.OutputPath, MakeSafeFileName(layout.Name) + ".pdf");
                    if (File.Exists(path)) File.Delete(path);
                }
            }
            catch
            {
                // If a PDF is open/locked, let PublishDsd report the real failure path.
            }
        }

        private static void PublishDsd(
            Document doc,
            IList<PrintableLayout> layouts,
            string setupName,
            string outputPath,
            SheetType sheetType)
        {
            // PDF/file output ignores the number of copies (it always produces one file), so we publish
            // a single set of sheets. Copies are honored only when plotting to a physical printer.
            var entries = new DsdEntryCollection();
            foreach (var layout in layouts)
            {
                entries.Add(new DsdEntry
                {
                    DwgName = doc.Name,
                    Layout = layout.Name,
                    Title = layout.Name,
                    Nps = setupName,
                    NpsSourceDwg = doc.Name
                });
            }

            var dsdPath = Path.Combine(
                Path.GetTempPath(),
                "PLPRINT_" + Guid.NewGuid().ToString("N") + ".dsd");

            try
            {
                using (var dsd = new DsdData())
                {
                    dsd.SetDsdEntryCollection(entries);
                    dsd.SheetType = sheetType;
                    dsd.NoOfCopies = 1; // copies are handled by repeating entries above (PDF ignores NoOfCopies)
                    dsd.DestinationName = outputPath;
                    dsd.ProjectPath = Path.GetDirectoryName(outputPath);
                    dsd.SheetSetName = Path.GetFileNameWithoutExtension(outputPath);
                    dsd.PromptForDwfName = false;
                    dsd.IsHomogeneous = true;
                    dsd.LogFilePath = Path.ChangeExtension(outputPath, ".log");
                    dsd.WriteDsd(dsdPath);
                }

                using (var progress = new PlotProgressDialog(false, entries.Count, true))
                {
                    progress.set_PlotMsgString(PlotMessageIndex.DialogTitle, "Publish PDF");
                    progress.set_PlotMsgString(PlotMessageIndex.CancelJobButtonMessage, "Cancel job");
                    progress.set_PlotMsgString(PlotMessageIndex.CancelSheetButtonMessage, "Cancel sheet");
                    progress.set_PlotMsgString(PlotMessageIndex.SheetSetProgressCaption, "Sheet set progress");
                    progress.set_PlotMsgString(PlotMessageIndex.SheetProgressCaption, "Sheet progress");
                    progress.IsVisible = true;
                    AcadApp.Publisher.PublishDsd(dsdPath, progress);
                }
            }
            finally
            {
                try { if (File.Exists(dsdPath)) File.Delete(dsdPath); } catch { }
            }
        }

        private static void CreateNamedPageSetup(Database db, string setupName, PrintJobOptions options)
        {
            var psv = PlotSettingsValidator.Current;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var dict = (DBDictionary)tr.GetObject(db.PlotSettingsDictionaryId, OpenMode.ForWrite);
                if (dict.Contains(setupName))
                {
                    var old = tr.GetObject(dict.GetAt(setupName), OpenMode.ForWrite) as PlotSettings;
                    old?.Erase();
                }

                var ps = new PlotSettings(false)
                {
                    PlotSettingsName = setupName
                };
                psv.SetPlotConfigurationName(ps, options.DeviceName, options.MediaName);
                psv.RefreshLists(ps);
                psv.SetPlotType(ps, options.PlotArea == PlotAreaChoice.Extents
                    ? Autodesk.AutoCAD.DatabaseServices.PlotType.Extents
                    : Autodesk.AutoCAD.DatabaseServices.PlotType.Layout);
                if (options.PlotArea == PlotAreaChoice.Extents)
                    psv.SetPlotCentered(ps, true);
                psv.SetUseStandardScale(ps, true);
                psv.SetStdScaleType(ps, (options.ScaleToFit || options.PlotArea == PlotAreaChoice.Extents)
                    ? StdScaleType.ScaleToFit
                    : StdScaleType.StdScale1To1);
                if (!string.IsNullOrWhiteSpace(options.StyleSheet))
                    psv.SetCurrentStyleSheet(ps, options.StyleSheet);

                switch (options.Orientation)
                {
                    case PlotOrientationChoice.Portrait:
                        psv.SetPlotRotation(ps, PlotRotation.Degrees000);
                        break;
                    case PlotOrientationChoice.Landscape:
                        psv.SetPlotRotation(ps, PlotRotation.Degrees090);
                        break;
                }

                dict.SetAt(setupName, ps);
                tr.AddNewlyCreatedDBObject(ps, true);
                tr.Commit();
            }
        }

        private static void PlotDocument(
            Document doc,
            LayoutManager lm,
            List<PreparedPlotInfo> plotInfos,
            int copies,
            bool plotToFile,
            string outputPath,
            Action<string> log)
        {
            int sheetCount = plotInfos.Count * copies;
            using (var engine = PlotFactory.CreatePublishEngine())
            using (var progress = new PlotProgressDialog(false, sheetCount, true))
            {
                progress.set_PlotMsgString(PlotMessageIndex.DialogTitle, "Print Layout");
                progress.set_PlotMsgString(PlotMessageIndex.CancelJobButtonMessage, "Cancel job");
                progress.set_PlotMsgString(PlotMessageIndex.CancelSheetButtonMessage, "Cancel sheet");
                progress.set_PlotMsgString(PlotMessageIndex.SheetSetProgressCaption, "Sheet set progress");
                progress.set_PlotMsgString(PlotMessageIndex.SheetProgressCaption, "Sheet progress");
                progress.LowerPlotProgressRange = 0;
                progress.UpperPlotProgressRange = 100;
                progress.IsVisible = true;

                progress.OnBeginPlot();
                engine.BeginPlot(progress, null);

                lm.CurrentLayout = plotInfos[0].LayoutName;
                engine.BeginDocument(plotInfos[0].Info, doc.Name, null, sheetCount, plotToFile, outputPath);

                int sheetIndex = 0;
                for (int copy = 1; copy <= copies; copy++)
                {
                    foreach (var pi in plotInfos)
                    {
                        sheetIndex++;
                        log?.Invoke($"Plotting {pi.LayoutName} ({sheetIndex}/{sheetCount})...");

                        progress.StatusMsgString = $"Plotting {pi.LayoutName}";
                        progress.OnBeginSheet();
                        progress.LowerSheetProgressRange = 0;
                        progress.UpperSheetProgressRange = 100;
                        progress.SheetProgressPos = 0;

                        var pageInfo = new PlotPageInfo();
                        bool lastSheet = sheetIndex == sheetCount;
                        engine.BeginPage(pageInfo, pi.Info, lastSheet, null);
                        engine.BeginGenerateGraphics(null);
                        engine.EndGenerateGraphics(null);
                        engine.EndPage(null);

                        progress.SheetProgressPos = 100;
                        progress.OnEndSheet();
                    }
                }

                engine.EndDocument(null);
                progress.PlotProgressPos = 100;
                progress.OnEndPlot();
                engine.EndPlot(null);
            }
        }

        private class PreparedPlotInfo
        {
            public string LayoutName;
            public PlotInfo Info;
        }

        private static List<PreparedPlotInfo> BuildPlotInfos(Database db, PrintJobOptions options, LayoutManager lm)
        {
            var result = new List<PreparedPlotInfo>();
            var psv = PlotSettingsValidator.Current;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var item in options.Layouts)
                {
                    lm.CurrentLayout = item.Name;
                    db.TileMode = false;
                    var layout = tr.GetObject(item.LayoutId, OpenMode.ForRead) as Layout;
                    if (layout == null) continue;

                    var ps = new PlotSettings(layout.ModelType);
                    ps.CopyFrom(layout);
                    psv.SetPlotConfigurationName(ps, options.DeviceName, options.MediaName);
                    psv.RefreshLists(ps);
                    psv.SetPlotType(ps, options.PlotArea == PlotAreaChoice.Extents
                        ? Autodesk.AutoCAD.DatabaseServices.PlotType.Extents
                        : Autodesk.AutoCAD.DatabaseServices.PlotType.Layout);
                    if (options.PlotArea == PlotAreaChoice.Extents)
                        psv.SetPlotCentered(ps, true);
                    psv.SetUseStandardScale(ps, true);
                    psv.SetStdScaleType(ps, (options.ScaleToFit || options.PlotArea == PlotAreaChoice.Extents)
                        ? StdScaleType.ScaleToFit
                        : StdScaleType.StdScale1To1);

                    if (!string.IsNullOrWhiteSpace(options.StyleSheet))
                        psv.SetCurrentStyleSheet(ps, options.StyleSheet);

                    switch (options.Orientation)
                    {
                        case PlotOrientationChoice.Portrait:
                            psv.SetPlotRotation(ps, PlotRotation.Degrees000);
                            break;
                        case PlotOrientationChoice.Landscape:
                            psv.SetPlotRotation(ps, PlotRotation.Degrees090);
                            break;
                    }

                    var pi = new PlotInfo
                    {
                        Layout = layout.ObjectId,
                        OverrideSettings = ps
                    };
                    var piv = new PlotInfoValidator
                    {
                        MediaMatchingPolicy = MatchingPolicy.MatchEnabled
                    };
                    piv.Validate(pi);
                    result.Add(new PreparedPlotInfo
                    {
                        LayoutName = item.Name,
                        Info = pi
                    });
                }
                tr.Commit();
            }
            return result;
        }

        public static string MakeSafeFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "Layout";
            var invalid = Path.GetInvalidFileNameChars();
            var chars = name.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray();
            var safe = new string(chars).Trim();
            return safe.Length == 0 ? "Layout" : safe;
        }

        private static List<string> ToList(StringCollection strings)
        {
            var result = new List<string>();
            if (strings == null) return result;
            foreach (string s in strings)
                if (!string.IsNullOrWhiteSpace(s)) result.Add(s);
            return result;
        }

        private static int PreferredDeviceRank(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return 100;
            if (name.IndexOf("DWG To PDF", StringComparison.OrdinalIgnoreCase) >= 0) return 0;
            if (name.IndexOf("PDF", StringComparison.OrdinalIgnoreCase) >= 0) return 1;
            return 10;
        }

        private static int PreferredStyleRank(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return 100;
            if (name.Equals("monochrome.ctb", StringComparison.OrdinalIgnoreCase)) return 0;
            if (name.EndsWith(".ctb", StringComparison.OrdinalIgnoreCase)) return 1;
            if (name.EndsWith(".stb", StringComparison.OrdinalIgnoreCase)) return 2;
            return 10;
        }

        private static int MediaRank(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return 100;
            var n = name.ToUpperInvariant();
            if (n.Contains("ISO") && n.Contains("A0")) return 0;
            if (n.Contains("ISO") && n.Contains("A1")) return 1;
            if (n.Contains("ISO") && n.Contains("A2")) return 2;
            if (n.Contains("ISO") && n.Contains("A3")) return 3;
            if (n.Contains("ISO") && n.Contains("A4")) return 4;
            return 20;
        }
    }
}
