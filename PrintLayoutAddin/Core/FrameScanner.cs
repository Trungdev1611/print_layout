using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace PrintLayoutAddin.Core
{
    public class FrameInfo
    {
        public ObjectId BlockRefId { get; set; }
        public string Stt { get; set; }
        public Point3d Min { get; set; }
        public Point3d Max { get; set; }

        public Point3d Center => new Point3d((Min.X + Max.X) / 2.0, (Min.Y + Max.Y) / 2.0, 0.0);
        public double Width => Max.X - Min.X;
        public double Height => Max.Y - Min.Y;
    }

    public class BlockChoice
    {
        public string Name { get; set; }
        public int Count { get; set; }
        public bool IsXref { get; set; }

        public override string ToString()
        {
            return $"{Name}  ({Count} instance{(Count == 1 ? "" : "s")}{(IsXref ? ", XREF" : "")})";
        }
    }

    public static class FrameScanner
    {
        public static List<BlockChoice> ListBlocksInModelSpace(Database db)
        {
            var result = new Dictionary<string, BlockChoice>(StringComparer.OrdinalIgnoreCase);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var ms = (BlockTableRecord)tr.GetObject(
                    SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

                foreach (ObjectId id in ms)
                {
                    if (id.ObjectClass.DxfName != "INSERT") continue;
                    var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                    if (br == null) continue;

                    var btr = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                    var name = br.Name;

                    if (!result.TryGetValue(name, out var choice))
                    {
                        choice = new BlockChoice
                        {
                            Name = name,
                            Count = 0,
                            IsXref = btr.IsFromExternalReference
                        };
                        result[name] = choice;
                    }
                    choice.Count++;
                }
                tr.Commit();
            }

            return result.Values.OrderBy(c => c.Name).ToList();
        }

        public static List<FrameInfo> CollectFrames(Database db, string blockName, bool requireStt)
        {
            var cfg = Config.Instance;
            var frames = new List<FrameInfo>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var ms = (BlockTableRecord)tr.GetObject(
                    SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);

                foreach (ObjectId id in ms)
                {
                    if (id.ObjectClass.DxfName != "INSERT") continue;
                    var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                    if (br == null) continue;
                    if (!string.Equals(br.Name, blockName, StringComparison.OrdinalIgnoreCase)) continue;

                    Extents3d ext;
                    try { ext = br.GeometricExtents; }
                    catch { continue; }

                    var info = new FrameInfo
                    {
                        BlockRefId = id,
                        Min = ext.MinPoint,
                        Max = ext.MaxPoint,
                        Stt = ReadStt(br, cfg)
                    };

                    if (requireStt && string.IsNullOrWhiteSpace(info.Stt)) continue;
                    frames.Add(info);
                }
                tr.Commit();
            }

            return frames;
        }

        public static string ReadStt(BlockReference br, Config cfg)
        {
            foreach (ObjectId attId in br.AttributeCollection)
            {
                var att = attId.GetObject(OpenMode.ForRead) as AttributeReference;
                if (att == null) continue;
                if (string.Equals(att.Tag, cfg.AttributeTag, StringComparison.OrdinalIgnoreCase))
                {
                    if (!string.IsNullOrEmpty(att.TextString)) return att.TextString;
                }
            }

            var rb = br.GetXDataForApplication(cfg.XdataAppName);
            if (rb != null)
            {
                try
                {
                    foreach (var tv in rb)
                    {
                        if (tv.TypeCode == (int)Autodesk.AutoCAD.DatabaseServices.DxfCode.ExtendedDataAsciiString)
                        {
                            var s = tv.Value?.ToString();
                            if (!string.IsNullOrEmpty(s)) return s;
                        }
                    }
                }
                finally { rb.Dispose(); }
            }
            return null;
        }

        public static string ReadDrawingName(BlockReference br, Config cfg)
        {
            return ReadAttribute(br, cfg.DrawingNameTag);
        }

        public static Dictionary<string, string> CollectDrawingNamesByStt(Database db)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var cfg = Config.Instance;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var ms = (BlockTableRecord)tr.GetObject(
                    SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);
                foreach (ObjectId id in ms)
                {
                    if (id.ObjectClass.DxfName != "INSERT") continue;
                    var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                    if (br == null) continue;
                    var stt = ReadStt(br, cfg);
                    if (string.IsNullOrWhiteSpace(stt)) continue;
                    var drawingName = ReadDrawingName(br, cfg);
                    if (!string.IsNullOrWhiteSpace(drawingName))
                        result[stt.Trim()] = drawingName.Trim();
                }
                tr.Commit();
            }
            return result;
        }

        public static bool WriteStt(BlockReference br, string value, Config cfg, Transaction tr)
        {
            if (WriteAttribute(br, cfg.AttributeTag, value, tr)) return true;

            // STT keeps its legacy XData fallback because downstream layout
            // generation can read it when an xref does not expose attributes.
            EnsureRegApp(br.Database, cfg.XdataAppName, tr);
            var brW = (BlockReference)tr.GetObject(br.ObjectId, OpenMode.ForWrite);
            brW.XData = new ResultBuffer(
                new TypedValue((int)DxfCode.ExtendedDataRegAppName, cfg.XdataAppName),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, value ?? ""));
            return true;
        }

        public static bool WriteDrawingName(BlockReference br, string value, Config cfg, Transaction tr)
        {
            if (string.IsNullOrWhiteSpace(value)) return false;
            return WriteAttribute(br, cfg.DrawingNameTag, value, tr);
        }

        private static bool WriteAttribute(BlockReference br, string tag, string value, Transaction tr)
        {
            if (string.IsNullOrWhiteSpace(tag)) return false;
            foreach (ObjectId attId in br.AttributeCollection)
            {
                var att = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                if (att == null) continue;
                if (string.Equals(att.Tag, tag, StringComparison.OrdinalIgnoreCase))
                {
                    att.TextString = value;
                    return true;
                }
            }
            return false;
        }

        private static string ReadAttribute(BlockReference br, string tag)
        {
            if (string.IsNullOrWhiteSpace(tag)) return null;
            foreach (ObjectId attId in br.AttributeCollection)
            {
                var att = attId.GetObject(OpenMode.ForRead) as AttributeReference;
                if (att == null) continue;
                if (string.Equals(att.Tag, tag, StringComparison.OrdinalIgnoreCase))
                    return att.TextString;
            }
            return null;
        }

        private static void EnsureRegApp(Database db, string appName, Transaction tr)
        {
            var regTable = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);
            if (regTable.Has(appName)) return;
            regTable.UpgradeOpen();
            var rec = new RegAppTableRecord { Name = appName };
            regTable.Add(rec);
            tr.AddNewlyCreatedDBObject(rec, true);
        }

        // Sort STT strings: if a trailing numeric segment exists, sort by that number; else by string.
        public static IComparer<string> SttComparer { get; } = new SttStringComparer();

        private class SttStringComparer : IComparer<string>
        {
            private static readonly Regex TailNum = new Regex(@"^(?<prefix>.*?)(?<num>\d+)(?<suffix>\D*)$");

            public int Compare(string x, string y)
            {
                if (x == null && y == null) return 0;
                if (x == null) return -1;
                if (y == null) return 1;

                var mx = TailNum.Match(x);
                var my = TailNum.Match(y);
                if (mx.Success && my.Success
                    && mx.Groups["prefix"].Value == my.Groups["prefix"].Value
                    && mx.Groups["suffix"].Value == my.Groups["suffix"].Value)
                {
                    long nx = long.Parse(mx.Groups["num"].Value);
                    long ny = long.Parse(my.Groups["num"].Value);
                    return nx.CompareTo(ny);
                }
                return string.Compare(x, y, StringComparison.OrdinalIgnoreCase);
            }
        }
    }
}
