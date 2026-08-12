using System;
using System.Collections.Generic;
using System.Globalization;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace PrintLayoutAddin.Core
{
    /// <summary>
    /// Finds the axis-aligned title-block cell around a label by collecting
    /// horizontal / vertical Line and Polyline edges (including inside blocks / xrefs).
    /// </summary>
    public static class TitleBlockCellFinder
    {
        public const double DefaultClusterTol = 0.35;
        /// <summary>Small right nudge so seed is inside the cell, not on the left wall.</summary>
        public const double DefaultSeedNudgeX = 2.0;
        /// <summary>Tiny down nudge — large values drop into the cell below (PHIÊN BẢN).</summary>
        public const double DefaultSeedNudgeY = 0.5;
        /// <summary>Ignore tiny ticks; cell walls are longer than decorative stubs.</summary>
        public const double DefaultMinSegment = 5.0;
        /// <summary>Vertical must cover this fraction of the Y-band to count as a cell wall.</summary>
        public const double MinVerticalCoverage = 0.55;

        public readonly struct Cell
        {
            public Cell(double xMin, double yMin, double xMax, double yMax)
            {
                XMin = xMin;
                YMin = yMin;
                XMax = xMax;
                YMax = yMax;
            }

            public double XMin { get; }
            public double YMin { get; }
            public double XMax { get; }
            public double YMax { get; }
            public double Width => XMax - XMin;
            public double Height => YMax - YMin;
            public Point3d Center => new Point3d((XMin + XMax) * 0.5, (YMin + YMax) * 0.5, 0);
            public Point3d BottomLeft => new Point3d(XMin, YMin, 0);
            public Point3d BottomRight => new Point3d(XMax, YMin, 0);
            public Point3d TopRight => new Point3d(XMax, YMax, 0);
            public Point3d TopLeft => new Point3d(XMin, YMax, 0);

            public bool IsValid => Width > 1e-6 && Height > 1e-6;

            public string FormatCorners() =>
                string.Format(
                    CultureInfo.InvariantCulture,
                    "BL=({0:F2},{1:F2}) BR=({2:F2},{3:F2}) TR=({4:F2},{5:F2}) TL=({6:F2},{7:F2})",
                    BottomLeft.X, BottomLeft.Y,
                    BottomRight.X, BottomRight.Y,
                    TopRight.X, TopRight.Y,
                    TopLeft.X, TopLeft.Y);

            public override string ToString() =>
                string.Format(
                    CultureInfo.InvariantCulture,
                    "({0:F2},{1:F2})-({2:F2},{3:F2}) c=({4:F2},{5:F2})",
                    XMin, YMin, XMax, YMax, Center.X, Center.Y);
        }

        readonly struct Segment
        {
            public Segment(Point3d a, Point3d b)
            {
                A = a;
                B = b;
            }

            public Point3d A { get; }
            public Point3d B { get; }
            public double MidX => (A.X + B.X) * 0.5;
            public double MidY => (A.Y + B.Y) * 0.5;
            public double Len
            {
                get
                {
                    double dx = B.X - A.X;
                    double dy = B.Y - A.Y;
                    return Math.Sqrt(dx * dx + dy * dy);
                }
            }
        }

        /// <summary>
        /// Build a grid from linework near the title strip, then return the cell
        /// containing a point nudged slightly inward from the label insert.
        /// </summary>
        public static bool TryFindCell(
            Database db,
            Point3d labelPosition,
            ViewportCornerGeometry.Rect searchArea,
            out Cell cell,
            out string detail,
            double clusterTol = DefaultClusterTol,
            double seedNudgeX = DefaultSeedNudgeX,
            double seedNudgeY = DefaultSeedNudgeY)
        {
            cell = default;
            detail = "";
            if (db == null)
            {
                detail = "database null";
                return false;
            }

            // Keep seed in the SAME row as the label (do not drop into the cell below).
            var seed = new Point3d(
                labelPosition.X + Math.Abs(seedNudgeX),
                labelPosition.Y - Math.Abs(seedNudgeY),
                0);

            var horizontals = new List<Segment>();
            var verticals = new List<Segment>();

            var area = new ViewportCornerGeometry.Rect(
                searchArea.XMin - 5,
                searchArea.YMin - 5,
                searchArea.XMax + 5,
                searchArea.YMax + 5);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
                var visited = new HashSet<ObjectId>();
                CollectSegments(space, Matrix3d.Identity, tr, area, horizontals, verticals, visited);
                tr.Commit();
            }

            var ys = new List<double>(horizontals.Count);
            foreach (var h in horizontals)
                ys.Add(h.MidY);

            var gridY = Cluster(ys, clusterTol);
            if (gridY.Count < 2)
            {
                detail = string.Format(
                    CultureInfo.InvariantCulture,
                    "not enough horizontal lines (Y={0})",
                    gridY.Count);
                return false;
            }

            if (!TryBracket(gridY, seed.Y, out double yMin, out double yMax))
            {
                detail = string.Format(
                    CultureInfo.InvariantCulture,
                    "seed Y={0:F2} not inside a row (Ylines={1})",
                    seed.Y, gridY.Count);
                return false;
            }

            // Label insert must stay in this row — otherwise we picked the wrong band.
            if (labelPosition.Y < yMin - 0.05 || labelPosition.Y > yMax + 0.05)
            {
                detail = string.Format(
                    CultureInfo.InvariantCulture,
                    "label Y={0:F2} outside row ({1:F2},{2:F2}) seed=({3:F2},{4:F2})",
                    labelPosition.Y, yMin, yMax, seed.X, seed.Y);
                return false;
            }

            double rowH = yMax - yMin;
            var xs = new List<double>();
            foreach (var v in verticals)
            {
                double overlap = VerticalOverlap(v, yMin, yMax);
                if (overlap < DefaultMinSegment) continue;
                if (rowH > 1e-9 && overlap / rowH < MinVerticalCoverage) continue;
                xs.Add(v.MidX);
            }

            var gridX = Cluster(xs, clusterTol);
            if (gridX.Count < 2)
            {
                detail = string.Format(
                    CultureInfo.InvariantCulture,
                    "not enough spanning verticals in row H={0:F2} (X={1})",
                    rowH, gridX.Count);
                return false;
            }

            if (!TryBracket(gridX, seed.X, out double xMin, out double xMax))
            {
                // Label often sits on the left wall — try slightly further right once.
                double seedX2 = seed.X + 8.0;
                if (!TryBracket(gridX, seedX2, out xMin, out xMax))
                {
                    detail = string.Format(
                        CultureInfo.InvariantCulture,
                        "seed X={0:F2} not inside a column (Xlines={1})",
                        seed.X, gridX.Count);
                    return false;
                }
            }

            cell = new Cell(xMin, yMin, xMax, yMax);
            if (!cell.IsValid || cell.Width > 500 || cell.Height > 500)
            {
                detail = "cell size unrealistic: " + cell;
                cell = default;
                return false;
            }

            detail = string.Format(
                CultureInfo.InvariantCulture,
                "cell {0} W={1:F2} H={2:F2} seed=({3:F2},{4:F2}) Xspan={5} Yspan={6}",
                cell, cell.Width, cell.Height, seed.X, seed.Y, gridX.Count, gridY.Count);
            return true;
        }

        static double VerticalOverlap(Segment v, double yMin, double yMax)
        {
            double a = Math.Min(v.A.Y, v.B.Y);
            double b = Math.Max(v.A.Y, v.B.Y);
            double lo = Math.Max(a, yMin);
            double hi = Math.Min(b, yMax);
            return Math.Max(0, hi - lo);
        }

        static bool TryBracket(List<double> sorted, double value, out double lo, out double hi)
        {
            lo = hi = 0;
            if (sorted == null || sorted.Count < 2) return false;

            double? left = null, right = null;
            foreach (double v in sorted)
            {
                if (v < value - 1e-9) left = v;
                else if (v > value + 1e-9)
                {
                    right = v;
                    break;
                }
            }

            if (!left.HasValue || !right.HasValue) return false;
            lo = left.Value;
            hi = right.Value;
            return hi > lo + 1e-9;
        }

        static List<double> Cluster(List<double> raw, double tol)
        {
            var list = new List<double>();
            if (raw == null || raw.Count == 0) return list;
            raw.Sort();
            double sum = raw[0];
            int n = 1;
            double cur = raw[0];
            for (int i = 1; i < raw.Count; i++)
            {
                if (Math.Abs(raw[i] - cur) <= tol)
                {
                    sum += raw[i];
                    n++;
                    cur = sum / n;
                }
                else
                {
                    list.Add(cur);
                    sum = raw[i];
                    n = 1;
                    cur = raw[i];
                }
            }
            list.Add(cur);
            return list;
        }

        static void CollectSegments(
            BlockTableRecord space,
            Matrix3d toPaper,
            Transaction tr,
            ViewportCornerGeometry.Rect area,
            List<Segment> horizontals,
            List<Segment> verticals,
            HashSet<ObjectId> visitedBlocks)
        {
            if (space == null || space.ObjectId.IsNull) return;
            if (!visitedBlocks.Add(space.ObjectId)) return;

            foreach (ObjectId id in space)
            {
                if (id.IsNull || id.IsErased) continue;
                var ent = tr.GetObject(id, OpenMode.ForRead, false) as Entity;
                if (ent == null) continue;

                if (ent is Line line)
                {
                    AddSegment(
                        line.StartPoint.TransformBy(toPaper),
                        line.EndPoint.TransformBy(toPaper),
                        area, horizontals, verticals);
                    continue;
                }

                if (ent is Polyline pl)
                {
                    int n = pl.NumberOfVertices;
                    for (int i = 0; i < n; i++)
                    {
                        Point3d a = pl.GetPoint3dAt(i).TransformBy(toPaper);
                        Point3d b = pl.GetPoint3dAt((i + 1) % n).TransformBy(toPaper);
                        if (!pl.Closed && i == n - 1) break;
                        AddSegment(a, b, area, horizontals, verticals);
                    }
                    continue;
                }

                if (ent is Polyline2d pl2)
                {
                    var pts = new List<Point3d>();
                    foreach (ObjectId vid in pl2)
                    {
                        if (tr.GetObject(vid, OpenMode.ForRead, false) is Vertex2d v)
                            pts.Add(new Point3d(v.Position.X, v.Position.Y, 0).TransformBy(toPaper));
                    }
                    for (int i = 0; i + 1 < pts.Count; i++)
                        AddSegment(pts[i], pts[i + 1], area, horizontals, verticals);
                    if (pl2.Closed && pts.Count >= 2)
                        AddSegment(pts[pts.Count - 1], pts[0], area, horizontals, verticals);
                    continue;
                }

                if (ent is BlockReference br)
                {
                    try
                    {
                        var btr = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                        CollectSegments(
                            btr, toPaper * br.BlockTransform, tr, area, horizontals, verticals, visitedBlocks);
                    }
                    catch { }
                }
            }
        }

        static void AddSegment(
            Point3d a,
            Point3d b,
            ViewportCornerGeometry.Rect area,
            List<Segment> horizontals,
            List<Segment> verticals)
        {
            double dx = b.X - a.X;
            double dy = b.Y - a.Y;
            double len = Math.Sqrt(dx * dx + dy * dy);
            if (len < DefaultMinSegment) return;

            double mx = (a.X + b.X) * 0.5;
            double my = (a.Y + b.Y) * 0.5;
            if (!area.Contains(new Point3d(mx, my, 0), margin: 20)
                && !area.Contains(a, 20)
                && !area.Contains(b, 20))
                return;

            const double axisTol = 0.15;
            double ax = Math.Abs(dx);
            double ay = Math.Abs(dy);
            var seg = new Segment(a, b);

            if (ay <= axisTol * len || ay < 0.5)
                horizontals.Add(seg);
            else if (ax <= axisTol * len || ax < 0.5)
                verticals.Add(seg);
        }
    }
}
