using System;
using Autodesk.AutoCAD.Geometry;

namespace PrintLayoutAddin.Core
{
    /// <summary>
    /// Paper-space geometry derived from the two PLAYOUT viewport corners.
    /// Pick order does not matter — bounds are always min/max normalized.
    /// </summary>
    public static class ViewportCornerGeometry
    {
        /// <summary>Default width (drawing units / mm on paper) of the title-block strip to the right of the viewport.</summary>
        public const double DefaultTitleStripScanWidth = 200.0;

        /// <summary>Normalized axis-aligned rectangle of the presentation / viewport area.</summary>
        public readonly struct Bounds
        {
            public Bounds(double xMin, double yMin, double xMax, double yMax)
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
            public double YMid => (YMin + YMax) * 0.5;

            public Point3d BottomLeft => new Point3d(XMin, YMin, 0);
            public Point3d TopRight => new Point3d(XMax, YMax, 0);
            public Point3d TopLeft => new Point3d(XMin, YMax, 0);
            public Point3d BottomRight => new Point3d(XMax, YMin, 0);
        }

        /// <summary>Scan / placement box in paper space.</summary>
        public readonly struct Rect
        {
            public Rect(double xMin, double yMin, double xMax, double yMax)
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

            public bool Contains(Point3d p, double margin = 0) =>
                p.X >= XMin - margin && p.X <= XMax + margin
                && p.Y >= YMin - margin && p.Y <= YMax + margin;

            public override string ToString() =>
                $"({XMin:F2},{YMin:F2})-({XMax:F2},{YMax:F2})";
        }

        public static Bounds Normalize(Point3d p1, Point3d p2) =>
            new Bounds(
                Math.Min(p1.X, p2.X),
                Math.Min(p1.Y, p2.Y),
                Math.Max(p1.X, p2.X),
                Math.Max(p1.Y, p2.Y));

        public static Bounds Normalize((Point3d P1, Point3d P2) corners) =>
            Normalize(corners.P1, corners.P2);

        /// <summary>
        /// Title-strip band for PHIÊN BẢN / revision label. Wider than the mid-only band:
        /// from 20% height up to 90% (skips bottom title/number cells' lowest zone only loosely,
        /// and the top logo / Rev table). Callers may also accept LowerBox — see scanner.
        /// </summary>
        public static Rect TitleStripMidScanBox(
            Bounds viewport,
            double stripWidth = DefaultTitleStripScanWidth)
        {
            double w = stripWidth > 1e-9 ? stripWidth : DefaultTitleStripScanWidth;
            double yLo = viewport.YMin + viewport.Height * 0.20;
            double yHi = viewport.YMin + viewport.Height * 0.90;
            if (yHi < yLo) yHi = yLo;
            return new Rect(
                viewport.XMax,
                yLo,
                viewport.XMax + w,
                yHi);
        }

        /// <summary>
        /// Full title strip to the right of the viewport (for revision fallback / diagnostics).
        /// </summary>
        public static Rect TitleStripFullScanBox(
            Bounds viewport,
            double stripWidth = DefaultTitleStripScanWidth)
        {
            double w = stripWidth > 1e-9 ? stripWidth : DefaultTitleStripScanWidth;
            return new Rect(
                viewport.XMax,
                viewport.YMin,
                viewport.XMax + w,
                viewport.YMax);
        }

        /// <summary>
        /// Title-block strip to the right of the viewport, lower half only
        /// (Sheet Title / Sheet Number labels — avoids Rev table / logo at the top).
        /// </summary>
        public static Rect TitleStripScanBox(
            Bounds viewport,
            double stripWidth = DefaultTitleStripScanWidth)
        {
            double w = stripWidth > 1e-9 ? stripWidth : DefaultTitleStripScanWidth;
            return new Rect(
                viewport.XMax,
                viewport.YMin,
                viewport.XMax + w,
                viewport.YMid);
        }

        /// <summary>
        /// Insertion point for the revision Table (= top-left of the table in AutoCAD).
        /// Anchored at the top-right corner of the viewport so the table grows into the title strip.
        /// </summary>
        public static Point3d RevTableInsertPoint(Bounds viewport) =>
            viewport.TopRight;

        /// <summary>Default distance from viewport bottom to center drawing title insertion Y.</summary>
        public const double DefaultCenterTitleLift = 40.0;

        /// <summary>
        /// Center-bottom of the presentation area for the secondary drawing title.
        /// <paramref name="lift"/> is offset above the bottom edge (default 40 drawing units).
        /// </summary>
        public static Point3d CenterTitlePoint(Bounds viewport, double textHeight = 5.0, double? lift = null)
        {
            double dy = lift ?? DefaultCenterTitleLift;
            return new Point3d(
                (viewport.XMin + viewport.XMax) * 0.5,
                viewport.YMin + dy,
                0);
        }

        /// <summary>One-line summary for command-line / log checks.</summary>
        public static string Describe(Bounds viewport, double stripWidth = DefaultTitleStripScanWidth)
        {
            var scan = TitleStripScanBox(viewport, stripWidth);
            var rev = RevTableInsertPoint(viewport);
            var center = CenterTitlePoint(viewport);
            return string.Format(
                System.Globalization.CultureInfo.InvariantCulture,
                "vp=({0:F1},{1:F1})-({2:F1},{3:F1}) W={4:F1} H={5:F1} | "
                + "scan={6} | revInsert=({7:F1},{8:F1}) | centerTitle=({9:F1},{10:F1})",
                viewport.XMin, viewport.YMin, viewport.XMax, viewport.YMax,
                viewport.Width, viewport.Height,
                scan,
                rev.X, rev.Y,
                center.X, center.Y);
        }
    }
}
