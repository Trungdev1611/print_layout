using System.Globalization;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace PrintLayoutAddin.Core
{
    /// <summary>
    /// Interactive P1/P2 pick for the presentation viewport area (shared by PLAYOUT setup UI).
    /// </summary>
    public static class ViewportCornerPicker
    {
        public static bool TryPrompt(Editor ed, out (Point3d P1, Point3d P2) corners)
        {
            corners = default;
            if (ed == null) return false;

            ed.WriteMessage(
                "\nPick 2 viewport corners (saved per DWG — used by Auto Frame Setup and Build Layouts).");

            var p1opt = new PromptPointOptions("\nSpecify first viewport corner: ")
            {
                AllowNone = false,
            };
            var p1res = ed.GetPoint(p1opt);
            if (p1res.Status != PromptStatus.OK) return false;

            var p2opt = new PromptCornerOptions("\nSpecify opposite viewport corner: ", p1res.Value);
            var p2res = ed.GetCorner(p2opt);
            if (p2res.Status != PromptStatus.OK) return false;

            corners = (p1res.Value, p2res.Value);
            return true;
        }

        /// <summary>Short summary for palette / command line.</summary>
        public static string FormatSaved((Point3d P1, Point3d P2) corners) =>
            string.Format(
                CultureInfo.InvariantCulture,
                "P1=({0:F2},{1:F2})  P2=({2:F2},{3:F2})",
                corners.P1.X, corners.P1.Y,
                corners.P2.X, corners.P2.Y);
    }
}
