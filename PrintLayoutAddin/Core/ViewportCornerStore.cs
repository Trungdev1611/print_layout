using System;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Autodesk.AutoCAD.Geometry;
using Microsoft.Win32;

namespace PrintLayoutAddin.Core
{
    /// <summary>
    /// Persists PLAYOUT viewport corner picks per drawing file (full path).
    /// Legacy global VpX1..VpY2 under <c>Software\PrintLayoutAddin</c> are migrated
    /// once into the current DWG's slot, then removed so other drawings pick fresh.
    /// </summary>
    public static class ViewportCornerStore
    {
        private const string RootKey = @"Software\PrintLayoutAddin";
        private const string CornersKey = RootKey + @"\ViewportCorners";

        /// <summary>Full path when the DWG exists on disk; otherwise null.</summary>
        public static string TryNormalizePath(string dwgPath)
        {
            if (string.IsNullOrWhiteSpace(dwgPath)) return null;
            try
            {
                if (!File.Exists(dwgPath)) return null;
                return Path.GetFullPath(dwgPath);
            }
            catch
            {
                return null;
            }
        }

        public static (Point3d P1, Point3d P2)? Load(string dwgPath)
        {
            string path = TryNormalizePath(dwgPath);
            if (path == null) return null;

            try
            {
                using (var k = Registry.CurrentUser.OpenSubKey(CornersKey + "\\" + SlotName(path)))
                {
                    if (k != null
                        && TryGetDouble(k, "X1", out var x1)
                        && TryGetDouble(k, "Y1", out var y1)
                        && TryGetDouble(k, "X2", out var x2)
                        && TryGetDouble(k, "Y2", out var y2))
                    {
                        return (new Point3d(x1, y1, 0), new Point3d(x2, y2, 0));
                    }
                }

                // One-shot migration from the old machine-wide corners.
                var legacy = LoadLegacyGlobal();
                if (!legacy.HasValue) return null;

                Save(path, legacy.Value.P1, legacy.Value.P2);
                ClearLegacyGlobal();
                return legacy;
            }
            catch
            {
                return null;
            }
        }

        public static void Save(string dwgPath, Point3d p1, Point3d p2)
        {
            string path = TryNormalizePath(dwgPath);
            if (path == null) return;

            try
            {
                using (var k = Registry.CurrentUser.CreateSubKey(CornersKey + "\\" + SlotName(path)))
                {
                    if (k == null) return;
                    k.SetValue("Path", path);
                    k.SetValue("X1", p1.X.ToString("R", CultureInfo.InvariantCulture));
                    k.SetValue("Y1", p1.Y.ToString("R", CultureInfo.InvariantCulture));
                    k.SetValue("X2", p2.X.ToString("R", CultureInfo.InvariantCulture));
                    k.SetValue("Y2", p2.Y.ToString("R", CultureInfo.InvariantCulture));
                }
            }
            catch { }
        }

        public static void Save(string dwgPath, (Point3d P1, Point3d P2) corners) =>
            Save(dwgPath, corners.P1, corners.P2);

        /// <summary>Clears corners for this DWG only.</summary>
        public static bool Clear(string dwgPath)
        {
            string path = TryNormalizePath(dwgPath);
            if (path == null) return false;

            try
            {
                using (var parent = Registry.CurrentUser.OpenSubKey(CornersKey, writable: true))
                {
                    if (parent == null) return false;
                    string slot = SlotName(path);
                    try { parent.DeleteSubKeyTree(slot); }
                    catch { return false; }
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        private static string SlotName(string normalizedFullPath)
        {
            // Registry subkey names cannot contain '\' — hash the path.
            using (var sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(
                    normalizedFullPath.ToUpperInvariant()));
                var sb = new StringBuilder(hash.Length * 2);
                foreach (byte b in hash)
                    sb.Append(b.ToString("x2", CultureInfo.InvariantCulture));
                return sb.ToString();
            }
        }

        private static (Point3d P1, Point3d P2)? LoadLegacyGlobal()
        {
            try
            {
                using (var k = Registry.CurrentUser.OpenSubKey(RootKey))
                {
                    if (k == null) return null;
                    if (!TryGetDouble(k, "VpX1", out var x1)) return null;
                    if (!TryGetDouble(k, "VpY1", out var y1)) return null;
                    if (!TryGetDouble(k, "VpX2", out var x2)) return null;
                    if (!TryGetDouble(k, "VpY2", out var y2)) return null;
                    return (new Point3d(x1, y1, 0), new Point3d(x2, y2, 0));
                }
            }
            catch
            {
                return null;
            }
        }

        private static void ClearLegacyGlobal()
        {
            try
            {
                using (var k = Registry.CurrentUser.CreateSubKey(RootKey))
                {
                    if (k == null) return;
                    k.DeleteValue("VpX1", false);
                    k.DeleteValue("VpY1", false);
                    k.DeleteValue("VpX2", false);
                    k.DeleteValue("VpY2", false);
                }
            }
            catch { }
        }

        private static bool TryGetDouble(RegistryKey k, string name, out double v)
        {
            v = 0;
            var o = k.GetValue(name);
            if (o == null) return false;
            return double.TryParse(o.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out v);
        }
    }
}
