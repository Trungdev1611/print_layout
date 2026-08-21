using System;
using System.IO;
using FlexiLicense;

namespace PrintLayoutAddin.Core
{
    /// <summary>
    /// Ghi lai moi lan check license (tu Initialize hoac tung command) de biet
    /// dung lan nao thuc su goi mang, lan nao roi xuong OfflineCache -- OfflineCache
    /// luon tra ForceUpdate=false nen se lam PLAUTO/... chay tiep du server dang
    /// yeu cau cap nhat. Chi de debug, khong anh huong logic chinh.
    /// </summary>
    internal static class LicenseLog
    {
        private static string LogPath
        {
            get
            {
                var dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "Inno", "Licenses");
                Directory.CreateDirectory(dir);
                return Path.Combine(dir, "Print_layout.debug.log");
            }
        }

        public static void Write(string source, VerifyResult result)
        {
            try
            {
                var line = string.Format(
                    "{0:yyyy-MM-dd HH:mm:ss} [{1}] Status={2} ForceUpdate={3} FromCache={4} Latest={5} Msg={6}\r\n",
                    DateTime.Now, source, result.Status, result.ForceUpdate, result.FromCache,
                    result.LatestVersion, result.Message);
                File.AppendAllText(LogPath, line);
            }
            catch { /* logging khong duoc lam vo add-in */ }
        }
    }
}
