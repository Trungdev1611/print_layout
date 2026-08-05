using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace PrintLayoutAddin.LicenseGenerator
{
    // Shared license-key logic used by both the GUI (MainForm) and the CLI (Program).
    //
    // !!! The secret here MUST stay byte-for-byte identical to
    //     PrintLayoutAddin/Core/LicenseManager.BuildSecret() in the addin. !!!
    internal static class LicenseCore
    {
        public const string KeyPrefix = "PLA1-";

        // Column-header aliases for CSV / pasted-table detection (normalized form).
        public static readonly string[] MidAliases =
            { "machineid", "machine", "mid", "id", "mamay", "maid" };
        public static readonly string[] ExpAliases =
            { "expire", "expiry", "expiration", "expiredate", "expirydate",
              "hethan", "ngayhethan", "ngayhet", "han", "hansudung", "handung", "hsd" };
        // "user" and "note" are kept separate so a sheet with BOTH columns loses neither;
        // they are merged into the single name/note field (see CombineNote).
        public static readonly string[] UserAliases =
            { "user", "username", "nguoidung", "ten", "hoten", "name", "customer", "khachhang", "khach" };
        public static readonly string[] NoteAliases =
            { "note", "notes", "ghichu", "company", "congty", "duan", "project" };

        // Merge the user/name column and the note column into the single note that gets
        // embedded in the key. If both are present they are joined; otherwise the one with data.
        public static string CombineNote(string user, string note)
        {
            user = (user ?? "").Trim();
            note = (note ?? "").Trim();
            if (user.Length > 0 && note.Length > 0) return user + " - " + note;
            return user.Length > 0 ? user : note;
        }

        private static byte[] BuildSecret()
        {
            byte[] a = Encoding.UTF8.GetBytes("HxZRnVKIvZcnsvBc5zbo4h3qh9T21Yry");
            byte[] b = Encoding.UTF8.GetBytes("n9J2EV8LsFiOkITbiayLbT3yu2pqC3Jh");
            byte[] c = Encoding.UTF8.GetBytes("QG6TFg25AzVs4RwtzNyi4W1hyZ2WhNyx");
            int n = Math.Max(a.Length, Math.Max(b.Length, c.Length));
            byte[] r = new byte[n];
            for (int i = 0; i < n; i++)
                r[i] = (byte)(a[i % a.Length] ^ b[i % b.Length] ^ c[i % c.Length]);
            return r;
        }

        public static string MakeKey(string machineId, DateTime expire, string note, out string payload)
        {
            note = (note ?? "").Replace(';', ',').Replace('=', '-');
            string issued = DateTime.UtcNow.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            payload = string.Format("m={0};e={1};i={2};n={3}",
                (machineId ?? "").Trim(),
                expire.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                issued,
                note);

            byte[] payloadBytes = Encoding.UTF8.GetBytes(payload);
            byte[] sig;
            using (var hmac = new HMACSHA256(BuildSecret()))
                sig = hmac.ComputeHash(payloadBytes);
            byte[] sig16 = new byte[16];
            Array.Copy(sig, sig16, 16);

            return KeyPrefix + Base64UrlEncode(payloadBytes) + "." + Base64UrlEncode(sig16);
        }

        // Accept the common date formats people type / Excel & Sheets export, then
        // normalize to YYYY-MM-DD. Ambiguous d/M vs M/d resolves M/d first (US export).
        public static bool TryParseDate(string s, out DateTime d)
        {
            d = default(DateTime);
            s = (s ?? "").Trim();
            if (s.Length == 0) return false;
            string[] formats =
            {
                "yyyy-MM-dd", "yyyy/MM/dd", "yyyy.MM.dd",
                "M/d/yyyy", "MM/dd/yyyy",
                "d/M/yyyy", "dd/MM/yyyy",
                "d-M-yyyy", "dd-MM-yyyy",
                "d.M.yyyy", "dd.MM.yyyy",
            };
            if (DateTime.TryParseExact(s, formats, CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out d))
                return true;
            return DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.None, out d);
        }

        // ---------- delimited-table helpers (CSV from file, TSV from clipboard) ----------

        public static List<string[]> ParseDelimited(string text, char delimiter)
        {
            var rows = new List<string[]>();
            var record = new List<string>();
            var field = new StringBuilder();
            bool inQuotes = false;
            int i = 0, n = text.Length;

            while (i < n)
            {
                char c = text[i];
                if (inQuotes)
                {
                    if (c == '"')
                    {
                        if (i + 1 < n && text[i + 1] == '"') { field.Append('"'); i += 2; continue; }
                        inQuotes = false; i++; continue;
                    }
                    field.Append(c); i++; continue;
                }

                if (c == '"') { inQuotes = true; i++; continue; }
                if (c == delimiter) { record.Add(field.ToString()); field.Clear(); i++; continue; }
                if (c == '\r' || c == '\n')
                {
                    record.Add(field.ToString()); field.Clear();
                    rows.Add(record.ToArray()); record = new List<string>();
                    if (c == '\r' && i + 1 < n && text[i + 1] == '\n') i += 2; else i++;
                    continue;
                }
                field.Append(c); i++;
            }
            if (field.Length > 0 || record.Count > 0)
            {
                record.Add(field.ToString());
                rows.Add(record.ToArray());
            }
            return rows.Where(r => !(r.Length == 1 && string.IsNullOrWhiteSpace(r[0]))).ToList();
        }

        // Auto-detect tab (clipboard paste from Sheet/Excel) vs comma (downloaded CSV).
        public static List<string[]> ParseTable(string text)
        {
            char delim = text.Contains("\t") ? '\t' : ',';
            return ParseDelimited(text, delim);
        }

        public static string CsvEscape(string s)
        {
            s = s ?? "";
            if (s.IndexOf(',') >= 0 || s.IndexOf('"') >= 0 || s.IndexOf('\n') >= 0 || s.IndexOf('\r') >= 0)
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }

        // Alias-priority: earlier aliases win, so an exact "note" column beats a
        // softer "user"/"company" alias even if "user" sits in an earlier column.
        public static int FindColumn(string[] header, string[] aliases)
        {
            var norm = header.Select(NormalizeHeader).ToArray();
            foreach (var a in aliases)
                for (int i = 0; i < norm.Length; i++)
                    if (norm[i] == a) return i;
            return -1;
        }

        // Lowercase, strip diacritics + separators so "Ngày hết hạn" -> "ngayhethan".
        public static string NormalizeHeader(string s)
        {
            if (s == null) return "";
            s = s.Trim().ToLowerInvariant().Replace("đ", "d");
            var formD = s.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder();
            foreach (char ch in formD)
                if (CharUnicodeInfo.GetUnicodeCategory(ch) != UnicodeCategory.NonSpacingMark)
                    sb.Append(ch);
            s = sb.ToString().Normalize(NormalizationForm.FormC);
            return s.Replace(" ", "").Replace("_", "").Replace("-", "").Replace(".", "");
        }

        // ---------- logging ----------

        public static string LogPath()
        {
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "issued_keys.log");
        }

        public static void LogIssued(string machineId, DateTime expire, string note, string key)
        {
            try
            {
                File.AppendAllText(LogPath(),
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "\t" +
                    (machineId ?? "").Trim() + "\t" +
                    expire.ToString("yyyy-MM-dd") + "\t" +
                    (note ?? "").Replace('\t', ' ') + "\t" +
                    key + Environment.NewLine,
                    new UTF8Encoding(false));
            }
            catch { /* logging is best-effort */ }
        }

        private static string Base64UrlEncode(byte[] data)
        {
            return Convert.ToBase64String(data).TrimEnd('=').Replace('+', '-').Replace('/', '_');
        }
    }
}
