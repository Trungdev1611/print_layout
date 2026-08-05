using System;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Win32;

namespace PrintLayoutAddin.Core
{
    public enum LicenseState
    {
        Valid,
        NotActivated,
        Expired,
        WrongMachine,
        Corrupted,
    }

    public class LicenseInfo
    {
        public LicenseState State { get; set; }
        public string MachineId { get; set; }
        public DateTime? Expiration { get; set; }
        public DateTime? Issued { get; set; }
        public string Note { get; set; }
        public string Message { get; set; }

        public bool IsValid { get { return State == LicenseState.Valid; } }

        public int DaysRemaining
        {
            get
            {
                if (Expiration == null) return -1;
                return (int)Math.Ceiling((Expiration.Value.Date - DateTime.UtcNow.Date).TotalDays);
            }
        }
    }

    // License key format:
    //     PLA1-<base64url(payload)>.<base64url(signature)>
    //
    // Payload (UTF-8 plain text, semicolon-separated key=value):
    //     m=<MachineId>;e=YYYY-MM-DD;i=YYYY-MM-DD;n=<note>
    //
    // Signature: first 16 bytes of HMAC-SHA256(secret, payloadBytes).
    //
    // The same protocol is implemented in license-generator/Program.cs.
    // If you change the secret or format here, update the generator too.
    public static class LicenseManager
    {
        // !!! CHANGE THIS SECRET BEFORE DISTRIBUTING TO CUSTOMERS !!!
        // Keep the secret identical in the LicenseGenerator tool.
        // Obfuscated as XOR of three parts to make casual extraction harder.
        // A determined attacker with a .NET decompiler can still recover it.
        private static readonly byte[] Secret = BuildSecret();

        private const string KeyPrefix = "PLA1-";
        public const int WarnDaysThreshold = 7;

        private static readonly object _lock = new object();
        private static LicenseInfo _cached;

        public static LicenseInfo Current
        {
            get
            {
                lock (_lock)
                {
                    if (_cached == null) _cached = Load();
                    return _cached;
                }
            }
        }

        public static void Refresh()
        {
            lock (_lock) { _cached = Load(); }
        }

        public static string LicenseFilePath
        {
            get
            {
                var dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "PrintLayoutAddin");
                Directory.CreateDirectory(dir);
                return Path.Combine(dir, "license.dat");
            }
        }

        public static string GetMachineId()
        {
            string raw = ReadMachineGuid();
            if (string.IsNullOrEmpty(raw))
            {
                // Fallback when MachineGuid is missing — combine hostname + user SID-ish info.
                raw = Environment.MachineName + "|" + Environment.UserDomainName + "|" + Environment.OSVersion.VersionString;
            }
            using (var sha = SHA256.Create())
            {
                var hash = sha.ComputeHash(Encoding.UTF8.GetBytes("PLADDIN-MID-v1:" + raw));
                string body = Base32Encode(hash, 10); // 10 bytes -> 16 base32 chars
                return string.Format("{0}-{1}-{2}-{3}",
                    body.Substring(0, 4),
                    body.Substring(4, 4),
                    body.Substring(8, 4),
                    body.Substring(12, 4));
            }
        }

        public static LicenseInfo Validate(string key)
        {
            var info = new LicenseInfo { MachineId = GetMachineId() };

            if (string.IsNullOrWhiteSpace(key))
            {
                info.State = LicenseState.NotActivated;
                info.Message = "License not activated.";
                return info;
            }

            key = key.Trim();
            if (!key.StartsWith(KeyPrefix, StringComparison.OrdinalIgnoreCase))
            {
                info.State = LicenseState.Corrupted;
                info.Message = "Invalid key format (missing prefix).";
                return info;
            }

            string rest = key.Substring(KeyPrefix.Length);
            int dot = rest.IndexOf('.');
            if (dot <= 0 || dot >= rest.Length - 1)
            {
                info.State = LicenseState.Corrupted;
                info.Message = "Invalid key format (missing signature).";
                return info;
            }

            byte[] payloadBytes, sigBytes;
            try
            {
                payloadBytes = Base64UrlDecode(rest.Substring(0, dot));
                sigBytes = Base64UrlDecode(rest.Substring(dot + 1));
            }
            catch
            {
                info.State = LicenseState.Corrupted;
                info.Message = "Invalid key encoding.";
                return info;
            }

            if (sigBytes.Length != 16)
            {
                info.State = LicenseState.Corrupted;
                info.Message = "Invalid signature length.";
                return info;
            }

            byte[] expected;
            using (var hmac = new HMACSHA256(Secret))
                expected = hmac.ComputeHash(payloadBytes);

            // Constant-time compare on the first 16 bytes.
            int diff = 0;
            for (int i = 0; i < 16; i++) diff |= sigBytes[i] ^ expected[i];
            if (diff != 0)
            {
                info.State = LicenseState.Corrupted;
                info.Message = "Signature mismatch. The key is tampered or invalid.";
                return info;
            }

            string payload;
            try { payload = Encoding.UTF8.GetString(payloadBytes); }
            catch
            {
                info.State = LicenseState.Corrupted;
                info.Message = "Unreadable payload.";
                return info;
            }

            string mid = null, expStr = null, issStr = null, note = null;
            foreach (var kv in payload.Split(';'))
            {
                int idx = kv.IndexOf('=');
                if (idx <= 0) continue;
                string k = kv.Substring(0, idx);
                string v = kv.Substring(idx + 1);
                switch (k)
                {
                    case "m": mid = v; break;
                    case "e": expStr = v; break;
                    case "i": issStr = v; break;
                    case "n": note = v; break;
                }
            }

            DateTime exp;
            if (!DateTime.TryParseExact(expStr, "yyyy-MM-dd", CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out exp))
            {
                info.State = LicenseState.Corrupted;
                info.Message = "Invalid expiration date.";
                return info;
            }
            info.Expiration = exp;
            info.Note = note;

            DateTime iss;
            if (DateTime.TryParseExact(issStr, "yyyy-MM-dd", CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out iss))
                info.Issued = iss;

            if (!string.Equals(mid, info.MachineId, StringComparison.OrdinalIgnoreCase))
            {
                info.State = LicenseState.WrongMachine;
                info.Message = "Key was issued for a different machine.";
                return info;
            }

            if (DateTime.UtcNow.Date > exp.Date)
            {
                info.State = LicenseState.Expired;
                info.Message = "License expired on " + exp.ToString("yyyy-MM-dd") + " (UTC).";
                return info;
            }

            info.State = LicenseState.Valid;
            info.Message = "License valid until " + exp.ToString("yyyy-MM-dd") + ".";
            return info;
        }

        public static LicenseInfo Activate(string key)
        {
            var info = Validate(key);
            if (info.State == LicenseState.Valid)
            {
                try
                {
                    File.WriteAllText(LicenseFilePath, key.Trim(), new UTF8Encoding(false));
                }
                catch (Exception ex)
                {
                    info.State = LicenseState.Corrupted;
                    info.Message = "Could not save license file: " + ex.Message;
                }
            }
            lock (_lock) { _cached = info; }
            return info;
        }

        public static void Clear()
        {
            try { File.Delete(LicenseFilePath); }
            catch { }
            lock (_lock) { _cached = null; }
        }

        // ---------- internals ----------

        private static LicenseInfo Load()
        {
            string key = null;
            try
            {
                if (File.Exists(LicenseFilePath))
                    key = File.ReadAllText(LicenseFilePath, Encoding.UTF8);
            }
            catch { }
            return Validate(key);
        }

        private static string ReadMachineGuid()
        {
            try
            {
                using (var hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
                using (var k = hklm.OpenSubKey(@"SOFTWARE\Microsoft\Cryptography"))
                {
                    return k != null ? k.GetValue("MachineGuid") as string : null;
                }
            }
            catch { return null; }
        }

        private static byte[] BuildSecret()
        {
            // Three parts XOR'd together. Must stay identical to BuildSecret() in
            // license-generator/LicenseCore.cs. If you regenerate the secret, also
            // bump the addin version (existing customer keys will stop validating).
            byte[] a = Encoding.UTF8.GetBytes("HxZRnVKIvZcnsvBc5zbo4h3qh9T21Yry");
            byte[] b = Encoding.UTF8.GetBytes("n9J2EV8LsFiOkITbiayLbT3yu2pqC3Jh");
            byte[] c = Encoding.UTF8.GetBytes("QG6TFg25AzVs4RwtzNyi4W1hyZ2WhNyx");
            int n = Math.Max(a.Length, Math.Max(b.Length, c.Length));
            byte[] r = new byte[n];
            for (int i = 0; i < n; i++)
                r[i] = (byte)(a[i % a.Length] ^ b[i % b.Length] ^ c[i % c.Length]);
            return r;
        }

        public static string Base64UrlEncode(byte[] data)
        {
            return Convert.ToBase64String(data).TrimEnd('=').Replace('+', '-').Replace('/', '_');
        }

        public static byte[] Base64UrlDecode(string s)
        {
            s = s.Replace('-', '+').Replace('_', '/');
            switch (s.Length % 4)
            {
                case 2: s += "=="; break;
                case 3: s += "="; break;
            }
            return Convert.FromBase64String(s);
        }

        // RFC 4648 base32 alphabet.
        private static readonly char[] B32Alpha =
            "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567".ToCharArray();

        private static string Base32Encode(byte[] data, int bytes)
        {
            var sb = new StringBuilder();
            int buffer = 0;
            int bits = 0;
            for (int i = 0; i < bytes; i++)
            {
                buffer = (buffer << 8) | data[i];
                bits += 8;
                while (bits >= 5)
                {
                    bits -= 5;
                    sb.Append(B32Alpha[(buffer >> bits) & 0x1F]);
                }
            }
            if (bits > 0)
                sb.Append(B32Alpha[(buffer << (5 - bits)) & 0x1F]);
            return sb.ToString();
        }
    }
}
