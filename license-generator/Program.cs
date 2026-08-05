using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace PrintLayoutAddin.LicenseGenerator
{
    // PrintLayoutAddin license tool.
    //
    // !!! KEEP THIS TOOL PRIVATE. Anyone with this exe can generate license keys. !!!
    //
    //   PLLicenseGen                                  -> opens the GUI (issue many keys)
    //   PLLicenseGen --mid <id> --expire YYYY-MM-DD [--note "..."]   (CLI, 1 key)
    //   PLLicenseGen --batch input.csv [--out output.csv]            (CLI, many keys)
    //
    // The key crypto/format lives in LicenseCore and MUST match LicenseManager in the addin.
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            // No arguments -> friendly GUI. Arguments -> headless CLI (for scripting).
            if (args.Length == 0)
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainForm());
                return 0;
            }

            // Attach to the parent console so Console.WriteLine is visible when run from a shell.
            AttachConsole(ATTACH_PARENT_PROCESS);
            try
            {
                if (GetArg(args, "--batch") != null)
                    return RunBatch(args);
                return RunSingle(args);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Loi: " + ex.Message);
                return 1;
            }
        }

        // ---------- single-key CLI ----------

        private static int RunSingle(string[] args)
        {
            string machineId = GetArg(args, "--mid") ?? GetArg(args, "--machine-id");
            string expireStr = GetArg(args, "--expire");
            string note = GetArg(args, "--note") ?? "";

            if (string.IsNullOrWhiteSpace(machineId))
            {
                Console.Write("Machine ID (paste from customer's PLLICENSE dialog): ");
                machineId = (Console.ReadLine() ?? "").Trim();
            }
            if (string.IsNullOrWhiteSpace(machineId))
            {
                Console.Error.WriteLine("Machine ID is required.");
                return 2;
            }

            if (string.IsNullOrWhiteSpace(expireStr))
            {
                Console.Write("Ngay het han (YYYY-MM-DD): ");
                expireStr = (Console.ReadLine() ?? "").Trim();
            }
            DateTime expire;
            if (!LicenseCore.TryParseDate(expireStr, out expire))
            {
                Console.Error.WriteLine("Ngay het han khong hop le. Dinh dang YYYY-MM-DD, vd 2027-06-30.");
                return 2;
            }
            if (expire.Date < DateTime.UtcNow.Date)
                Console.Error.WriteLine("Canh bao: ngay het han nam trong qua khu (UTC).");

            if (string.IsNullOrWhiteSpace(note) && !HasArg(args, "--no-note"))
            {
                Console.Write("Ghi chu (ten KH / cong ty, Enter de bo qua): ");
                note = (Console.ReadLine() ?? "").Trim();
            }

            string payload;
            string key = LicenseCore.MakeKey(machineId, expire, note, out payload);

            Console.WriteLine();
            Console.WriteLine("=== LICENSE KEY ===");
            Console.WriteLine(key);
            Console.WriteLine();
            Console.WriteLine("Payload : " + payload);
            Console.WriteLine("Length  : " + key.Length + " chars");
            Console.WriteLine();
            Console.WriteLine("Gui ca chuoi tren cho khach. Khach mo AutoCAD, go 'PLLICENSE', dan key vao, bam Kich hoat.");

            LicenseCore.LogIssued(machineId, expire, note, key);
            Console.WriteLine("Log    : " + LicenseCore.LogPath());
            return 0;
        }

        // ---------- batch CLI (CSV in -> CSV out) ----------

        private static int RunBatch(string[] args)
        {
            string inPath = GetArg(args, "--batch");
            string outPath = GetArg(args, "--out");
            string defaultExpire = GetArg(args, "--expire");
            string defaultNote = GetArg(args, "--note");

            if (!File.Exists(inPath))
            {
                Console.Error.WriteLine("Khong tim thay file: " + inPath);
                return 2;
            }
            if (string.IsNullOrWhiteSpace(outPath))
            {
                string dir = Path.GetDirectoryName(Path.GetFullPath(inPath));
                outPath = Path.Combine(dir, Path.GetFileNameWithoutExtension(inPath) + "_keys.csv");
            }

            var rows = LicenseCore.ParseTable(File.ReadAllText(inPath, Encoding.UTF8));
            if (rows.Count < 2)
            {
                Console.Error.WriteLine("File khong co du lieu (can dong tieu de + it nhat 1 dong).");
                return 2;
            }

            string[] header = rows[0];
            int ciMid = LicenseCore.FindColumn(header, LicenseCore.MidAliases);
            int ciExp = LicenseCore.FindColumn(header, LicenseCore.ExpAliases);
            int ciUser = LicenseCore.FindColumn(header, LicenseCore.UserAliases);
            int ciNote = LicenseCore.FindColumn(header, LicenseCore.NoteAliases);

            if (ciMid < 0 || ciExp < 0)
            {
                Console.Error.WriteLine("Khong nhan dien duoc cot machine_id va/hoac expire o dong tieu de.");
                Console.Error.WriteLine("Tieu de doc duoc: " + string.Join(" | ", header));
                Console.Error.WriteLine("Hay dat ten cot (dong dau tien): machine_id, expire, note");
                return 2;
            }

            var outLines = new List<string>();
            var outHeader = header.ToList();
            outHeader.Add("license_key");
            outHeader.Add("status");
            outLines.Add(string.Join(",", outHeader.Select(LicenseCore.CsvEscape)));

            int ok = 0, fail = 0;
            for (int r = 1; r < rows.Count; r++)
            {
                string[] row = rows[r];
                string mid = Cell(row, ciMid);
                string expStr = Cell(row, ciExp);
                string note = LicenseCore.CombineNote(
                    ciUser >= 0 ? Cell(row, ciUser) : "",
                    ciNote >= 0 ? Cell(row, ciNote) : "");

                if (string.IsNullOrWhiteSpace(expStr) && !string.IsNullOrWhiteSpace(defaultExpire))
                    expStr = defaultExpire.Trim();
                if (string.IsNullOrWhiteSpace(note) && !string.IsNullOrWhiteSpace(defaultNote))
                    note = defaultNote;

                if (string.IsNullOrWhiteSpace(mid) && string.IsNullOrWhiteSpace(expStr))
                    continue;

                var outRow = row.ToList();
                while (outRow.Count < header.Length) outRow.Add("");

                string key = null;
                string status;
                DateTime exp;
                if (string.IsNullOrWhiteSpace(mid))
                {
                    status = "ERROR: thieu machine_id";
                    fail++;
                }
                else if (!LicenseCore.TryParseDate(expStr, out exp))
                {
                    status = "ERROR: ngay het han khong hop le (can YYYY-MM-DD): '" + expStr + "'";
                    fail++;
                }
                else
                {
                    string payload;
                    key = LicenseCore.MakeKey(mid, exp, note, out payload);
                    LicenseCore.LogIssued(mid, exp, note, key);
                    status = exp.Date < DateTime.UtcNow.Date ? "OK (canh bao: da het han)" : "OK";
                    ok++;
                }

                outRow.Add(key ?? "");
                outRow.Add(status);
                outLines.Add(string.Join(",", outRow.Select(LicenseCore.CsvEscape)));
                Console.WriteLine(string.Format("[{0,3}] {1,-22} {2,-12} -> {3}",
                    r, mid, expStr, key != null ? "OK" : status));
            }

            File.WriteAllText(outPath, string.Join("\r\n", outLines) + "\r\n", new UTF8Encoding(true));

            Console.WriteLine();
            Console.WriteLine("Da cap " + ok + " key, loi " + fail + ".");
            Console.WriteLine("Ket qua: " + Path.GetFullPath(outPath));
            Console.WriteLine("Log    : " + LicenseCore.LogPath());
            return fail > 0 ? 3 : 0;
        }

        // ---------- helpers ----------

        private static string Cell(string[] row, int idx)
        {
            return (idx >= 0 && idx < row.Length) ? (row[idx] ?? "").Trim() : "";
        }

        private static string GetArg(string[] args, string name)
        {
            for (int i = 0; i < args.Length - 1; i++)
                if (string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase))
                    return args[i + 1];
            return null;
        }

        private static bool HasArg(string[] args, string name)
        {
            foreach (var a in args)
                if (string.Equals(a, name, StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }

        [DllImport("kernel32.dll")]
        private static extern bool AttachConsole(int dwProcessId);
        private const int ATTACH_PARENT_PROCESS = -1;
    }
}
