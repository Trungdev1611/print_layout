using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Registry = Microsoft.Win32.Registry;
using RegistryKey = Microsoft.Win32.RegistryKey;
using RegistryValueKind = Microsoft.Win32.RegistryValueKind;

namespace PrintLayoutAddin.Core
{
    /// <summary>
    /// User-configurable command shortcuts for PrintLayoutAddin:
    ///   * Hotkeys - Ctrl/Alt key combos handled by a WinForms message filter (instant, no reload).
    ///   * Aliases - typed command abbreviations written into acad.pgp (e.g. 11 -> PLPRINT).
    /// Both are persisted in the registry; aliases are mirrored into acad.pgp and the PGP section
    /// is reloaded in-session via the RE-INIT system variable (bit 16).
    /// </summary>
    public static class ShortcutManager
    {
        private const string RegistryPath = @"Software\PrintLayoutAddin\Shortcuts";
        // Keep these byte-for-byte stable across versions so an older managed block in acad.pgp is
        // still recognized and replaced/removed (otherwise old aliases would be orphaned).
        private const string BlockStart = "; >>> PrintLayoutAddin shortcuts (auto-generated; edit via PLKEYS) >>>";
        private const string BlockEnd = "; <<< PrintLayoutAddin shortcuts <<<";
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_SYSKEYDOWN = 0x0104;

        public sealed class ShortcutDef
        {
            public ShortcutDef(string commandName, string label, Keys defaultKeys)
            {
                CommandName = commandName;
                Label = label;
                DefaultKeys = defaultKeys;
            }

            public string CommandName { get; }
            public string Label { get; }
            public Keys DefaultKeys { get; }
        }

        public sealed class ShortcutConfig
        {
            public Dictionary<string, Keys> Hotkeys { get; } =
                new Dictionary<string, Keys>(StringComparer.OrdinalIgnoreCase);

            public Dictionary<string, string> Aliases { get; } =
                new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        // Default hotkeys are None - the user opts in via the dialog so we never hijack a key.
        private static readonly ShortcutDef[] Defs =
        {
            new ShortcutDef("PLAUTO",    "Auto Frames",       Keys.None),
            new ShortcutDef("PLSTT",     "Number Frames",     Keys.None),
            new ShortcutDef("PLAYOUT",   "Build Layouts",     Keys.None),
            new ShortcutDef("PLPRINT",   "Print / Export PDF",Keys.None),
            new ShortcutDef("PLCLEAN",   "Cleanup Frames",    Keys.None),
            new ShortcutDef("PLVP",      "Reset Corners",     Keys.None),
            new ShortcutDef("PLLICENSE", "License",           Keys.None),
        };

        private static readonly Dictionary<string, Keys> Hotkeys =
            new Dictionary<string, Keys>(StringComparer.OrdinalIgnoreCase);

        private static readonly Dictionary<string, string> Aliases =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        private static ShortcutMessageFilter _filter;
        private static bool _suspended;

        public static IReadOnlyList<ShortcutDef> Definitions => Defs;

        public static void Install()
        {
            Load();
            if (_filter == null)
            {
                _filter = new ShortcutMessageFilter();
                System.Windows.Forms.Application.AddMessageFilter(_filter);
            }
        }

        public static void Uninstall()
        {
            if (_filter != null)
            {
                System.Windows.Forms.Application.RemoveMessageFilter(_filter);
                _filter = null;
            }
        }

        /// <summary>Pause hotkey interception (e.g. while the editor dialog is open).</summary>
        public static void Suspend(bool value) => _suspended = value;

        /// <summary>Re-read settings from the (shared) registry before showing the editor.</summary>
        public static void Reload() => Load();

        public static ShortcutConfig CurrentConfig()
        {
            var config = new ShortcutConfig();
            foreach (var def in Defs)
            {
                config.Hotkeys[def.CommandName] = Hotkeys.TryGetValue(def.CommandName, out var keys) ? keys : Keys.None;
                config.Aliases[def.CommandName] = Aliases.TryGetValue(def.CommandName, out var alias) ? alias : string.Empty;
            }
            return config;
        }

        public static ShortcutConfig DefaultConfig()
        {
            var config = new ShortcutConfig();
            foreach (var def in Defs)
            {
                config.Hotkeys[def.CommandName] = def.DefaultKeys;
                config.Aliases[def.CommandName] = string.Empty;
            }
            return config;
        }

        public static void Apply(ShortcutConfig config)
        {
            if (config == null) return;

            Hotkeys.Clear();
            Aliases.Clear();
            foreach (var def in Defs)
            {
                Hotkeys[def.CommandName] = config.Hotkeys.TryGetValue(def.CommandName, out var keys) ? keys : Keys.None;
                string alias = config.Aliases.TryGetValue(def.CommandName, out var value) ? value : string.Empty;
                Aliases[def.CommandName] = (alias ?? string.Empty).Trim().ToUpperInvariant();
            }

            Save();
            ApplyAliasesToPgp();
        }

        public static bool IsValidAlias(string alias)
        {
            if (string.IsNullOrWhiteSpace(alias)) return false;
            foreach (char c in alias)
                if (!char.IsLetterOrDigit(c)) return false;
            return true;
        }

        public static string Describe(Keys keys)
        {
            if (keys == Keys.None) return "(none)";
            var parts = new List<string>();
            if ((keys & Keys.Control) != 0) parts.Add("Ctrl");
            if ((keys & Keys.Alt) != 0) parts.Add("Alt");
            if ((keys & Keys.Shift) != 0) parts.Add("Shift");
            Keys code = keys & Keys.KeyCode;
            if (code != Keys.None) parts.Add(KeyName(code));
            return string.Join("+", parts);
        }

        /// <summary>Alias names defined in acad.pgp outside our managed block.</summary>
        public static HashSet<string> ReadForeignAliases()
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                string pgpPath = FindPgpPath();
                if (string.IsNullOrEmpty(pgpPath) || !File.Exists(pgpPath)) return set;
                var lines = ReadPgpWithRetry(pgpPath);
                if (lines == null) return set;

                bool inBlock = false;
                foreach (var raw in lines)
                {
                    string line = raw.Trim();
                    if (line.Equals(BlockStart, StringComparison.OrdinalIgnoreCase)) { inBlock = true; continue; }
                    if (line.Equals(BlockEnd, StringComparison.OrdinalIgnoreCase)) { inBlock = false; continue; }
                    if (inBlock || line.Length == 0 || line.StartsWith(";", StringComparison.Ordinal)) continue;

                    int comma = line.IndexOf(',');
                    if (comma <= 0) continue;
                    string name = line.Substring(0, comma).Trim();
                    if (name.Length > 0) set.Add(name);
                }
            }
            catch { }
            return set;
        }

        private static string KeyName(Keys code)
        {
            if (code >= Keys.D0 && code <= Keys.D9)
                return ((char)('0' + (code - Keys.D0))).ToString(CultureInfo.InvariantCulture);
            if (code >= Keys.NumPad0 && code <= Keys.NumPad9)
                return "Num" + (code - Keys.NumPad0).ToString(CultureInfo.InvariantCulture);
            return code.ToString();
        }

        private static void Load()
        {
            Hotkeys.Clear();
            Aliases.Clear();
            foreach (var def in Defs)
            {
                Hotkeys[def.CommandName] = def.DefaultKeys;
                Aliases[def.CommandName] = string.Empty;
            }

            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryPath))
                {
                    if (key == null) return;
                    foreach (var def in Defs)
                    {
                        object hotkeyValue = key.GetValue("Shortcut_" + def.CommandName);
                        if (hotkeyValue != null && int.TryParse(
                                Convert.ToString(hotkeyValue, CultureInfo.InvariantCulture),
                                NumberStyles.Integer, CultureInfo.InvariantCulture, out int raw))
                        {
                            Hotkeys[def.CommandName] = (Keys)raw;
                        }

                        object aliasValue = key.GetValue("Alias_" + def.CommandName);
                        if (aliasValue != null)
                            Aliases[def.CommandName] =
                                (Convert.ToString(aliasValue, CultureInfo.InvariantCulture) ?? string.Empty).Trim().ToUpperInvariant();
                    }
                }
            }
            catch { }
        }

        private static void Save()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegistryPath))
                {
                    if (key == null) return;
                    foreach (var def in Defs)
                    {
                        Keys keys = Hotkeys.TryGetValue(def.CommandName, out var value) ? value : Keys.None;
                        key.SetValue("Shortcut_" + def.CommandName, (int)keys, RegistryValueKind.DWord);
                        string alias = Aliases.TryGetValue(def.CommandName, out var aliasValue) ? aliasValue : string.Empty;
                        key.SetValue("Alias_" + def.CommandName, alias ?? string.Empty, RegistryValueKind.String);
                    }
                }
            }
            catch { }
        }

        private static void ApplyAliasesToPgp()
        {
            try
            {
                string pgpPath = FindPgpPath();
                if (string.IsNullOrEmpty(pgpPath) || !File.Exists(pgpPath))
                {
                    Write("[KEYS] acad.pgp not found; typed aliases were not applied (hotkeys still work).");
                    return;
                }

                var lines = ReadPgpWithRetry(pgpPath);
                if (lines == null)
                {
                    Write("[KEYS] acad.pgp is in use; aliases not saved. Try again. Hotkeys are unaffected.");
                    return;
                }

                StripManagedBlock(lines);

                var block = BuildManagedBlock();
                if (block.Count > 0)
                {
                    if (lines.Count > 0 && lines[lines.Count - 1].Trim().Length != 0) lines.Add(string.Empty);
                    lines.AddRange(block);
                }

                if (!WritePgpWithRetry(pgpPath, lines)) return;

                // Reload the alias (PGP) section silently. Bit 16 = PGP file. The REINIT *command*
                // would pop up the interactive dialog instead; the AutoLISP (reinit) function is not
                // available on every install. Setting the RE-INIT system variable always works.
                try { AcadApp.SetSystemVariable("RE-INIT", 16); } catch { }
                Write("[KEYS] Updated acad.pgp aliases and reloaded.");
            }
            catch (System.Exception ex)
            {
                Write("[KEYS] Cannot update acad.pgp: " + ex.Message);
            }
        }

        private static List<string> ReadPgpWithRetry(string path)
        {
            for (int attempt = 0; attempt < 4; attempt++)
            {
                try { return new List<string>(File.ReadAllLines(path)); }
                catch (IOException) { System.Threading.Thread.Sleep(150); }
            }
            return null;
        }

        private static bool WritePgpWithRetry(string path, List<string> lines)
        {
            try
            {
                FileAttributes attributes = File.GetAttributes(path);
                if ((attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                    File.SetAttributes(path, attributes & ~FileAttributes.ReadOnly);
            }
            catch (System.Exception ex)
            {
                Write("[KEYS] acad.pgp is read-only and cannot be changed (" + ex.Message + "). Hotkeys still work.");
                return false;
            }

            string tempPath = path + ".pla.tmp";
            for (int attempt = 0; attempt < 4; attempt++)
            {
                try
                {
                    File.WriteAllLines(tempPath, lines);
                    // Temp-then-swap so a concurrent session never reads a half-written acad.pgp.
                    if (File.Exists(path)) File.Replace(tempPath, path, null);
                    else File.Move(tempPath, path);
                    return true;
                }
                catch (UnauthorizedAccessException ex)
                {
                    CleanupTemp(tempPath);
                    Write("[KEYS] No permission to write acad.pgp (" + ex.Message + "). Hotkeys still work.");
                    return false;
                }
                catch (IOException)
                {
                    CleanupTemp(tempPath);
                    if (attempt == 3)
                    {
                        Write("[KEYS] acad.pgp is locked by another AutoCAD session; aliases not saved. Hotkeys are unaffected.");
                        return false;
                    }
                    System.Threading.Thread.Sleep(200);
                }
            }
            return false;
        }

        private static void CleanupTemp(string tempPath)
        {
            try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { }
        }

        private static List<string> BuildManagedBlock()
        {
            var entries = new List<string>();
            foreach (var def in Defs)
            {
                if (Aliases.TryGetValue(def.CommandName, out var alias) && !string.IsNullOrWhiteSpace(alias))
                    entries.Add(alias.Trim().ToUpperInvariant() + ",*" + def.CommandName);
            }

            var block = new List<string>();
            if (entries.Count == 0) return block;
            block.Add(BlockStart);
            block.AddRange(entries);
            block.Add(BlockEnd);
            return block;
        }

        private static void StripManagedBlock(List<string> lines)
        {
            int start = lines.FindIndex(l => l.Trim().Equals(BlockStart, StringComparison.OrdinalIgnoreCase));
            if (start < 0) return;
            int end = lines.FindIndex(start, l => l.Trim().Equals(BlockEnd, StringComparison.OrdinalIgnoreCase));
            if (end < 0) end = lines.Count - 1;
            int removeFrom = start;
            if (removeFrom > 0 && lines[removeFrom - 1].Trim().Length == 0) removeFrom--;
            lines.RemoveRange(removeFrom, end - removeFrom + 1);
        }

        private static string FindPgpPath()
        {
            try
            {
                var doc = AcadApp.DocumentManager.MdiActiveDocument;
                Database db = doc?.Database ?? HostApplicationServices.WorkingDatabase;
                return HostApplicationServices.Current.FindFile("acad.pgp", db, FindFileHint.Default);
            }
            catch { return null; }
        }

        private static void Write(string message)
        {
            try { AcadApp.DocumentManager.MdiActiveDocument?.Editor.WriteMessage("\n" + message); } catch { }
        }

        // ---------- Hotkey interception ----------

        private static string FindCommand(Keys combo)
        {
            foreach (var pair in Hotkeys)
                if (pair.Value != Keys.None && pair.Value == combo) return pair.Key;
            return null;
        }

        private static void RunCommand(string commandName)
        {
            try
            {
                var doc = AcadApp.DocumentManager.MdiActiveDocument;
                doc?.SendStringToExecute(commandName + " ", true, false, true);
            }
            catch { }
        }

        private static bool HandleKeyMessage(ref Message m)
        {
            if (_suspended) return false;
            if (m.Msg != WM_KEYDOWN && m.Msg != WM_SYSKEYDOWN) return false;

            Keys code = (Keys)((long)m.WParam & 0xFFFF);
            if (code == Keys.None || code == Keys.ControlKey || code == Keys.ShiftKey || code == Keys.Menu)
                return false;

            Keys modifiers = Control.ModifierKeys;
            // Require Ctrl or Alt so plain typing is never hijacked.
            if ((modifiers & (Keys.Control | Keys.Alt)) == 0) return false;

            string command = FindCommand(code | modifiers);
            if (command == null) return false;

            RunCommand(command);
            return true;
        }

        private sealed class ShortcutMessageFilter : IMessageFilter
        {
            public bool PreFilterMessage(ref Message m) => HandleKeyMessage(ref m);
        }
    }
}
