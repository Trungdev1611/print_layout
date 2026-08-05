using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using PrintLayoutAddin.Core;

namespace PrintLayoutAddin.UI
{
    public class ShortcutsDialog : Form
    {
        private readonly ShortcutManager.ShortcutConfig _defaults;
        private readonly Dictionary<string, Keys> _hotkeys = new Dictionary<string, Keys>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, TextBox> _hotkeyBoxes = new Dictionary<string, TextBox>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, TextBox> _aliasBoxes = new Dictionary<string, TextBox>(StringComparer.OrdinalIgnoreCase);

        public ShortcutManager.ShortcutConfig Result { get; private set; }

        public ShortcutsDialog(ShortcutManager.ShortcutConfig current, ShortcutManager.ShortcutConfig defaults)
        {
            _defaults = defaults;

            Text = "Keyboard Shortcuts & Aliases";
            Width = 580;
            Height = 430;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MinimizeBox = false;
            MaximizeBox = false;

            BuildUi(current ?? new ShortcutManager.ShortcutConfig());
        }

        private void BuildUi(ShortcutManager.ShortcutConfig current)
        {
            var grid = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                Padding = new Padding(12, 12, 12, 0)
            };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));

            grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            grid.Controls.Add(Header("Command"), 0, 0);
            grid.Controls.Add(Header("Hotkey"), 1, 0);
            grid.Controls.Add(Header(""), 2, 0);
            grid.Controls.Add(Header("Type alias"), 3, 0);

            int row = 1;
            foreach (var def in ShortcutManager.Definitions)
            {
                string command = def.CommandName;
                _hotkeys[command] = current.Hotkeys.TryGetValue(command, out var k) ? k : Keys.None;
                string alias = current.Aliases.TryGetValue(command, out var a) ? (a ?? "") : "";

                grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));

                grid.Controls.Add(new Label
                {
                    Text = $"{def.Label} ({command})",
                    Dock = DockStyle.Fill,
                    TextAlign = ContentAlignment.MiddleLeft
                }, 0, row);

                var hotkeyBox = new TextBox
                {
                    Dock = DockStyle.Fill,
                    ReadOnly = true,
                    Cursor = Cursors.Hand,
                    TextAlign = HorizontalAlignment.Center,
                    Text = ShortcutManager.Describe(_hotkeys[command])
                };
                hotkeyBox.GotFocus += (s, e) => hotkeyBox.Text = "Press a combo...";
                hotkeyBox.LostFocus += (s, e) => hotkeyBox.Text = ShortcutManager.Describe(_hotkeys[command]);
                hotkeyBox.KeyDown += (s, e) => CaptureHotkey(command, hotkeyBox, e);
                _hotkeyBoxes[command] = hotkeyBox;
                grid.Controls.Add(hotkeyBox, 1, row);

                var clearButton = new Button { Text = "Clear", Dock = DockStyle.Fill, Height = 26 };
                clearButton.Click += (s, e) =>
                {
                    _hotkeys[command] = Keys.None;
                    hotkeyBox.Text = ShortcutManager.Describe(Keys.None);
                };
                grid.Controls.Add(clearButton, 2, row);

                var aliasBox = new TextBox
                {
                    Dock = DockStyle.Fill,
                    CharacterCasing = CharacterCasing.Upper,
                    Text = alias
                };
                _aliasBoxes[command] = aliasBox;
                grid.Controls.Add(aliasBox, 3, row);

                row++;
            }

            var hint = new Label
            {
                Text = "Hotkey: click the box and press a combo (Ctrl or Alt required), Esc/Delete to clear.\n" +
                       "Type alias: enter a short code (e.g. 11) then press Enter at the command line to run. Blank = none.",
                Dock = DockStyle.Fill,
                AutoSize = false
            };
            grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 56));
            grid.Controls.Add(hint, 0, row);
            grid.SetColumnSpan(hint, 4);

            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft,
                Height = 46,
                Padding = new Padding(12, 8, 12, 8)
            };
            var ok = new Button { Text = "OK", Width = 88, Height = 28 };
            var cancel = new Button { Text = "Cancel", Width = 88, Height = 28, DialogResult = DialogResult.Cancel };
            var reset = new Button { Text = "Reset", Width = 88, Height = 28 };
            ok.Click += (s, e) => TryAccept();
            reset.Click += (s, e) => ResetToDefaults();
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(ok);
            buttons.Controls.Add(reset);

            AcceptButton = ok;
            CancelButton = cancel;

            Controls.Add(grid);
            Controls.Add(buttons);
        }

        private static Label Header(string text) => new Label
        {
            Text = text,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            Font = new System.Drawing.Font(System.Drawing.SystemFonts.DefaultFont, FontStyle.Bold)
        };

        private void CaptureHotkey(string command, TextBox box, KeyEventArgs e)
        {
            Keys keyCode = e.KeyCode;
            if (keyCode == Keys.ControlKey || keyCode == Keys.ShiftKey || keyCode == Keys.Menu || keyCode == Keys.Tab)
                return;

            e.SuppressKeyPress = true;
            e.Handled = true;

            if (keyCode == Keys.Escape || keyCode == Keys.Back || keyCode == Keys.Delete)
            {
                _hotkeys[command] = Keys.None;
                box.Text = ShortcutManager.Describe(Keys.None);
                return;
            }

            Keys modifiers = e.Modifiers;
            if ((modifiers & (Keys.Control | Keys.Alt)) == 0)
            {
                box.Text = "Ctrl or Alt required!";
                return;
            }

            _hotkeys[command] = keyCode | modifiers;
            box.Text = ShortcutManager.Describe(_hotkeys[command]);
        }

        private void ResetToDefaults()
        {
            foreach (var def in ShortcutManager.Definitions)
            {
                Keys keyValue = _defaults != null && _defaults.Hotkeys.TryGetValue(def.CommandName, out var dk) ? dk : Keys.None;
                string aliasValue = _defaults != null && _defaults.Aliases.TryGetValue(def.CommandName, out var da) ? (da ?? "") : "";
                _hotkeys[def.CommandName] = keyValue;
                _hotkeyBoxes[def.CommandName].Text = ShortcutManager.Describe(keyValue);
                _aliasBoxes[def.CommandName].Text = aliasValue;
            }
        }

        private void TryAccept()
        {
            var config = new ShortcutManager.ShortcutConfig();
            foreach (var def in ShortcutManager.Definitions)
            {
                config.Hotkeys[def.CommandName] = _hotkeys.TryGetValue(def.CommandName, out var k) ? k : Keys.None;
                config.Aliases[def.CommandName] = _aliasBoxes[def.CommandName].Text.Trim().ToUpperInvariant();
            }

            // No duplicate hotkeys.
            var seenHotkeys = new Dictionary<Keys, string>();
            foreach (var pair in config.Hotkeys)
            {
                if (pair.Value == Keys.None) continue;
                if (seenHotkeys.ContainsKey(pair.Value))
                {
                    Warn($"Hotkey {ShortcutManager.Describe(pair.Value)} is assigned to more than one command.");
                    return;
                }
                seenHotkeys[pair.Value] = pair.Key;
            }

            // Valid + unique aliases.
            var seenAliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var pair in config.Aliases)
            {
                string alias = pair.Value;
                if (string.IsNullOrEmpty(alias)) continue;
                if (!ShortcutManager.IsValidAlias(alias))
                {
                    Warn($"Alias \"{alias}\" is invalid. Use letters and digits only (e.g. 11, PP, AR1).");
                    return;
                }
                if (seenAliases.ContainsKey(alias))
                {
                    Warn($"Alias \"{alias}\" is assigned to more than one command.");
                    return;
                }
                seenAliases[alias] = pair.Key;
            }

            // Warn (but allow) when a typed alias overrides an existing acad.pgp alias.
            var foreign = ShortcutManager.ReadForeignAliases();
            foreach (var alias in seenAliases.Keys)
            {
                if (foreign.Contains(alias))
                {
                    var choice = MessageBox.Show(this,
                        $"Alias \"{alias}\" is already used by another AutoCAD command. Override it with the PrintLayout command?",
                        "Keyboard Shortcuts", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (choice != DialogResult.Yes)
                    {
                        DialogResult = DialogResult.None;
                        return;
                    }
                }
            }

            Result = config;
            DialogResult = DialogResult.OK;
            Close();
        }

        private void Warn(string msg)
        {
            MessageBox.Show(this, msg, "Keyboard Shortcuts", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            DialogResult = DialogResult.None;
        }
    }
}
