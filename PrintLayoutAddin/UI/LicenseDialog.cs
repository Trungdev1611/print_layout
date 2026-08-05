using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using PrintLayoutAddin.Core;

namespace PrintLayoutAddin.UI
{
    public class LicenseDialog : Form
    {
        // --- Version / build identity (shown to the user for support) ---
        // Version comes from the assembly (set via <Version> in the .csproj).
        private static readonly string AppVersion = ReadAppVersion();

        // Which runtime/build is actually loaded — helps diagnose the AutoCAD edition.
#if NET8_0_OR_GREATER
        private const string RuntimeLabel = ".NET 8 build (AutoCAD 2025+)";
#else
        private const string RuntimeLabel = ".NET Framework build (AutoCAD 2018-2024)";
#endif

        private static string ReadAppVersion()
        {
            var v = Assembly.GetExecutingAssembly().GetName().Version;
            return v != null ? $"{v.Major}.{v.Minor}.{v.Build}" : "1.0.0";
        }

        // --- Layout constants (single source of truth for spacing) ---
        // NOTE: actual pixel sizes are scaled at runtime by WinForms font/DPI
        // auto-scaling. The layout uses TableLayoutPanel + AutoSize controls so
        // nothing clips on high-DPI laptops (125%/150% scaling).
        private const int FormWidth = 620;
        private const int Pad = 20;
        private const int SectionGap = 18;
        private const int ContentWidth = FormWidth - Pad * 2;

        // --- Colors ---
        private static readonly Color AccentBlue = Color.FromArgb(0, 120, 215);
        private static readonly Color SuccessGreen = Color.FromArgb(30, 138, 60);
        private static readonly Color WarnAmber = Color.FromArgb(196, 130, 0);
        private static readonly Color ErrorRed = Color.FromArgb(196, 60, 60);
        private static readonly Color NeutralGray = Color.FromArgb(100, 100, 100);
        private static readonly Color StatusBgValid = Color.FromArgb(232, 245, 233);
        private static readonly Color StatusBgWarn = Color.FromArgb(255, 248, 225);
        private static readonly Color StatusBgError = Color.FromArgb(253, 232, 232);
        private static readonly Color StatusBgNeutral = Color.FromArgb(245, 247, 250);
        private static readonly Color BorderColor = Color.FromArgb(218, 220, 224);
        private static readonly Color HintGray = Color.FromArgb(96, 102, 110);

        // --- Controls ---
        private Panel _statusPanel;
        private Label _statusTitle;
        private Label _statusDetail;
        private TextBox _txtMachineId;
        private TextBox _txtKey;
        private Button _btnCopy;
        private Button _btnActivate;
        private Button _btnClose;

        private System.Windows.Forms.Timer _copyResetTimer;

        public LicenseDialog()
        {
            BuildUI();
            RefreshStatus(LicenseManager.Current);
        }

        // ---------------- UI construction ----------------

        private void BuildUI()
        {
            Text = $"PrintLayoutAddin v{AppVersion} — License";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            Font = new Font("Segoe UI", 9F);
            BackColor = Color.White;

            // DPI-robust: let the form size itself to the content and let WinForms
            // font-scale every child. Combined with AutoSize labels this guarantees
            // text never clips on high-DPI laptops.
            AutoScaleMode = AutoScaleMode.Font;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 1,
                Padding = new Padding(Pad),
                BackColor = Color.White,
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, ContentWidth));

            int rowAt = 0;
            void AddRow(Control c)
            {
                root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                root.Controls.Add(c, 0, rowAt++);
            }

            // ----- Status banner -----
            _statusPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = StatusBgNeutral,
                BorderStyle = BorderStyle.None,
                Padding = new Padding(14, 10, 14, 12),
                Margin = new Padding(0, 0, 0, SectionGap),
            };
            _statusPanel.Paint += DrawPanelBorder;

            var statusInner = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 1,
                Margin = new Padding(0),
                BackColor = Color.Transparent,
            };
            statusInner.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, ContentWidth - 28));
            _statusTitle = new Label
            {
                AutoSize = true,
                MaximumSize = new Size(ContentWidth - 28, 0),
                Font = new Font("Segoe UI Semibold", 11F, FontStyle.Bold),
                ForeColor = NeutralGray,
                Margin = new Padding(0, 0, 0, 4),
                Text = "License status"
            };
            _statusDetail = new Label
            {
                AutoSize = true,
                MaximumSize = new Size(ContentWidth - 28, 0),
                Font = new Font("Segoe UI", 9F),
                ForeColor = NeutralGray,
                Margin = new Padding(0),
                Text = ""
            };
            statusInner.Controls.Add(_statusTitle, 0, 0);
            statusInner.Controls.Add(_statusDetail, 0, 1);
            _statusPanel.Controls.Add(statusInner);
            AddRow(_statusPanel);

            // ----- Machine ID section -----
            AddRow(new Label
            {
                Text = "Your Machine ID",
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 4),
                Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold)
            });

            var midRow = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                Margin = new Padding(0, 0, 0, SectionGap),
            };
            midRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            midRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _txtMachineId = new TextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Font = new Font("Consolas", 10F),
                BackColor = Color.FromArgb(250, 251, 252),
                TabStop = false,
                Margin = new Padding(0, 0, 8, 0),
            };
            _btnCopy = new Button
            {
                Text = "Copy",
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                MinimumSize = new Size(96, 28),
                FlatStyle = FlatStyle.System,
                Margin = new Padding(0),
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                TabIndex = 0
            };
            _btnCopy.Click += OnCopyMachineId;
            midRow.Controls.Add(_txtMachineId, 0, 0);
            midRow.Controls.Add(_btnCopy, 1, 0);
            AddRow(midRow);

            // ----- License key section -----
            AddRow(new Label
            {
                Text = "License Key",
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 4),
                Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold)
            });

            _txtKey = new TextBox
            {
                Width = ContentWidth,
                Height = 76,
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Consolas", 9F),
                WordWrap = true,
                AcceptsReturn = false,
                AcceptsTab = false,
                Margin = new Padding(0, 0, 0, 10),
                TabIndex = 1
            };
            AddRow(_txtKey);

            // ----- Activate button -----
            _btnActivate = new Button
            {
                Text = "Activate",
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                MinimumSize = new Size(130, 34),
                Padding = new Padding(10, 4, 10, 4),
                Font = new Font("Segoe UI Semibold", 9.5F, FontStyle.Bold),
                FlatStyle = FlatStyle.Flat,
                BackColor = AccentBlue,
                ForeColor = Color.White,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 0, 0, SectionGap),
                TabIndex = 2,
                Cursor = Cursors.Hand
            };
            _btnActivate.FlatAppearance.BorderSize = 0;
            _btnActivate.FlatAppearance.MouseOverBackColor = Color.FromArgb(16, 110, 190);
            _btnActivate.FlatAppearance.MouseDownBackColor = Color.FromArgb(0, 100, 180);
            _btnActivate.Click += OnActivate;
            AddRow(_btnActivate);

            // ----- Hint text -----
            AddRow(new Label
            {
                AutoSize = true,
                MaximumSize = new Size(ContentWidth, 0),
                Font = new Font("Segoe UI", 8.5F, FontStyle.Italic),
                ForeColor = HintGray,
                Margin = new Padding(0, 0, 0, 12),
                Text = "Send your Machine ID above to your supplier to receive a license key. " +
                       "The key is bound to this machine and cannot be used on another computer."
            });

            // ----- Version / build footer -----
            AddRow(new Label
            {
                AutoSize = true,
                MaximumSize = new Size(ContentWidth, 0),
                Font = new Font("Segoe UI", 8.5F),
                ForeColor = NeutralGray,
                Margin = new Padding(0, 0, 0, 8),
                Text = $"PrintLayoutAddin  v{AppVersion}   •   {RuntimeLabel}"
            });

            // ----- Close button (bottom-right) -----
            _btnClose = new Button
            {
                Text = "Close",
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                MinimumSize = new Size(110, 32),
                FlatStyle = FlatStyle.System,
                Anchor = AnchorStyles.Right,
                Margin = new Padding(0),
                DialogResult = DialogResult.OK,
                TabIndex = 3
            };
            _btnClose.Click += (s, e) => { DialogResult = DialogResult.OK; Close(); };
            AddRow(_btnClose);

            Controls.Add(root);

            AcceptButton = _btnActivate;
            CancelButton = _btnClose;
        }

        private static void DrawPanelBorder(object sender, PaintEventArgs e)
        {
            var p = (Panel)sender;
            using (var pen = new Pen(BorderColor))
                e.Graphics.DrawRectangle(pen, 0, 0, p.Width - 1, p.Height - 1);
        }

        // ---------------- Event handlers ----------------

        private void OnCopyMachineId(object sender, EventArgs e)
        {
            try
            {
                Clipboard.SetText(_txtMachineId.Text);
                _btnCopy.Text = "Copied";
            }
            catch
            {
                _btnCopy.Text = "Copy failed";
            }
            ResetCopyButtonAfterDelay();
        }

        private void ResetCopyButtonAfterDelay()
        {
            if (_copyResetTimer != null)
            {
                _copyResetTimer.Stop();
                _copyResetTimer.Dispose();
            }
            _copyResetTimer = new System.Windows.Forms.Timer { Interval = 1500 };
            _copyResetTimer.Tick += (s, e) =>
            {
                _btnCopy.Text = "Copy";
                _copyResetTimer.Stop();
                _copyResetTimer.Dispose();
                _copyResetTimer = null;
            };
            _copyResetTimer.Start();
        }

        private void OnActivate(object sender, EventArgs e)
        {
            var key = _txtKey.Text;
            if (string.IsNullOrWhiteSpace(key))
            {
                MessageBox.Show(this,
                    "Please paste a license key before clicking Activate.",
                    "License",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                _txtKey.Focus();
                return;
            }

            var info = LicenseManager.Activate(key);
            RefreshStatus(info);

            if (info.IsValid)
            {
                _txtKey.ReadOnly = true;
                _btnActivate.Enabled = false;
                _btnActivate.Text = "Activated";
            }
        }

        // ---------------- Status rendering ----------------

        private void RefreshStatus(LicenseInfo info)
        {
            _txtMachineId.Text = info.MachineId ?? "";

            switch (info.State)
            {
                case LicenseState.Valid:
                    SetStatus(
                        title: "License active",
                        detail: BuildValidDetail(info),
                        bg: StatusBgValid,
                        fg: SuccessGreen);
                    _btnActivate.Enabled = false;
                    _btnActivate.Text = "Activated";
                    _btnActivate.BackColor = SuccessGreen;
                    _txtKey.ReadOnly = true;
                    break;

                case LicenseState.NotActivated:
                    SetStatus(
                        title: "Not activated",
                        detail: "Send your Machine ID to your supplier, then paste the key here and click Activate.",
                        bg: StatusBgNeutral,
                        fg: NeutralGray);
                    break;

                case LicenseState.Expired:
                    SetStatus(
                        title: "License expired",
                        detail: (info.Expiration.HasValue
                            ? "Expired on " + info.Expiration.Value.ToString("yyyy-MM-dd") + ". "
                            : "")
                            + "Contact your supplier to renew.",
                        bg: StatusBgWarn,
                        fg: WarnAmber);
                    break;

                case LicenseState.WrongMachine:
                    SetStatus(
                        title: "License bound to a different machine",
                        detail: "This key was issued for another computer. Send the Machine ID above to request a new key.",
                        bg: StatusBgError,
                        fg: ErrorRed);
                    break;

                default:
                    SetStatus(
                        title: "Invalid key",
                        detail: info.Message ?? "The license key is malformed or has been tampered with.",
                        bg: StatusBgError,
                        fg: ErrorRed);
                    break;
            }
        }

        private static string BuildValidDetail(LicenseInfo info)
        {
            string head = info.Expiration.HasValue
                ? "Valid until " + info.Expiration.Value.ToString("yyyy-MM-dd")
                    + " (" + info.DaysRemaining + " days remaining)"
                : "Valid";
            if (!string.IsNullOrEmpty(info.Note))
                head += "\r\nLicensed to: " + info.Note;
            return head;
        }

        private void SetStatus(string title, string detail, Color bg, Color fg)
        {
            _statusPanel.BackColor = bg;
            _statusTitle.Text = title;
            _statusTitle.ForeColor = fg;
            _statusDetail.Text = detail;
            _statusDetail.ForeColor = fg;
            _statusPanel.Invalidate();
        }

        // ---------------- Cleanup ----------------

        protected override void Dispose(bool disposing)
        {
            if (disposing && _copyResetTimer != null)
            {
                _copyResetTimer.Stop();
                _copyResetTimer.Dispose();
                _copyResetTimer = null;
            }
            base.Dispose(disposing);
        }
    }
}
