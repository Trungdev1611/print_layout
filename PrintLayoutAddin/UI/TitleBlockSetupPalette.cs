using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using PrintLayoutAddin.Core;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace PrintLayoutAddin.UI
{
    /// <summary>
    /// Modeless palette for <c>PLFRAME_SETUP</c>. Hides during entity/point picks.
    /// </summary>
    public sealed class TitleBlockSetupPalette : Form
    {
        private readonly Label _contextLabel;
        private readonly Label _hintLabel;
        private readonly Label _notesLabel;
        private readonly TextBox _textHeightBox;
        private readonly TextBox _tableWidthBox;
        private readonly TextBox _tableHeaderHtBox;
        private readonly TextBox _tableDataHtBox;
        private readonly ComboBox _textStyleCombo;
        private readonly ComboBox _headerStyleCombo;
        private readonly ComboBox _dataStyleCombo;
        private readonly Button[] _actionButtons;
        private readonly Button _activateBtn;
        private readonly ToolTip _toolTips = new ToolTip();

        public TitleBlockSetupPalette()
        {
            Text = "Title Block Setup (PLFRAME_SETUP)";
            Width = 520;
            Height = 620;
            MinimumSize = new Size(480, 560);
            StartPosition = FormStartPosition.Manual;
            Location = new Point(80, 80);
            FormBorderStyle = FormBorderStyle.SizableToolWindow;
            MinimizeBox = true;
            MaximizeBox = false;
            ShowInTaskbar = false;
            TopMost = false;

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 11,
                Padding = new Padding(10),
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));   // context
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));   // hint
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 120));  // options
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));   // btn1
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));   // btn2
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));   // btn3
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));   // btn4
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));   // btn5
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));   // refresh
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));   // notes
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 0));

            _contextLabel = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new System.Drawing.Font(DefaultFont.FontFamily, 9f, System.Drawing.FontStyle.Bold),
            };
            root.Controls.Add(_contextLabel, 0, 0);

            _hintLabel = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.TopLeft,
                Text =
                    "1–3: placeholder. 4: revision table. 5: activate Sheet Set.\n"
                    + "OK in Model Space (xref title-block source), Paper Space, or BEDIT.",
            };
            root.Controls.Add(_hintLabel, 0, 1);

            var options = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 4,
            };
            options.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            options.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            options.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            options.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            for (int r = 0; r < 4; r++)
                options.RowStyles.Add(new RowStyle(SizeType.Percent, 25));

            options.Controls.Add(MakeOptLabel("Text height"), 0, 0);
            _textHeightBox = MakeOptBox(TitleBlockSetupService.DefaultTextHeight);
            options.Controls.Add(_textHeightBox, 1, 0);

            options.Controls.Add(MakeOptLabel("Text style"), 2, 0);
            _textStyleCombo = MakeStyleCombo();
            options.Controls.Add(_textStyleCombo, 3, 0);

            options.Controls.Add(MakeOptLabel("Table width"), 0, 1);
            _tableWidthBox = MakeOptBox(TitleBlockSetupService.DefaultTableWidth);
            options.Controls.Add(_tableWidthBox, 1, 1);

            options.Controls.Add(MakeOptLabel(""), 2, 1);
            options.Controls.Add(MakeOptLabel(""), 3, 1);

            options.Controls.Add(MakeOptLabel("Header ht"), 0, 2);
            _tableHeaderHtBox = MakeOptBox(TitleBlockSetupService.DefaultTableHeaderTextHeight);
            options.Controls.Add(_tableHeaderHtBox, 1, 2);

            options.Controls.Add(MakeOptLabel("Header style"), 2, 2);
            _headerStyleCombo = MakeStyleCombo();
            options.Controls.Add(_headerStyleCombo, 3, 2);

            options.Controls.Add(MakeOptLabel("Data ht"), 0, 3);
            _tableDataHtBox = MakeOptBox(TitleBlockSetupService.DefaultTableDataTextHeight);
            options.Controls.Add(_tableDataHtBox, 1, 3);

            options.Controls.Add(MakeOptLabel("Data style"), 2, 3);
            _dataStyleCombo = MakeStyleCombo();
            options.Controls.Add(_dataStyleCombo, 3, 3);

            root.Controls.Add(options, 0, 2);

            _activateBtn = MakeButton(
                "5. Activate Sheet Set fields",
                "Paper Space → layout fields; BEDIT → block attributes",
                RunActivateFields);

            _actionButtons = new[]
            {
                MakeButton(
                    "1. Place Field — Sheet Number",
                    "Pick point → DRAWING_NO (Text height + Text style)",
                    () => RunFieldPlace(TitleBlockSetupService.SheetSetFieldKind.SheetNumber)),
                MakeButton(
                    "2. Place Field — Sheet Title",
                    "Pick point → DRAWING_NAME (Text height + Text style)",
                    () => RunFieldPlace(TitleBlockSetupService.SheetSetFieldKind.SheetTitle)),
                MakeButton(
                    "3. Place Field — Revision",
                    "Pick point → REVISION (Text height + Text style)",
                    () => RunFieldPlace(TitleBlockSetupService.SheetSetFieldKind.Revision)),
                MakeButton(
                    "4. Insert Revision Table",
                    "Pick point → width, header/data height & style",
                    RunInsertTable),
                _activateBtn,
            };

            for (int i = 0; i < _actionButtons.Length; i++)
                root.Controls.Add(_actionButtons[i], 0, i + 3);

            var closePanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft,
            };
            var refreshBtn = new Button { Text = "Refresh context", AutoSize = true };
            refreshBtn.Click += (s, e) => RefreshContext();
            closePanel.Controls.Add(refreshBtn);
            root.Controls.Add(closePanel, 0, 8);

            _notesLabel = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.TopLeft,
                ForeColor = Color.DimGray,
                Font = new System.Drawing.Font(DefaultFont.FontFamily, 8f),
                Text =
                    "Ghi chú / Notes:\n"
                    + "• Text height — chiều cao chữ số hiệu, tên bản vẽ, revision (nút 1–3).\n"
                    + "• Text style — kiểu chữ (Text Style) cho các field nút 1–3.\n"
                    + "• Table width — tổng chiều rộng bảng revision (nút 4), mặc định 75.\n"
                    + "• Header ht / Header style — cỡ chữ & kiểu chữ hàng tiêu đề (Rev / Description / Date).\n"
                    + "• Data ht / Data style — cỡ chữ & kiểu chữ các dòng dữ liệu revision.\n"
                    + "• Nút 5 — gắn field Sheet Set (Model / Paper / BEDIT).\n"
                    + "• Model Space — dùng khi setup file khung tên rồi xref sang host.",
            };
            root.Controls.Add(_notesLabel, 0, 9);

            Controls.Add(root);

            Load += (s, e) => RefreshContext();
            AcadApp.DocumentManager.DocumentActivated += OnDocumentActivated;
            FormClosed += (s, e) =>
                AcadApp.DocumentManager.DocumentActivated -= OnDocumentActivated;
        }

        private static Label MakeOptLabel(string text) =>
            new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
            };

        private static TextBox MakeOptBox(double defaultValue) =>
            new TextBox
            {
                Dock = DockStyle.Fill,
                Text = defaultValue.ToString("0.##", CultureInfo.InvariantCulture),
            };

        private static ComboBox MakeStyleCombo() =>
            new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                IntegralHeight = false,
            };

        private void OnDocumentActivated(object sender, DocumentCollectionEventArgs e) =>
            RefreshContext();

        public void RefreshContext()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                _contextLabel.Text = "No active drawing.";
                SetActionButtonsEnabled(false);
                return;
            }

            var db = doc.Database;
            ReloadTextStyles(db);

            _contextLabel.ForeColor = Color.DarkGreen;
            _contextLabel.Text = "Context: " + TitleBlockSetupService.DescribeCurrentSpace(db);
            SetActionButtonsEnabled(true);
            _toolTips.SetToolTip(_activateBtn, TitleBlockSetupService.ActivateButtonHint(db));
        }

        private void ReloadTextStyles(Database db)
        {
            string keepText = _textStyleCombo.SelectedItem as string;
            string keepHeader = _headerStyleCombo.SelectedItem as string;
            string keepData = _dataStyleCombo.SelectedItem as string;

            var names = TitleBlockSetupService.ListTextStyleNames(db);
            FillStyleCombo(_textStyleCombo, names, keepText);
            FillStyleCombo(_headerStyleCombo, names, keepHeader);
            FillStyleCombo(_dataStyleCombo, names, keepData);
        }

        private static void FillStyleCombo(ComboBox combo, System.Collections.Generic.List<string> names, string preferred)
        {
            if (combo == null) return;
            combo.BeginUpdate();
            try
            {
                combo.Items.Clear();
                foreach (var n in names)
                    combo.Items.Add(n);

                if (!string.IsNullOrWhiteSpace(preferred))
                {
                    int idx = combo.FindStringExact(preferred);
                    if (idx >= 0)
                    {
                        combo.SelectedIndex = idx;
                        return;
                    }
                }

                int std = combo.FindStringExact("Standard");
                combo.SelectedIndex = std >= 0 ? std : (combo.Items.Count > 0 ? 0 : -1);
            }
            finally
            {
                combo.EndUpdate();
            }
        }

        private static string SelectedStyle(ComboBox combo)
        {
            if (combo?.SelectedItem == null) return "Standard";
            return combo.SelectedItem.ToString();
        }

        private void SetActionButtonsEnabled(bool enabled)
        {
            if (_actionButtons == null) return;
            foreach (var btn in _actionButtons)
                btn.Enabled = enabled;
        }

        private bool EnsureAllowedSpace(Document doc)
        {
            if (doc == null) return false;
            if (doc.Database != null) return true;
            return false;
        }

        private bool TryReadPositive(TextBox box, string label, double fallback, out double value)
        {
            value = fallback;
            string raw = (box?.Text ?? "").Trim().Replace(',', '.');
            if (string.IsNullOrEmpty(raw))
                return true;
            if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out value)
                || value <= 0)
            {
                MessageBox.Show(
                    this,
                    $"{label} must be a positive number.\nUsing default {fallback:0.##}.",
                    "Title Block Setup",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                value = fallback;
                if (box != null)
                    box.Text = fallback.ToString("0.##", CultureInfo.InvariantCulture);
                return true;
            }
            return true;
        }

        private Button MakeButton(string text, string tip, Action action)
        {
            var btn = new Button
            {
                Text = text,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
            };
            _toolTips.SetToolTip(btn, tip);
            btn.Click += (s, e) => action();
            return btn;
        }

        private void RunFieldPlace(TitleBlockSetupService.SheetSetFieldKind kind)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            if (!EnsureAllowedSpace(doc)) return;
            TryReadPositive(
                _textHeightBox, "Text height",
                TitleBlockSetupService.DefaultTextHeight, out double textHeight);
            string textStyle = SelectedStyle(_textStyleCombo);

            var ppo = new PromptPointOptions(
                "\n" + TitleBlockSetupService.FieldPlacementPrompt(kind));
            ppo.AllowNone = false;

            Hide();
            PromptPointResult ppr;
            try
            {
                ppr = ed.GetPoint(ppo);
            }
            finally
            {
                Show();
                BringToFront();
            }

            if (ppr.Status != PromptStatus.OK) return;

            bool ok;
            string message;
            using (doc.LockDocument())
            {
                ok = TitleBlockSetupService.InsertFieldAtPoint(
                    db, ppr.Value, kind, out message, textHeight, textStyle);
            }

            ed.WriteMessage("\nPLFRAME_SETUP: " + (message ?? (ok ? "Done." : "Failed.")));
        }

        private void RunInsertTable()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            if (!EnsureAllowedSpace(doc)) return;
            TryReadPositive(
                _tableWidthBox, "Table width",
                TitleBlockSetupService.DefaultTableWidth, out double tableWidth);
            TryReadPositive(
                _tableHeaderHtBox, "Header ht",
                TitleBlockSetupService.DefaultTableHeaderTextHeight, out double headerHt);
            TryReadPositive(
                _tableDataHtBox, "Data ht",
                TitleBlockSetupService.DefaultTableDataTextHeight, out double dataHt);
            string headerStyle = SelectedStyle(_headerStyleCombo);
            string dataStyle = SelectedStyle(_dataStyleCombo);

            var ppo = new PromptPointOptions("\nPick revision table insertion point (top-left): ");
            ppo.AllowNone = false;

            Hide();
            PromptPointResult ppr;
            try
            {
                ppr = ed.GetPoint(ppo);
            }
            finally
            {
                Show();
                BringToFront();
            }

            if (ppr.Status != PromptStatus.OK) return;

            bool ok;
            string message;
            using (doc.LockDocument())
            {
                ok = TitleBlockSetupService.InsertRevisionTable(
                    db, ppr.Value, out message, tableWidth, headerHt, dataHt,
                    headerStyle, dataStyle);
            }

            ed.WriteMessage("\nPLFRAME_SETUP: " + (message ?? (ok ? "Done." : "Failed.")));
        }

        private void RunActivateFields()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            if (!EnsureAllowedSpace(doc)) return;

            bool ok;
            string message;
            using (doc.LockDocument())
            {
                ok = TitleBlockSetupService.ActivateSheetSetFieldsOnPlaceholders(db, out message);
            }

            ed.WriteMessage("\nPLFRAME_SETUP: " + (message ?? (ok ? "Done." : "Failed.")));
            if (ok)
            {
                try
                {
                    doc.SendStringToExecute("_.UPDATEFIELD _All ", true, false, false);
                }
                catch { }
                ed.WriteMessage(
                    "\nTip: if text still shows #### from an older activate, erase those texts, "
                    + "place 1–3 again, then press 5.");
            }
        }
    }
}
