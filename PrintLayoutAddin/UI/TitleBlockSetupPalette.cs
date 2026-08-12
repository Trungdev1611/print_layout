using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
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
        private readonly Label _cornersStatusLabel;
        private readonly Button _pickCornersBtn;
        private readonly Button _resetCornersBtn;

        // Typography table (5 rows)
        private readonly ComboBox _centerStyleCombo;
        private readonly TextBox _centerHeightBox;
        private readonly ComboBox _titleStyleCombo;
        private readonly TextBox _titleHeightBox;
        private readonly ComboBox _numberStyleCombo;
        private readonly TextBox _numberHeightBox;
        private readonly ComboBox _headerStyleCombo;
        private readonly TextBox _tableHeaderHtBox;
        private readonly ComboBox _dataStyleCombo;
        private readonly TextBox _tableDataHtBox;
        private readonly TextBox _tableWidthBox;

        private readonly Button[] _actionButtons;
        private readonly Button _activateBtn;
        private readonly Button _autoBtn;
        private readonly Panel _advancedBody;
        private readonly Button _advancedToggle;
        private bool _advancedOpen;
        private readonly TableLayoutPanel _root;
        private readonly ToolTip _toolTips = new ToolTip();

        private const int HeightCollapsed = 458;
        private const int HeightExpanded = 598;

        public TitleBlockSetupPalette()
        {
            Text = "Title Block Setup (PLFRAME_SETUP)";
            Width = 560;
            Height = HeightCollapsed;
            MinimumSize = new Size(520, HeightCollapsed);
            StartPosition = FormStartPosition.Manual;
            Location = new Point(80, 80);
            FormBorderStyle = FormBorderStyle.SizableToolWindow;
            MinimizeBox = true;
            MaximizeBox = false;
            ShowInTaskbar = false;
            TopMost = false;

            _root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5,
                Padding = new Padding(8),
            };
            _root.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));   // context
            _root.RowStyles.Add(new RowStyle(SizeType.Absolute, 98));   // viewport corners
            _root.RowStyles.Add(new RowStyle(SizeType.Absolute, 192));  // typography table
            _root.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));   // auto
            _root.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));   // advanced toggle only

            // --- Context ---
            var contextRow = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
            };
            contextRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            contextRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
            _contextLabel = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new System.Drawing.Font(DefaultFont.FontFamily, 9f, System.Drawing.FontStyle.Bold),
            };
            var refreshBtn = new Button { Text = "Reload styles", Dock = DockStyle.Fill };
            refreshBtn.Click += (s, e) => RefreshContext();
            contextRow.Controls.Add(_contextLabel, 0, 0);
            contextRow.Controls.Add(refreshBtn, 1, 0);
            _root.Controls.Add(contextRow, 0, 0);

            // --- Viewport corners (P1/P2) ---
            var cornersGroup = new GroupBox
            {
                Text = "Vùng trình bày (P1/P2)",
                Dock = DockStyle.Fill,
                Padding = new Padding(8, 4, 8, 8),
            };
            var cornersOuter = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
            };
            cornersOuter.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            cornersOuter.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));

            _cornersStatusLabel = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.TopLeft,
                Font = new System.Drawing.Font(DefaultFont.FontFamily, 8.75f, System.Drawing.FontStyle.Regular),
                AutoEllipsis = true,
            };

            var cornersBtnRow = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
            };
            cornersBtnRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            cornersBtnRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            _pickCornersBtn = new Button { Text = "Pick P1/P2", Dock = DockStyle.Fill };
            _resetCornersBtn = new Button { Text = "Reset corners", Dock = DockStyle.Fill };
            _toolTips.SetToolTip(
                _pickCornersBtn,
                "Pick two opposite corners of the presentation viewport area (saved per DWG).");
            _toolTips.SetToolTip(
                _resetCornersBtn,
                "Clear saved P1/P2 for this drawing.");
            _pickCornersBtn.Click += (s, e) => RunPickCorners();
            _resetCornersBtn.Click += (s, e) => RunResetCorners();
            cornersBtnRow.Controls.Add(_pickCornersBtn, 0, 0);
            cornersBtnRow.Controls.Add(_resetCornersBtn, 1, 0);

            cornersOuter.Controls.Add(_cornersStatusLabel, 0, 0);
            cornersOuter.Controls.Add(cornersBtnRow, 0, 1);
            cornersGroup.Controls.Add(cornersOuter);
            _root.Controls.Add(cornersGroup, 0, 1);

            // --- Typography table ---
            var styleGroup = new GroupBox
            {
                Text = "Cấu hình Text Styles & Heights",
                Dock = DockStyle.Fill,
                Padding = new Padding(8, 8, 8, 6),
            };

            var styleOuter = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
            };
            styleOuter.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            styleOuter.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));

            var grid = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 6,
                CellBorderStyle = TableLayoutPanelCellBorderStyle.Single,
                BackColor = Color.White,
            };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 48));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 22));
            for (int r = 0; r < 6; r++)
                grid.RowStyles.Add(new RowStyle(SizeType.Percent, 100f / 6f));

            grid.Controls.Add(MakeHeaderLabel("Đối tượng Text (Item)"), 0, 0);
            grid.Controls.Add(MakeHeaderLabel("Text Style"), 1, 0);
            grid.Controls.Add(MakeHeaderLabel("Height (mm)"), 2, 0);

            _centerStyleCombo = MakeStyleCombo();
            _centerHeightBox = MakeOptBox(TitleBlockSetupService.DefaultCenterTitleHeight);
            AddTypoRow(grid, 1, "1. Tên bản vẽ phụ (Giữa-Dưới)", _centerStyleCombo, _centerHeightBox);

            _titleStyleCombo = MakeStyleCombo();
            _titleHeightBox = MakeOptBox(TitleBlockSetupService.DefaultTextHeight);
            AddTypoRow(grid, 2, "2. Tên bản vẽ (Khung tên)", _titleStyleCombo, _titleHeightBox);

            _numberStyleCombo = MakeStyleCombo();
            _numberHeightBox = MakeOptBox(TitleBlockSetupService.DefaultTextHeight);
            AddTypoRow(grid, 3, "3. Số hiệu bản vẽ", _numberStyleCombo, _numberHeightBox);

            _headerStyleCombo = MakeStyleCombo();
            _tableHeaderHtBox = MakeOptBox(TitleBlockSetupService.DefaultTableHeaderTextHeight);
            AddTypoRow(grid, 4, "4. Tiêu đề Bảng Rev (Header)", _headerStyleCombo, _tableHeaderHtBox);

            _dataStyleCombo = MakeStyleCombo();
            _tableDataHtBox = MakeOptBox(TitleBlockSetupService.DefaultTableDataTextHeight);
            AddTypoRow(grid, 5, "5. Nội dung Bảng Rev (Data)", _dataStyleCombo, _tableDataHtBox);

            styleOuter.Controls.Add(grid, 0, 0);

            var widthRow = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
            };
            widthRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160));
            widthRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70));
            widthRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            widthRow.Controls.Add(MakeOptLabel("Bảng Revision Width:"), 0, 0);
            _tableWidthBox = MakeOptBox(TitleBlockSetupService.DefaultTableWidth);
            widthRow.Controls.Add(_tableWidthBox, 1, 0);
            widthRow.Controls.Add(MakeOptLabel("mm"), 2, 0);
            styleOuter.Controls.Add(widthRow, 0, 1);

            styleGroup.Controls.Add(styleOuter);
            _root.Controls.Add(styleGroup, 0, 2);

            // --- Auto ---
            _autoBtn = new Button
            {
                Text = "AUTO FRAME SETUP",
                Dock = DockStyle.Fill,
                Font = new System.Drawing.Font(DefaultFont.FontFamily, 9.5f, System.Drawing.FontStyle.Regular),
                TextAlign = ContentAlignment.MiddleCenter,
            };
            _toolTips.SetToolTip(
                _autoBtn,
                "Scan labels → place fields + rev table + center title + activate");
            _autoBtn.Click += (s, e) => RunAutoFrameSetup();
            _root.Controls.Add(_autoBtn, 0, 3);

            // --- Advanced (collapsed) ---
            var advancedWrap = new Panel { Dock = DockStyle.Fill };
            _advancedToggle = new Button
            {
                Text = "► Advanced / Manual Setup",
                Dock = DockStyle.Top,
                Height = 28,
                TextAlign = ContentAlignment.MiddleLeft,
                FlatStyle = FlatStyle.Flat,
            };
            _advancedToggle.Click += (s, e) => ToggleAdvanced();

            _advancedBody = new Panel
            {
                Dock = DockStyle.Top,
                Visible = false,
                Height = 150,
                Padding = new Padding(0, 2, 0, 2),
            };
            var advInner = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5,
            };
            for (int i = 0; i < 5; i++)
                advInner.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));

            _activateBtn = MakeButton(
                "5. Activate Sheet Set fields",
                "Paper Space → layout fields; BEDIT → block attributes",
                RunActivateFields);

            _actionButtons = new[]
            {
                MakeButton(
                    "1. Place Field — Sheet Number",
                    "Pick point → DRAWING_NO (row 3 style/height)",
                    () => RunFieldPlace(TitleBlockSetupService.SheetSetFieldKind.SheetNumber)),
                MakeButton(
                    "2. Place Field — Sheet Title",
                    "Pick point → DRAWING_NAME (row 2 style/height)",
                    () => RunFieldPlace(TitleBlockSetupService.SheetSetFieldKind.SheetTitle)),
                MakeButton(
                    "3. Place Field — Revision",
                    "Pick point → REVISION (uses Số hiệu style/height)",
                    () => RunFieldPlace(TitleBlockSetupService.SheetSetFieldKind.Revision)),
                MakeButton(
                    "4. Insert Revision Table",
                    "Pick point → rows 4–5 + Revision Width",
                    RunInsertTable),
                _activateBtn,
            };
            for (int i = 0; i < _actionButtons.Length; i++)
                advInner.Controls.Add(_actionButtons[i], 0, i);

            _advancedBody.Controls.Add(advInner);
            advancedWrap.Controls.Add(_advancedBody);
            advancedWrap.Controls.Add(_advancedToggle);
            _root.Controls.Add(advancedWrap, 0, 4);

            Controls.Add(_root);

            Load += (s, e) => RefreshContext();
            AcadApp.DocumentManager.DocumentActivated += OnDocumentActivated;
            FormClosed += (s, e) =>
                AcadApp.DocumentManager.DocumentActivated -= OnDocumentActivated;
        }

        private static void AddTypoRow(
            TableLayoutPanel grid, int row, string itemLabel, ComboBox style, TextBox height)
        {
            grid.Controls.Add(MakeOptLabel(itemLabel), 0, row);
            grid.Controls.Add(style, 1, row);
            grid.Controls.Add(height, 2, row);
        }

        private void ToggleAdvanced()
        {
            _advancedOpen = !_advancedOpen;
            _advancedBody.Visible = _advancedOpen;
            _advancedToggle.Text = (_advancedOpen ? "▼" : "►") + " Advanced / Manual Setup";
            _root.RowStyles[4] = new RowStyle(
                SizeType.Absolute,
                _advancedOpen ? 180 : 30);
            Height = _advancedOpen ? HeightExpanded : HeightCollapsed;
        }

        private static Label MakeHeaderLabel(string text) =>
            new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font(DefaultFont.FontFamily, 8.25f, System.Drawing.FontStyle.Regular),
            };

        private static Label MakeOptLabel(string text) =>
            new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(4, 0, 0, 0),
            };

        private static TextBox MakeOptBox(double defaultValue) =>
            new TextBox
            {
                Dock = DockStyle.Fill,
                Text = defaultValue.ToString("0.##", CultureInfo.InvariantCulture),
                Margin = new Padding(2),
            };

        private static ComboBox MakeStyleCombo() =>
            new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                IntegralHeight = false,
                Margin = new Padding(2),
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
                if (_autoBtn != null) _autoBtn.Enabled = false;
                RefreshCornersDisplay(null, null);
                return;
            }

            var db = doc.Database;
            ReloadTextStyles(db);

            _contextLabel.ForeColor = Color.DarkGreen;
            _contextLabel.Text = "Context: " + TitleBlockSetupService.DescribeCurrentSpace(db);
            SetActionButtonsEnabled(true);
            _toolTips.SetToolTip(_activateBtn, TitleBlockSetupService.ActivateButtonHint(db));

            string dwgPath = null;
            try { dwgPath = doc.Name; } catch { }
            RefreshCornersDisplay(doc, dwgPath);
        }

        private void SetCornersStatus(string status, Color statusColor, string coords = null)
        {
            _cornersStatusLabel.ForeColor = statusColor;
            if (string.IsNullOrEmpty(coords))
            {
                _cornersStatusLabel.Font = new System.Drawing.Font(
                    DefaultFont.FontFamily, 8.75f, System.Drawing.FontStyle.Bold);
                _cornersStatusLabel.Text = status;
            }
            else
            {
                _cornersStatusLabel.Font = new System.Drawing.Font(
                    DefaultFont.FontFamily, 8.75f, System.Drawing.FontStyle.Regular);
                _cornersStatusLabel.Text = status + "\r\n" + coords;
                _cornersStatusLabel.ForeColor = Color.FromArgb(40, 40, 40);
            }
        }

        private void RefreshCornersDisplay(Document doc, string dwgPath)
        {
            if (_cornersStatusLabel == null) return;

            if (doc == null)
            {
                SetCornersStatus("No active drawing.", Color.Gray);
                _pickCornersBtn.Enabled = false;
                _resetCornersBtn.Enabled = false;
                if (_autoBtn != null) _autoBtn.Enabled = false;
                return;
            }

            bool paperSpace = !doc.Database.TileMode;
            string normalized = ViewportCornerStore.TryNormalizePath(dwgPath);
            var corners = ViewportCornerStore.Load(dwgPath);

            if (!paperSpace)
            {
                SetCornersStatus("Switch to paper space to pick P1/P2.", Color.DarkOrange);
            }
            else if (normalized == null)
            {
                SetCornersStatus("Save the DWG to disk before picking corners.", Color.DarkOrange);
            }
            else if (corners.HasValue)
            {
                SetCornersStatus(
                    "Corners saved for this drawing.",
                    Color.DarkGreen,
                    ViewportCornerPicker.FormatSaved(corners.Value));
            }
            else
            {
                SetCornersStatus(
                    "No corners saved — pick P1/P2 before Auto or Build Layouts.",
                    Color.DarkOrange);
            }

            _pickCornersBtn.Enabled = paperSpace && normalized != null;
            _resetCornersBtn.Enabled = corners.HasValue && normalized != null;
            if (_autoBtn != null)
                _autoBtn.Enabled = corners.HasValue && paperSpace;
        }

        private void RunPickCorners()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            if (db.TileMode)
            {
                MessageBox.Show(
                    this,
                    "Switch to a paper-space layout first.",
                    "Title Block Setup",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            string dwgPath = null;
            try { dwgPath = doc.Name; } catch { }
            if (ViewportCornerStore.TryNormalizePath(dwgPath) == null)
            {
                MessageBox.Show(
                    this,
                    "Save this drawing to disk first.\nCorners are stored per file path.",
                    "Title Block Setup",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            Hide();
            bool picked;
            (Point3d P1, Point3d P2) corners;
            try
            {
                picked = ViewportCornerPicker.TryPrompt(ed, out corners);
            }
            finally
            {
                Show();
                BringToFront();
            }

            if (!picked) return;

            ViewportCornerStore.Save(dwgPath, corners.P1, corners.P2);
            ed.WriteMessage("\nPLFRAME_SETUP: Saved viewport corners:\n  "
                + ViewportCornerPicker.FormatSaved(corners).Replace("\r\n", "\n  "));
            RefreshCornersDisplay(doc, dwgPath);
        }

        private void RunResetCorners()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            string dwgPath = null;
            try { dwgPath = doc.Name; } catch { }
            if (!ViewportCornerStore.Load(dwgPath).HasValue)
            {
                RefreshCornersDisplay(doc, dwgPath);
                return;
            }

            var confirm = MessageBox.Show(
                this,
                "Clear saved P1/P2 for this drawing?\nAuto Frame Setup and Build Layouts will require new picks.",
                "Title Block Setup",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            if (confirm != DialogResult.Yes) return;

            ViewportCornerStore.Clear(dwgPath);
            doc.Editor.WriteMessage("\nPLFRAME_SETUP: Cleared viewport corners for this drawing.");
            RefreshCornersDisplay(doc, dwgPath);
        }

        private void ReloadTextStyles(Database db)
        {
            string keepCenter = _centerStyleCombo.SelectedItem as string;
            string keepTitle = _titleStyleCombo.SelectedItem as string;
            string keepNumber = _numberStyleCombo.SelectedItem as string;
            string keepHeader = _headerStyleCombo.SelectedItem as string;
            string keepData = _dataStyleCombo.SelectedItem as string;

            var names = TitleBlockSetupService.ListTextStyleNames(db);
            FillStyleCombo(_centerStyleCombo, names, keepCenter);
            FillStyleCombo(_titleStyleCombo, names, keepTitle);
            FillStyleCombo(_numberStyleCombo, names, keepNumber);
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

        private bool EnsureAllowedSpace(Document doc) =>
            doc?.Database != null;

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

        private void GetFieldStyleHeight(
            TitleBlockSetupService.SheetSetFieldKind kind,
            out double height,
            out string style)
        {
            // Revision field has no dedicated row → reuse Số hiệu (row 3).
            if (kind == TitleBlockSetupService.SheetSetFieldKind.SheetTitle)
            {
                TryReadPositive(_titleHeightBox, "Title height", TitleBlockSetupService.DefaultTextHeight, out height);
                style = SelectedStyle(_titleStyleCombo);
            }
            else
            {
                TryReadPositive(_numberHeightBox, "Number height", TitleBlockSetupService.DefaultTextHeight, out height);
                style = SelectedStyle(_numberStyleCombo);
            }
        }

        private void RunFieldPlace(TitleBlockSetupService.SheetSetFieldKind kind)
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            if (!EnsureAllowedSpace(doc)) return;
            GetFieldStyleHeight(kind, out double textHeight, out string textStyle);

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
            }
        }

        private void RunAutoFrameSetup()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            if (!EnsureAllowedSpace(doc)) return;

            if (db.TileMode)
            {
                MessageBox.Show(
                    this,
                    "Switch to a paper-space layout first.\nAuto Frame Setup uses PLAYOUT viewport corners.",
                    "Title Block Setup",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            string dwgPath = null;
            try { dwgPath = doc.Name; } catch { }
            var corners = ViewportCornerStore.Load(dwgPath);
            if (!corners.HasValue)
            {
                var pickNow = MessageBox.Show(
                    this,
                    "No saved viewport corners for this DWG.\nPick P1/P2 now?",
                    "Title Block Setup",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if (pickNow != DialogResult.Yes) return;

                RunPickCorners();
                corners = ViewportCornerStore.Load(dwgPath);
                if (!corners.HasValue) return;
            }

            TryReadPositive(
                _centerHeightBox, "Center title height",
                TitleBlockSetupService.DefaultCenterTitleHeight, out double centerHeight);
            TryReadPositive(
                _titleHeightBox, "Title height",
                TitleBlockSetupService.DefaultTextHeight, out double titleHeight);
            TryReadPositive(
                _numberHeightBox, "Number height",
                TitleBlockSetupService.DefaultTextHeight, out double numberHeight);
            TryReadPositive(
                _tableWidthBox, "Table width",
                TitleBlockSetupService.DefaultTableWidth, out double tableWidth);
            TryReadPositive(
                _tableHeaderHtBox, "Header ht",
                TitleBlockSetupService.DefaultTableHeaderTextHeight, out double headerHt);
            TryReadPositive(
                _tableDataHtBox, "Data ht",
                TitleBlockSetupService.DefaultTableDataTextHeight, out double dataHt);

            string centerStyle = SelectedStyle(_centerStyleCombo);
            string titleStyle = SelectedStyle(_titleStyleCombo);
            string numberStyle = SelectedStyle(_numberStyleCombo);
            string headerStyle = SelectedStyle(_headerStyleCombo);
            string dataStyle = SelectedStyle(_dataStyleCombo);

            var bounds = ViewportCornerGeometry.Normalize(corners.Value);
            bool ok;
            string message;
            using (doc.LockDocument())
            {
                // Place title/number/rev with per-row styles, then center + table + activate.
                ok = TitleBlockSetupService.RunAutoFrameSetup(
                    db, bounds, out message,
                    titleHeight, titleStyle,
                    centerHeight, centerStyle,
                    tableWidth, headerHt, dataHt,
                    headerStyle, dataStyle,
                    numberHeight, numberStyle,
                    numberHeight, numberStyle);
            }

            ed.WriteMessage("\nPLFRAME_SETUP Auto:\n" + (message ?? (ok ? "Done." : "Failed.")));
            if (ok)
            {
                try
                {
                    doc.SendStringToExecute("_.UPDATEFIELD _All ", true, false, false);
                }
                catch { }
            }
            else
            {
                MessageBox.Show(
                    this,
                    message ?? "Auto Frame Setup failed.",
                    "Title Block Setup",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }
    }
}
