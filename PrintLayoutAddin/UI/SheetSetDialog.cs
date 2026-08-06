using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Win32;
using PrintLayoutAddin.Core;

namespace PrintLayoutAddin.UI
{
    public class SheetSetDialog : Form
    {
        private const string RegKey = @"Software\PrintLayoutAddin\SheetSet";

        private readonly BindingList<SheetSetEntry> _entries;
        private readonly BindingSource _source = new BindingSource();
        private DataGridView _grid;
        private TextBox _nameBox;
        private TextBox _pathBox;
        private Label _status;
        private Button _createBtn;
        private Button _exportBtn;

        public List<PrintableLayout> ExportLayouts { get; private set; }

        public SheetSetDialog(
            IEnumerable<PrintableLayout> layouts,
            IDictionary<string, string> drawingNames,
            string dwgPath,
            string defaultDstPath)
        {
            var fixedDstPath = string.IsNullOrWhiteSpace(defaultDstPath)
                ? PublishPaths.DefaultDstPath(dwgPath)
                : defaultDstPath;

            var items = (layouts ?? Enumerable.Empty<PrintableLayout>())
                .Select((layout, index) => new SheetSetEntry
                {
                    Include = true,
                    Order = index + 1,
                    SheetNumber = layout.Name,
                    Title = drawingNames != null
                            && drawingNames.TryGetValue(layout.Name, out var title)
                            && !string.IsNullOrWhiteSpace(title)
                        ? title
                        : layout.Name,
                    DwgPath = dwgPath,
                    Layout = layout,
                })
                .ToList();

            string sheetSetName = Path.GetFileNameWithoutExtension(dwgPath ?? "") ?? "SheetSet";
            string loadedFrom = null;
            var existing = SheetSetService.TryRead(fixedDstPath);
            if (existing != null)
            {
                loadedFrom = fixedDstPath;
                if (!string.IsNullOrWhiteSpace(existing.SheetSetName))
                    sheetSetName = existing.SheetSetName;

                var byLayout = existing.Sheets
                    .Where(s => !string.IsNullOrWhiteSpace(s.LayoutName))
                    .GroupBy(s => s.LayoutName, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

                foreach (var entry in items)
                {
                    if (!byLayout.TryGetValue(entry.LayoutName, out var sheet)) continue;
                    if (!string.IsNullOrWhiteSpace(sheet.Number))
                        entry.SheetNumber = sheet.Number;
                    if (!string.IsNullOrWhiteSpace(sheet.Title))
                        entry.Title = sheet.Title;
                }
            }

            _entries = new BindingList<SheetSetEntry>(items);

            Text = "Create Sheet Set";
            Width = 1050;
            Height = 680;
            MinimumSize = new Size(850, 560);
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = true;

            BuildUi();
            ApplyFixedPath(fixedDstPath, sheetSetName, loadedFrom);
            RefreshOrders();
            ValidateUi();
        }

        private void BuildUi()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 1,
                RowCount = 4,
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 106));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 56));

            root.Controls.Add(BuildHeader(), 0, 0);
            root.Controls.Add(BuildToolbar(), 0, 1);
            root.Controls.Add(BuildGrid(), 0, 2);
            root.Controls.Add(BuildFooter(), 0, 3);
            Controls.Add(root);
        }

        private Control BuildHeader()
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 5,
                RowCount = 3,
            };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 105));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 38));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 62));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));

            panel.Controls.Add(LabelFor("Sheet set name"), 0, 0);
            _nameBox = new TextBox { Dock = DockStyle.Fill };
            panel.Controls.Add(_nameBox, 1, 0);

            panel.Controls.Add(LabelFor("DST file"), 0, 1);
            _pathBox = new TextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackColor = SystemColors.Control,
            };
            panel.Controls.Add(_pathBox, 1, 1);
            panel.SetColumnSpan(_pathBox, 4);

            var hint = new Label
            {
                Text = "DST is fixed under the drawing's sheetset_manager folder. "
                    + "Existing DST values are loaded when present; otherwise Sheet Number/Title "
                    + "start from layout STT / " + Config.Instance.DrawingNameTag + ".",
                Dock = DockStyle.Fill,
                ForeColor = Color.DimGray,
                TextAlign = ContentAlignment.MiddleLeft,
            };
            panel.Controls.Add(hint, 2, 0);
            panel.SetColumnSpan(hint, 3);

            var fieldHint = new Label
            {
                Text = "Title block setup: use SheetSet fields CurrentSheetNumber and CurrentSheetTitle. "
                    + "After Create / Update DST, save the DWG and REGEN to refresh those fields.",
                Dock = DockStyle.Fill,
                ForeColor = Color.FromArgb(45, 95, 150),
                TextAlign = ContentAlignment.MiddleLeft,
            };
            panel.Controls.Add(fieldHint, 0, 2);
            panel.SetColumnSpan(fieldHint, 5);
            return panel;
        }

        private Control BuildToolbar()
        {
            var bar = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(0, 5, 0, 3),
            };
            var up = ToolbarButton("Move Up", () => MoveSelected(-1));
            var down = ToolbarButton("Move Down", () => MoveSelected(1));
            var renumber = ToolbarButton("Renumber 1…N", Renumber);
            var selectAll = ToolbarButton("Select All", () => SetAllIncluded(true));
            var selectNone = ToolbarButton("Select None", () => SetAllIncluded(false));
            bar.Controls.Add(up);
            bar.Controls.Add(down);
            bar.Controls.Add(renumber);
            bar.Controls.Add(new Label { Width = 14 });
            bar.Controls.Add(selectAll);
            bar.Controls.Add(selectNone);
            return bar;
        }

        private Control BuildGrid()
        {
            _grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToOrderColumns = false,
                MultiSelect = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                BackgroundColor = Color.White,
            };
            _grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(246, 248, 251);
            _grid.Columns.Add(new DataGridViewCheckBoxColumn
            {
                HeaderText = "Use",
                DataPropertyName = nameof(SheetSetEntry.Include),
                Width = 48,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
            });
            _grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Order",
                DataPropertyName = nameof(SheetSetEntry.Order),
                ReadOnly = true,
                Width = 58,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
            });
            _grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Sheet Number (initially STT)",
                DataPropertyName = nameof(SheetSetEntry.SheetNumber),
                FillWeight = 22,
            });
            _grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Drawing Name / Sheet Title",
                DataPropertyName = nameof(SheetSetEntry.Title),
                FillWeight = 42,
            });
            _grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Layout",
                DataPropertyName = nameof(SheetSetEntry.LayoutName),
                ReadOnly = true,
                FillWeight = 24,
            });
            _grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "DWG",
                DataPropertyName = nameof(SheetSetEntry.DwgName),
                ReadOnly = true,
                FillWeight = 22,
            });

            _source.DataSource = _entries;
            _grid.DataSource = _source;
            // Only auto-commit checkbox toggles. Committing text cells here ends the
            // edit after the first typed character (classic DataGridView gotcha).
            _grid.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (!_grid.IsCurrentCellDirty) return;
                if (_grid.CurrentCell is DataGridViewCheckBoxCell)
                    _grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };
            _grid.CellValueChanged += (s, e) => ValidateUi();
            return _grid;
        }

        private Control BuildFooter()
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
            };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            _status = new Label
            {
                Text = "",
                Dock = DockStyle.Fill,
                ForeColor = Color.DimGray,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true,
            };
            panel.Controls.Add(_status, 0, 0);

            var buttons = new FlowLayoutPanel
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Padding = new Padding(0, 9, 0, 0),
            };
            var close = new Button { Text = "Close", Width = 100, Height = 32, DialogResult = DialogResult.Cancel };
            _exportBtn = new Button { Text = "Export PDF…", Width = 120, Height = 32 };
            _createBtn = new Button { Text = "Create / Update DST", Width = 155, Height = 32 };
            _exportBtn.Click += (s, e) => RequestExport();
            _createBtn.Click += (s, e) => CreateDst();
            buttons.Controls.Add(close);
            buttons.Controls.Add(_exportBtn);
            buttons.Controls.Add(_createBtn);
            panel.Controls.Add(buttons, 1, 0);
            return panel;
        }

        private void MoveSelected(int delta)
        {
            CommitGrid();
            if (_grid.CurrentRow == null) return;
            int oldIndex = _grid.CurrentRow.Index;
            int newIndex = oldIndex + delta;
            if (oldIndex < 0 || newIndex < 0 || newIndex >= _entries.Count) return;
            var item = _entries[oldIndex];
            _entries.RaiseListChangedEvents = false;
            _entries.RemoveAt(oldIndex);
            _entries.Insert(newIndex, item);
            _entries.RaiseListChangedEvents = true;
            RefreshOrders();
            _source.ResetBindings(false);
            _grid.ClearSelection();
            _grid.Rows[newIndex].Selected = true;
            _grid.CurrentCell = _grid.Rows[newIndex].Cells[0];
        }

        private void Renumber()
        {
            CommitGrid();
            int number = 1;
            foreach (var entry in _entries)
            {
                if (!entry.Include) continue;
                entry.SheetNumber = number.ToString();
                number++;
            }
            _source.ResetBindings(false);
            SetStatus($"Renumbered {number - 1} selected sheet(s).", false);
        }

        private void SetAllIncluded(bool include)
        {
            CommitGrid();
            foreach (var entry in _entries) entry.Include = include;
            _source.ResetBindings(false);
            ValidateUi();
        }

        private void CreateDst()
        {
            CommitGrid();
            var selected = GetSelectedEntries();
            if (selected.Count == 0) return;
            var path = NormalizeDstPath(_pathBox.Text);
            if (string.IsNullOrWhiteSpace(path))
            {
                SetStatus("DST path is missing.", true);
                return;
            }
            if (File.Exists(path))
            {
                var answer = MessageBox.Show(
                    this,
                    "The DST file already exists and will be rebuilt from the current table:\n\n"
                    + path + "\n\nContinue?",
                    "Create / Update Sheet Set",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if (answer != DialogResult.Yes) return;
            }

            try
            {
                UseWaitCursor = true;
                _createBtn.Enabled = false;
                SheetSetService.CreateOrReplace(path, _nameBox.Text, selected);
                try
                {
                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
                        .MdiActiveDocument?.Editor.Regen();
                }
                catch { }
                _pathBox.Text = path;
                SaveDefaults();
                SetStatus($"Created {selected.Count} sheet(s): {path}", false);
                MessageBox.Show(
                    this,
                    $"Sheet set created successfully.\n\nSheets: {selected.Count}\nFile: {path}\n\n"
                    + "Sheet Set fields were refreshed. Save the DWG to persist its sheet-set association.",
                    "Create Sheet Set",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                SetStatus(ex.Message, true);
                MessageBox.Show(this, ex.Message, "Create Sheet Set", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                UseWaitCursor = false;
                ValidateUi();
            }
        }

        private void RequestExport()
        {
            CommitGrid();
            var selected = GetSelectedEntries();
            if (selected.Count == 0) return;
            ExportLayouts = selected.Select(x => x.Layout).ToList();
            SaveDefaults();
            DialogResult = DialogResult.OK;
            Close();
        }

        private List<SheetSetEntry> GetSelectedEntries()
        {
            return _entries.Where(x => x.Include).OrderBy(x => x.Order).ToList();
        }

        private void RefreshOrders()
        {
            for (int i = 0; i < _entries.Count; i++) _entries[i].Order = i + 1;
            _source.ResetBindings(false);
            ValidateUi();
        }

        private void CommitGrid()
        {
            try
            {
                _grid?.EndEdit();
                _source.EndEdit();
            }
            catch { }
        }

        private void ValidateUi()
        {
            if (_createBtn == null || _exportBtn == null) return;
            int count = _entries.Count(x => x.Include);
            bool hasSheets = count > 0;
            _createBtn.Enabled = hasSheets && !string.IsNullOrWhiteSpace(_pathBox.Text);
            _exportBtn.Enabled = hasSheets;
            if (string.IsNullOrWhiteSpace(_status.Text))
                _status.Text = $"{count} of {_entries.Count} layout(s) selected.";
        }

        private void ApplyFixedPath(string dstPath, string sheetSetName, string loadedFrom)
        {
            _pathBox.Text = dstPath ?? "";
            _nameBox.Text = sheetSetName ?? "";
            if (!string.IsNullOrWhiteSpace(loadedFrom))
                SetStatus("Loaded Number/Title from existing DST.", false);
            else
                SetStatus("No DST yet — seeded from layout STT / drawing-name attributes.", false);
        }

        private void SaveDefaults()
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(RegKey))
                {
                    key?.SetValue("LastDstPath", _pathBox.Text ?? "");
                    key?.SetValue("LastSheetSetName", _nameBox.Text ?? "");
                }
            }
            catch { }
        }

        private void SetStatus(string text, bool isError)
        {
            _status.Text = text ?? "";
            _status.ForeColor = isError ? Color.DarkRed : Color.DarkGreen;
        }

        private static Label LabelFor(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
            };
        }

        private static Button ToolbarButton(string text, Action action)
        {
            var button = new Button { Text = text, Width = 105, Height = 30, Margin = new Padding(0, 0, 8, 0) };
            button.Click += (s, e) => action();
            return button;
        }

        private static string NormalizeDstPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return "";
            var value = path.Trim();
            return value.EndsWith(".dst", StringComparison.OrdinalIgnoreCase) ? value : value + ".dst";
        }
    }
}
