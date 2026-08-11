using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using PrintLayoutAddin.Core;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace PrintLayoutAddin.UI
{
    public class SheetSetDialog : Form
    {
        private const string RegKey = @"Software\PrintLayoutAddin\SheetSet";

        private readonly string _dwgPath;
        private readonly string _fixedDstPath;
        private readonly IDictionary<string, string> _drawingNames;
        private readonly List<PrintableLayout> _layouts;
        private readonly BindingList<SheetSetEntry> _entries;
        private readonly BindingSource _source = new BindingSource();
        private DataGridView _grid;
        private TextBox _nameBox;
        private TextBox _pathBox;
        private Label _status;
        private Button _createBtn;
        private Button _exportBtn;
        private CheckBox _openSsmChk;

        public List<PrintableLayout> ExportLayouts { get; private set; }

        public SheetSetDialog(
            IEnumerable<PrintableLayout> layouts,
            IDictionary<string, string> drawingNames,
            string dwgPath,
            string defaultDstPath)
        {
            _dwgPath = dwgPath;
            _drawingNames = drawingNames;
            _layouts = (layouts ?? Enumerable.Empty<PrintableLayout>()).ToList();
            _fixedDstPath = string.IsNullOrWhiteSpace(defaultDstPath)
                ? PublishPaths.DefaultDstPath(dwgPath)
                : defaultDstPath;

            _entries = new BindingList<SheetSetEntry>();

            Text = "Create Sheet Set";
            Width = 1050;
            Height = 680;
            MinimumSize = new Size(850, 560);
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = true;

            BuildUi();
            var sheetSetName = ReloadFromDstOrModel();
            ApplyFixedPath(_fixedDstPath, sheetSetName);
            ValidateUi();
        }

        /// <summary>
        /// Rebuilds the table from DST (subset headers + sheets) when present;
        /// otherwise seeds from model layouts / drawing-name attributes.
        /// When DST has content, the grid mirrors it only (no auto-append of
        /// local layouts missing from the DST).
        /// </summary>
        private string ReloadFromDstOrModel()
        {
            _entries.RaiseListChangedEvents = false;
            _entries.Clear();

            string sheetSetName = Path.GetFileNameWithoutExtension(_dwgPath ?? "") ?? "SheetSet";
            var layoutByName = _layouts
                .Where(l => l != null && !string.IsNullOrWhiteSpace(l.Name))
                .GroupBy(l => l.Name, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

            var existing = SheetSetService.TryRead(_fixedDstPath);

            if (existing != null && existing.Nodes.Count > 0)
            {
                if (!string.IsNullOrWhiteSpace(existing.SheetSetName))
                    sheetSetName = existing.SheetSetName;

                foreach (var node in existing.Nodes)
                {
                    if (node == null) continue;
                    if (node.Kind == SheetSetRowKind.Subset)
                    {
                        _entries.Add(new SheetSetEntry
                        {
                            Kind = SheetSetRowKind.Subset,
                            Include = false,
                            SubsetName = node.SubsetName ?? node.Title,
                            SubsetLevel = Math.Max(1, node.SubsetLevel),
                            Title = node.Title ?? node.SubsetName,
                            SheetNumber = "",
                            DwgPath = _dwgPath,
                        });
                        continue;
                    }

                    layoutByName.TryGetValue(node.LayoutName ?? "", out var layout);

                    string dwgPath = !string.IsNullOrWhiteSpace(node.DwgPath)
                        ? node.DwgPath
                        : _dwgPath;
                    if (layout == null && !string.IsNullOrWhiteSpace(node.LayoutName))
                        layout = new PrintableLayout { Name = node.LayoutName };

                    _entries.Add(new SheetSetEntry
                    {
                        Kind = SheetSetRowKind.Sheet,
                        Include = layout != null && !string.IsNullOrWhiteSpace(dwgPath),
                        SubsetName = node.SubsetName,
                        SubsetLevel = Math.Max(0, node.SubsetLevel),
                        SheetNumber = !string.IsNullOrWhiteSpace(node.Number)
                            ? node.Number
                            : (node.LayoutName ?? ""),
                        Title = !string.IsNullOrWhiteSpace(node.Title)
                            ? node.Title
                            : ResolveDrawingName(node.LayoutName),
                        Revision = node.Revision ?? "",
                        DwgPath = dwgPath,
                        Layout = layout,
                    });
                }

                // Do not re-append local DWG layouts missing from the DST — that undoes
                // intentional sheet deletes after Create/Update (SSM stays correct while
                // this grid would resurrect them on Refresh / reopen).
                SetStatus("Loaded from DST (subsets mirrored for display).", false);
            }
            else
            {
                foreach (var layout in _layouts)
                {
                    if (layout == null) continue;
                    _entries.Add(new SheetSetEntry
                    {
                        Kind = SheetSetRowKind.Sheet,
                        Include = true,
                        SheetNumber = layout.Name,
                        Title = ResolveDrawingName(layout.Name),
                        Revision = "",
                        DwgPath = _dwgPath,
                        Layout = layout,
                    });
                }
                SetStatus(
                    existing == null
                        ? "No DST yet — seeded from layout STT / drawing-name attributes."
                        : "DST has no sheets — seeded from model.",
                    false);
            }

            _entries.RaiseListChangedEvents = true;
            RefreshOrders();
            ApplyRevisionSummaries();
            SeedRevisionFromHistoryIfEmpty();
            _source.DataSource = _entries;
            _source.ResetBindings(false);
            ApplyRowStyles();
            return sheetSetName;
        }

        /// <summary>
        /// If Current Version is empty, seed from the last history-table Rev No (text before " - ").
        /// </summary>
        private void SeedRevisionFromHistoryIfEmpty()
        {
            foreach (var entry in _entries)
            {
                if (entry == null || entry.IsSubset) continue;
                if (!string.IsNullOrWhiteSpace(entry.Revision)) continue;
                var summary = entry.LastRevisionSummary;
                if (string.IsNullOrWhiteSpace(summary)) continue;
                int sep = summary.IndexOf(" - ", StringComparison.Ordinal);
                entry.Revision = (sep > 0 ? summary.Substring(0, sep) : summary).Trim();
            }
        }

        private void ApplyRevisionSummaries()
        {
            var doc = AcadApp.DocumentManager?.MdiActiveDocument;
            if (doc == null) return;

            var names = _entries
                .Where(e => e != null && !e.IsSubset && e.Layout != null)
                .Select(e => e.Layout.Name)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (names.Count == 0) return;

            var map = RevisionTableService.ReadLastRevisionSummaries(doc.Database, names);
            foreach (var entry in _entries)
            {
                if (entry == null || entry.IsSubset || entry.Layout == null)
                {
                    if (entry != null) entry.LastRevisionSummary = "";
                    continue;
                }
                map.TryGetValue(entry.Layout.Name, out var summary);
                entry.LastRevisionSummary = summary ?? "";
            }
        }

        private string ResolveDrawingName(string layoutName)
        {
            if (_drawingNames != null
                && !string.IsNullOrWhiteSpace(layoutName)
                && _drawingNames.TryGetValue(layoutName, out var title)
                && !string.IsNullOrWhiteSpace(title))
                return title;
            return layoutName ?? "";
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
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 78));

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
                ColumnCount = 4,
                RowCount = 3,
            };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 105));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 65));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
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
            panel.SetColumnSpan(_pathBox, 3);

            var legend = BuildSubsetLevelLegend();
            panel.Controls.Add(legend, 2, 0);

            var importBtn = new Button
            {
                Text = "Import…",
                Dock = DockStyle.Fill,
                Margin = new Padding(6, 2, 0, 2),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(37, 99, 235),
                ForeColor = Color.White,
                UseVisualStyleBackColor = false,
                Cursor = Cursors.Hand,
            };
            importBtn.FlatAppearance.BorderSize = 0;
            importBtn.Click += (s, e) => StartImport();
            panel.Controls.Add(importBtn, 3, 0);

            var fieldHint = new Label
            {
                Text = "Title block: CurrentSheetNumber / CurrentSheetTitle / CurrentSheetRevisionNumber. "
                    + "Rev history: Table on layer TITLE_BLOCK_REV_TABLE; click 'Rev history' to edit.",
                Dock = DockStyle.Fill,
                ForeColor = Color.FromArgb(45, 95, 150),
                TextAlign = ContentAlignment.MiddleLeft,
            };
            panel.Controls.Add(fieldHint, 0, 2);
            panel.SetColumnSpan(fieldHint, 4);
            return panel;
        }

        private static Control BuildSubsetLevelLegend()
        {
            var wrap = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                AutoSize = false,
                Padding = new Padding(2, 0, 0, 0),
                Margin = Padding.Empty,
            };

            wrap.Controls.Add(new Label
            {
                Text = "Subset:",
                AutoSize = true,
                ForeColor = Color.DimGray,
                Margin = new Padding(0, 6, 6, 0),
            });

            AddLevelSwatch(wrap, 1, "Level1");
            AddLevelSwatch(wrap, 2, "Level2");
            AddLevelSwatch(wrap, 3, "Level3");
            AddLevelSwatch(wrap, 4, "Level4+");
            return wrap;
        }

        private static void AddLevelSwatch(FlowLayoutPanel wrap, int level, string label)
        {
            var chip = new Label
            {
                Text = "  " + label + "  ",
                AutoSize = true,
                Margin = new Padding(0, 2, 6, 0),
                Padding = new Padding(4, 2, 4, 2),
                BackColor = SubsetLevelColors.ForLevel(level),
                ForeColor = Color.FromArgb(30, 41, 59),
                BorderStyle = BorderStyle.FixedSingle,
                TextAlign = ContentAlignment.MiddleCenter,
            };
            wrap.Controls.Add(chip);
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

            // Cool slate — reorder
            var up = StyleButton(ToolbarButton("Move Up", () => MoveSelected(-1)),
                Color.FromArgb(71, 85, 105), Color.White);
            var down = StyleButton(ToolbarButton("Move Down", () => MoveSelected(1)),
                Color.FromArgb(100, 116, 139), Color.White);
            // Amber — renumber
            var renumber = StyleButton(ToolbarButton("Renumber 1…N", Renumber),
                Color.FromArgb(217, 119, 6), Color.White);
            // Rose — delete checked rows from table only (DST on Create/Update)
            var deleteSel = StyleButton(ToolbarButton("Delete sheet selected", DeleteSelectedRows),
                Color.FromArgb(255, 51, 51), Color.White);
            deleteSel.Width = 150;
            // Teal — sync from DST
            var refresh = StyleButton(ToolbarButton("Refresh from DST", () =>
            {
                var name = ReloadFromDstOrModel();
                if (!string.IsNullOrWhiteSpace(name)) _nameBox.Text = name;
                ValidateUi();
            }), Color.FromArgb(13, 148, 136), Color.White);
            refresh.Width = 130;
            // Green — export table
            var exportExcel = StyleButton(ToolbarButton("Export Excel…", ExportExcel),
                Color.FromArgb(22, 163, 74), Color.White);
            exportExcel.Width = 120;
            // Emerald — import editable columns only
            var importExcel = StyleButton(ToolbarButton("Import Excel…", ImportExcel),
                Color.FromArgb(5, 150, 105), Color.White);
            importExcel.Width = 120;

            bar.Controls.Add(up);
            bar.Controls.Add(down);
            bar.Controls.Add(renumber);
            bar.Controls.Add(refresh);
            bar.Controls.Add(exportExcel);
            bar.Controls.Add(importExcel);
            bar.Controls.Add(deleteSel);
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
                Name = "SelectCol",
                HeaderText = "",
                DataPropertyName = nameof(SheetSetEntry.Selected),
                Width = 36,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                ToolTipText = "Select rows to delete from this table (Create / Update DST to write)",
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
                Name = "TitleCol",
                HeaderText = "Drawing Name / Sheet Title",
                DataPropertyName = nameof(SheetSetEntry.Title),
                FillWeight = 30,
            });
            _grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Revision",
                DataPropertyName = nameof(SheetSetEntry.Revision),
                FillWeight = 12,
                ToolTipText = "CurrentSheetRevisionNumber (title block phiên bản)",
            });
            _grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "RevHistory",
                HeaderText = "Rev history (click)",
                DataPropertyName = nameof(SheetSetEntry.LastRevisionSummary),
                ReadOnly = true,
                FillWeight = 18,
                ToolTipText = "Click to edit revision history table on the layout",
            });
            _grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Layout",
                DataPropertyName = nameof(SheetSetEntry.DisplayLayout),
                ReadOnly = true,
                FillWeight = 16,
            });
            _grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "DWG",
                DataPropertyName = nameof(SheetSetEntry.DwgName),
                ReadOnly = true,
                FillWeight = 14,
            });

            _source.DataSource = _entries;
            _grid.DataSource = _source;
            _grid.CellValueChanged += (s, e) => ValidateUi();
            _grid.CellBeginEdit += (s, e) =>
            {
                if (e.RowIndex < 0 || e.RowIndex >= _entries.Count) return;
                if (_entries[e.RowIndex].IsSubset) e.Cancel = true;
            };
            _grid.CellClick += Grid_CellClick;
            _grid.CellFormatting += Grid_CellFormatting;
            return _grid;
        }

        private void Grid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            if (_grid.Columns[e.ColumnIndex].Name != "RevHistory") return;
            if (e.RowIndex >= _entries.Count) return;

            var entry = _entries[e.RowIndex];
            if (entry == null || entry.IsSubset || entry.Layout == null)
                return;

            OpenRevisionsDialog(entry.Layout.Name);
        }

        private void OpenRevisionsDialog(string layoutName)
        {
            var doc = AcadApp.DocumentManager?.MdiActiveDocument;
            if (doc == null)
            {
                MessageBox.Show(this, "No active document.", "Revisions",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var layoutNames = _entries
                .Where(x => x != null && !x.IsSubset && x.Layout != null)
                .Select(x => x.Layout.Name)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            using (var dlg = new RevisionsDialog(doc.Database, layoutNames, layoutName))
            {
                dlg.OnSaved = (savedLayout, summary) =>
                {
                    foreach (var entry in _entries)
                    {
                        if (entry?.Layout == null || entry.IsSubset) continue;
                        if (string.Equals(entry.Layout.Name, savedLayout, StringComparison.OrdinalIgnoreCase))
                            entry.LastRevisionSummary = summary ?? "";
                    }
                    _source.ResetBindings(false);
                    SetStatus($"Revisions saved for layout '{savedLayout}'.", false);
                };

                dlg.ShowDialog(this);
            }
        }

        private void Grid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= _entries.Count) return;
            var entry = _entries[e.RowIndex];
            var row = _grid.Rows[e.RowIndex];
            row.DefaultCellStyle.Padding = Padding.Empty;

            if (entry.IsSubset)
            {
                int level = Math.Max(1, entry.SubsetLevel);
                row.DefaultCellStyle.BackColor = SubsetLevelColors.ForLevel(level);
                row.DefaultCellStyle.ForeColor = Color.FromArgb(30, 41, 59);
                row.DefaultCellStyle.Font = new Font(_grid.Font, FontStyle.Bold);
                row.DefaultCellStyle.SelectionBackColor = SubsetLevelColors.SelectionForLevel(level);
                row.DefaultCellStyle.SelectionForeColor = Color.FromArgb(30, 41, 59);
            }
            else
            {
                row.DefaultCellStyle.BackColor = Color.White;
                row.DefaultCellStyle.ForeColor = _grid.DefaultCellStyle.ForeColor;
                row.DefaultCellStyle.Font = _grid.Font;
                row.DefaultCellStyle.SelectionBackColor = _grid.DefaultCellStyle.SelectionBackColor;
                row.DefaultCellStyle.SelectionForeColor = _grid.DefaultCellStyle.SelectionForeColor;
            }

            // Indent only Drawing Name / Sheet Title (subset + sheet under subset).
            var col = _grid.Columns[e.ColumnIndex];
            bool isTitleCol = col != null && (
                col.Name == "TitleCol"
                || col.DataPropertyName == nameof(SheetSetEntry.Title));
            if (isTitleCol)
            {
                int indentLevel = entry.IsSubset
                    ? Math.Max(0, entry.SubsetLevel - 1)
                    : Math.Max(0, entry.SubsetLevel);
                e.CellStyle.Padding = new Padding(indentLevel * 14, 0, 0, 0);
            }
            else
            {
                e.CellStyle.Padding = Padding.Empty;
            }
        }

        private void ApplyRowStyles()
        {
            if (_grid == null) return;
            _grid.Invalidate();
        }

        private Control BuildFooter()
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 2,
            };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            _openSsmChk = new CheckBox
            {
                Text = "Open Sheet Set Manager after Create / Update (continue subsets there)",
                AutoSize = true,
                Checked = true,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
            };
            panel.Controls.Add(_openSsmChk, 0, 0);
            panel.SetColumnSpan(_openSsmChk, 2);

            _status = new Label
            {
                Text = "",
                Dock = DockStyle.Fill,
                ForeColor = Color.DimGray,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true,
            };
            panel.Controls.Add(_status, 0, 1);

            var buttons = new FlowLayoutPanel
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Padding = new Padding(0, 4, 0, 0),
            };
            var close = new Button { Text = "Close", Width = 100, Height = 32, DialogResult = DialogResult.Cancel };
            _exportBtn = new Button { Text = "Export PDF…", Width = 120, Height = 32 };
            _createBtn = new Button { Text = "Create / Update DST", Width = 155, Height = 32 };
            StyleButton(close, Color.FromArgb(226, 232, 240), Color.FromArgb(51, 65, 85));
            StyleButton(_exportBtn, Color.FromArgb(37, 99, 235), Color.White);   // blue — PDF
            StyleButton(_createBtn, Color.FromArgb(180, 83, 9), Color.White);    // brown/orange — primary DST
            _exportBtn.Click += (s, e) => RequestExport();
            _createBtn.Click += (s, e) => CreateDst();
            buttons.Controls.Add(close);
            buttons.Controls.Add(_exportBtn);
            buttons.Controls.Add(_createBtn);
            panel.Controls.Add(buttons, 1, 1);
            return panel;
        }

        private void ExportExcel()
        {
            CommitGrid();
            ReassignSubsetMembership();
            RefreshOrders();
            if (_entries.Count == 0)
            {
                SetStatus("Nothing to export.", true);
                return;
            }

            string defaultName = Path.GetFileNameWithoutExtension(_pathBox.Text ?? "") ?? "SheetSet";
            string initialDir = null;
            try
            {
                initialDir = Path.GetDirectoryName(_pathBox.Text);
                if (!Directory.Exists(initialDir))
                    initialDir = PublishPaths.GetFolder(_dwgPath);
            }
            catch { }

            using (var sfd = new System.Windows.Forms.SaveFileDialog
            {
                Title = "Export Sheet Set table",
                Filter = "Excel workbook (*.xlsx)|*.xlsx|CSV file (*.csv)|*.csv",
                DefaultExt = "xlsx",
                AddExtension = true,
                FileName = defaultName + "_sheetset.xlsx",
                InitialDirectory = initialDir,
            })
            {
                if (sfd.ShowDialog(this) != DialogResult.OK) return;
                try
                {
                    SheetSetExporter.Export(sfd.FileName, _entries);
                    SetStatus($"Exported {_entries.Count} row(s): {sfd.FileName}", false);
                    var open = MessageBox.Show(
                        this,
                        $"Exported {_entries.Count} row(s).\n\n{sfd.FileName}\n\n"
                        + "Only edit SheetNumber, Title, Revision.\n"
                        + "Do not change Order, Kind, Subset, Layout, DWG "
                        + "(used only to match rows on Import).\n\n"
                        + "Open the file?",
                        "Export Excel",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Information);
                    if (open == DialogResult.Yes)
                    {
                        try
                        {
                            Process.Start(new ProcessStartInfo
                            {
                                FileName = sfd.FileName,
                                UseShellExecute = true,
                            });
                        }
                        catch (Exception openEx)
                        {
                            MessageBox.Show(this, "Could not open file: " + openEx.Message,
                                "Export Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }
                catch (Exception ex)
                {
                    SetStatus(ex.Message, true);
                    MessageBox.Show(this, "Export failed: " + ex.Message,
                        "Export Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void ImportExcel()
        {
            CommitGrid();
            if (!_entries.Any(e => e != null && !e.IsSubset))
            {
                SetStatus("No sheets in the table to update.", true);
                return;
            }

            string initialDir = null;
            try
            {
                initialDir = Path.GetDirectoryName(_pathBox.Text);
                if (!Directory.Exists(initialDir))
                    initialDir = PublishPaths.GetFolder(_dwgPath);
            }
            catch { }

            using (var ofd = new System.Windows.Forms.OpenFileDialog
            {
                Title = "Import Sheet Number / Title / Revision",
                Filter = "Excel or CSV (*.xlsx;*.csv)|*.xlsx;*.csv|Excel (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv",
                CheckFileExists = true,
                InitialDirectory = initialDir,
            })
            {
                if (ofd.ShowDialog(this) != DialogResult.OK) return;
                try
                {
                    var parsed = SheetSetImporter.ReadFile(ofd.FileName);
                    if (!parsed.Ok)
                    {
                        string err = parsed.Errors.Count > 0
                            ? string.Join("\n", parsed.Errors.Take(8))
                            : "Nothing to import.";
                        SetStatus(err, true);
                        MessageBox.Show(this, err, "Import Excel",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    var apply = SheetSetImporter.ApplyToEntries(_entries, parsed);
                    if (!apply.Ok)
                    {
                        string err = string.Join("\n", apply.Errors.Take(8));
                        SetStatus(err, true);
                        MessageBox.Show(this, err, "Import Excel",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    _source.ResetBindings(false);
                    ApplyRowStyles();
                    ValidateUi();
                    SetStatus(apply.Summary + "  " + ofd.FileName, false);

                    var detail = new StringBuilder();
                    detail.AppendLine(apply.Summary);
                    detail.AppendLine();
                    detail.AppendLine("Only SheetNumber / Title / Revision were applied.");
                    detail.AppendLine("Layout / Order / DWG were not changed.");
                    detail.AppendLine("Create / Update DST to write values into the .dst.");
                    if (apply.Warnings.Count > 0)
                    {
                        detail.AppendLine();
                        detail.AppendLine("Notes:");
                        foreach (var w in apply.Warnings.Take(12))
                            detail.AppendLine("• " + w);
                        if (apply.Warnings.Count > 12)
                            detail.AppendLine("• …");
                    }

                    MessageBox.Show(this, detail.ToString(), "Import Excel",
                        MessageBoxButtons.OK,
                        apply.Unmatched > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    SetStatus(ex.Message, true);
                    MessageBox.Show(this, "Import failed: " + ex.Message,
                        "Import Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void DeleteSelectedRows()
        {
            CommitGrid();
            var picked = _entries
                .Select((e, i) => (Entry: e, Index: i))
                .Where(x => x.Entry != null && x.Entry.Selected)
                .ToList();
            if (picked.Count == 0)
            {
                SetStatus("No rows selected.", true);
                return;
            }

            // Subsets with sheets still under them cannot be removed yet.
            var blocked = new List<string>();
            foreach (var item in picked.Where(x => x.Entry.IsSubset))
            {
                int sheets = CountSheetsUnderSubset(item.Index);
                if (sheets <= 0) continue;
                string name = item.Entry.Title ?? item.Entry.SubsetName ?? "Subset";
                blocked.Add(
                    $"Subset \"{name}\" still contains {sheets} sheet(s). "
                    + "Delete those sheets first, or uncheck the subset.");
            }
            if (blocked.Count > 0)
            {
                string msg = string.Join("\n\n", blocked.Take(8));
                if (blocked.Count > 8) msg += "\n\n…";
                SetStatus("Cannot delete non-empty subset(s).", true);
                MessageBox.Show(this, msg, "Delete selected",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int removeSheets = picked.Count(x => !x.Entry.IsSubset);
            int removeSubsets = picked.Count(x => x.Entry.IsSubset);
            var confirm = MessageBox.Show(
                this,
                $"Remove from this table only (not yet written to DST):\n\n"
                + $"• Sheets: {removeSheets}\n"
                + $"• Empty subsets: {removeSubsets}\n\n"
                + "Click Create / Update DST afterward to apply to Sheet Set Manager.",
                "Delete selected",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            if (confirm != DialogResult.Yes) return;

            // Remove high indices first.
            foreach (var item in picked.OrderByDescending(x => x.Index))
            {
                if (item.Index < 0 || item.Index >= _entries.Count) continue;
                _entries.RemoveAt(item.Index);
            }

            ReassignSubsetMembership();
            RefreshOrders();
            _source.ResetBindings(false);
            ApplyRowStyles();
            ValidateUi();
            SetStatus(
                $"Removed {removeSheets} sheet(s) and {removeSubsets} subset(s) from the table. "
                + "Create / Update DST to write changes.",
                false);
        }

        /// <summary>
        /// Sheets that belong under the subset at <paramref name="subsetIndex"/>
        /// until the next subset at the same or shallower level.
        /// </summary>
        private int CountSheetsUnderSubset(int subsetIndex)
        {
            if (subsetIndex < 0 || subsetIndex >= _entries.Count) return 0;
            var subset = _entries[subsetIndex];
            if (subset == null || !subset.IsSubset) return 0;

            int level = Math.Max(1, subset.SubsetLevel);
            int count = 0;
            for (int i = subsetIndex + 1; i < _entries.Count; i++)
            {
                var row = _entries[i];
                if (row == null) continue;
                if (row.IsSubset && Math.Max(1, row.SubsetLevel) <= level)
                    break;
                if (!row.IsSubset) count++;
            }
            return count;
        }

        private void MoveSelected(int delta)
        {
            CommitGrid();
            if (_grid.CurrentRow == null) return;
            int oldIndex = _grid.CurrentRow.Index;
            if (oldIndex < 0 || oldIndex >= _entries.Count) return;

            // Subset header rows are fixed anchors — only sheets may move.
            if (_entries[oldIndex].IsSubset)
            {
                SetStatus("Subset rows cannot be moved. Move sheets only.", true);
                return;
            }

            int newIndex = oldIndex + delta;
            if (newIndex < 0 || newIndex >= _entries.Count) return;

            var item = _entries[oldIndex];
            _entries.RaiseListChangedEvents = false;
            _entries.RemoveAt(oldIndex);
            _entries.Insert(newIndex, item);
            ReassignSubsetMembership();
            _entries.RaiseListChangedEvents = true;
            RefreshOrders();
            _source.ResetBindings(false);
            ApplyRowStyles();
            _grid.ClearSelection();
            _grid.Rows[newIndex].Selected = true;
            _grid.CurrentCell = _grid.Rows[newIndex].Cells[0];
            SetStatus("Sheet moved. Click Create / Update DST to write order/subsets to the DST.", false);
        }

        /// <summary>
        /// After moving sheets, membership follows the nearest subset header above
        /// with stack-aware nesting (SubsetLevel).
        /// </summary>
        private void ReassignSubsetMembership()
        {
            var stack = new List<(string Name, int Level)>();
            foreach (var entry in _entries)
            {
                if (entry.IsSubset)
                {
                    int level = entry.SubsetLevel > 0 ? entry.SubsetLevel : 1;
                    entry.SubsetLevel = level;
                    while (stack.Count > 0 && stack[stack.Count - 1].Level >= level)
                        stack.RemoveAt(stack.Count - 1);
                    var name = entry.Title ?? entry.SubsetName ?? "Subset";
                    entry.SubsetName = name;
                    entry.Title = name;
                    stack.Add((name, level));
                    continue;
                }

                if (stack.Count == 0)
                {
                    entry.SubsetName = null;
                    entry.SubsetLevel = 0;
                }
                else
                {
                    var top = stack[stack.Count - 1];
                    entry.SubsetName = top.Name;
                    entry.SubsetLevel = top.Level;
                }
            }
        }

        private void Renumber()
        {
            CommitGrid();
            int number = 1;
            foreach (var entry in _entries)
            {
                if (entry.IsSubset) continue;
                entry.SheetNumber = number.ToString();
                number++;
            }
            _source.ResetBindings(false);
            ApplyRowStyles();
            SetStatus($"Renumbered {number - 1} sheet(s).", false);
        }

        private void StartImport()
        {
            using (var dlg = new Form
            {
                Text = "Import",
                Width = 420,
                Height = 170,
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                ShowInTaskbar = false,
            })
            {
                var label = new Label
                {
                    Text = "What do you want to import?",
                    Left = 12,
                    Top = 16,
                    AutoSize = true,
                };
                var folderBtn = new Button
                {
                    Text = "Folder tree…",
                    Left = 12,
                    Top = 50,
                    Width = 180,
                    Height = 32,
                    DialogResult = DialogResult.Yes,
                };
                var filesBtn = new Button
                {
                    Text = "Drawing files (.dwg/.dwt)…",
                    Left = 208,
                    Top = 50,
                    Width = 180,
                    Height = 32,
                    DialogResult = DialogResult.No,
                };
                var cancel = new Button
                {
                    Text = "Cancel",
                    Left = 303,
                    Top = 98,
                    Width = 85,
                    DialogResult = DialogResult.Cancel,
                };
                dlg.Controls.Add(label);
                dlg.Controls.Add(folderBtn);
                dlg.Controls.Add(filesBtn);
                dlg.Controls.Add(cancel);
                dlg.CancelButton = cancel;

                var result = dlg.ShowDialog(this);
                if (result == DialogResult.Yes)
                    ImportFolderTree();
                else if (result == DialogResult.No)
                    ImportDrawingFiles();
            }
        }

        private void ImportFolderTree()
        {
            string folder;
            using (var fbd = new FolderBrowserDialog
            {
                Description = "Select the folder tree to import into the Sheet Set",
                ShowNewFolderButton = false,
            })
            {
                try
                {
                    var start = Path.GetDirectoryName(_dwgPath);
                    if (!string.IsNullOrWhiteSpace(start) && Directory.Exists(start))
                        fbd.SelectedPath = start;
                }
                catch { }

                if (fbd.ShowDialog(this) != DialogResult.OK) return;
                folder = fbd.SelectedPath;
            }

            FolderImportScanResult scan;
            UseWaitCursor = true;
            try
            {
                SetStatus("Scanning folder…", false);
                System.Windows.Forms.Application.DoEvents();
                scan = SheetSetFolderImport.Scan(folder);
            }
            finally
            {
                UseWaitCursor = false;
            }

            if (scan?.Root == null || !scan.Root.HasContent)
            {
                MessageBox.Show(this, scan?.Message ?? "Nothing to import.", "Import",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                SetStatus(scan?.Message ?? "Import cancelled.", true);
                return;
            }

            string parentPath = PickParentSubsetPath("Place imported folder tree under:");
            if (parentPath == null) return;

            int layouts = scan.Root.CountLayouts();
            int drawings = scan.Root.CountDrawings();
            int subsets = scan.Root.CountSubsets();
            var confirm = MessageBox.Show(
                this,
                $"{scan.Message}\n\n"
                + $"Parent: {(string.IsNullOrEmpty(parentPath) ? "(Sheet Set root)" : parentPath)}\n"
                + $"Will create/merge ~{subsets} subset(s), {drawings} drawing(s), {layouts} sheet(s).\n\n"
                + "Matching sheets (same DWG path + layout name) will be replaced.\nContinue?",
                "Import folder",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            if (confirm != DialogResult.Yes) return;

            FinishImportWrite(dst => SheetSetService.ImportFolderTree(
                dst, _nameBox.Text, parentPath, scan.Root));
        }

        private void ImportDrawingFiles()
        {
            string[] files;
            using (var ofd = new System.Windows.Forms.OpenFileDialog
            {
                Title = "Select drawing files to import as sheets",
                Filter = "AutoCAD drawings (*.dwg;*.dwt)|*.dwg;*.dwt|DWG (*.dwg)|*.dwg|DWT (*.dwt)|*.dwt",
                Multiselect = true,
                CheckFileExists = true,
            })
            {
                try
                {
                    var start = Path.GetDirectoryName(_dwgPath);
                    if (!string.IsNullOrWhiteSpace(start) && Directory.Exists(start))
                        ofd.InitialDirectory = start;
                }
                catch { }

                if (ofd.ShowDialog(this) != DialogResult.OK) return;
                files = ofd.FileNames;
            }

            FolderImportScanResult scan;
            UseWaitCursor = true;
            try
            {
                SetStatus("Reading drawings…", false);
                System.Windows.Forms.Application.DoEvents();
                scan = SheetSetFolderImport.ScanDrawingFiles(files);
            }
            finally
            {
                UseWaitCursor = false;
            }

            if (scan?.Root == null || !scan.Root.HasContent)
            {
                MessageBox.Show(this, scan?.Message ?? "Nothing to import.", "Import",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                SetStatus(scan?.Message ?? "Import cancelled.", true);
                return;
            }

            string parentPath = PickParentSubsetPath("Place imported sheets under:");
            if (parentPath == null) return;

            var confirm = MessageBox.Show(
                this,
                $"{scan.Message}\n\n"
                + $"Parent: {(string.IsNullOrEmpty(parentPath) ? "(Sheet Set root)" : parentPath)}\n"
                + "Selected files become sheets directly (no subset per file).\n\n"
                + "Matching sheets (same DWG path + layout name) will be replaced.\nContinue?",
                "Import drawings",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            if (confirm != DialogResult.Yes) return;

            FinishImportWrite(dst => SheetSetService.ImportDrawingFiles(
                dst, _nameBox.Text, parentPath, scan.Root.Drawings));
        }

        private void FinishImportWrite(Func<string, SheetSetService.FolderImportWriteResult> writeAction)
        {
            var dstPath = NormalizeDstPath(_pathBox.Text);
            if (string.IsNullOrWhiteSpace(dstPath))
            {
                SetStatus("DST path is missing.", true);
                return;
            }

            SheetSetService.FolderImportWriteResult write;
            UseWaitCursor = true;
            try
            {
                SetStatus("Writing Sheet Set…", false);
                System.Windows.Forms.Application.DoEvents();
                write = writeAction(dstPath);
            }
            finally
            {
                UseWaitCursor = false;
            }

            if (write == null || !write.Ok)
            {
                var err = write?.Message ?? "Import failed.";
                SetStatus(err, true);
                MessageBox.Show(this, err, "Import", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _pathBox.Text = dstPath;
            SaveDefaults();
            var name = ReloadFromDstOrModel();
            if (!string.IsNullOrWhiteSpace(name)) _nameBox.Text = name;
            ApplyRowStyles();
            SetStatus(write.Message, false);
            MessageBox.Show(this, write.Message, "Import", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Returns parent subset path ("" = root), or null if user cancelled.
        /// </summary>
        private string PickParentSubsetPath(string prompt)
        {
            var paths = CollectSubsetPaths();
            using (var dlg = new Form
            {
                Text = "Import parent subset",
                Width = 420,
                Height = 160,
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                ShowInTaskbar = false,
            })
            {
                var label = new Label
                {
                    Text = prompt ?? "Place import under:",
                    Left = 12,
                    Top = 14,
                    AutoSize = true,
                };
                var combo = new ComboBox
                {
                    Left = 12,
                    Top = 40,
                    Width = 380,
                    DropDownStyle = ComboBoxStyle.DropDownList,
                };
                combo.Items.Add("(Sheet Set root)");
                foreach (var p in paths)
                    combo.Items.Add(p);
                combo.SelectedIndex = 0;

                var ok = new Button
                {
                    Text = "OK",
                    DialogResult = DialogResult.OK,
                    Left = 216,
                    Top = 80,
                    Width = 85,
                };
                var cancel = new Button
                {
                    Text = "Cancel",
                    DialogResult = DialogResult.Cancel,
                    Left = 307,
                    Top = 80,
                    Width = 85,
                };
                dlg.Controls.Add(label);
                dlg.Controls.Add(combo);
                dlg.Controls.Add(ok);
                dlg.Controls.Add(cancel);
                dlg.AcceptButton = ok;
                dlg.CancelButton = cancel;

                if (dlg.ShowDialog(this) != DialogResult.OK)
                    return null;
                if (combo.SelectedIndex <= 0)
                    return "";
                return combo.SelectedItem as string ?? "";
            }
        }

        private List<string> CollectSubsetPaths()
        {
            var paths = new List<string>();
            var stack = new List<string>();
            foreach (var entry in _entries)
            {
                if (entry == null || !entry.IsSubset) continue;
                int level = Math.Max(1, entry.SubsetLevel);
                while (stack.Count >= level)
                    stack.RemoveAt(stack.Count - 1);
                var name = entry.Title ?? entry.SubsetName ?? "Subset";
                stack.Add(name);
                paths.Add(string.Join("/", stack));
            }
            return paths;
        }

        private void CreateDst()
        {
            CommitGrid();
            ReassignSubsetMembership();
            var writeList = GetOrderedEntriesForWrite();
            int sheetCount = writeList.Count(e => !e.IsSubset && e.Layout != null);
            var path = NormalizeDstPath(_pathBox.Text);
            if (string.IsNullOrWhiteSpace(path))
            {
                SetStatus("DST path is missing.", true);
                return;
            }

            if (sheetCount == 0)
            {
                var clear = MessageBox.Show(
                    this,
                    "The table has no sheets.\n\n"
                    + "Create / Update will write an empty sheet set "
                    + "(all sheets removed from the DST).\n\n"
                    + path + "\n\nContinue?",
                    "Create / Update Sheet Set",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                if (clear != DialogResult.Yes) return;
            }
            else if (File.Exists(path))
            {
                var answer = MessageBox.Show(
                    this,
                    "The DST file already exists and will be rebuilt from this table:\n\n"
                    + "• Sheet Number / Title\n"
                    + "• Subset grouping and sheet order (from Move Up/Down)\n\n"
                    + "Close the DST in Sheet Set Manager if Update fails.\n\n"
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
                SheetSetService.CreateOrReplace(path, _nameBox.Text, writeList);
                try
                {
                    // Same as typing RE — Editor.Regen() from this modal dialog often no-ops.
                    AcadApp.DocumentManager?.MdiActiveDocument
                        ?.SendStringToExecute("_.REGEN ", true, false, false);
                }
                catch { }
                _pathBox.Text = path;
                SaveDefaults();
                var name = ReloadFromDstOrModel();
                if (!string.IsNullOrWhiteSpace(name)) _nameBox.Text = name;
                SetStatus($"Wrote {sheetCount} sheet(s) to DST: {path}", false);

                string openNote = "";
                if (_openSsmChk != null && _openSsmChk.Checked)
                {
                    try { openNote = SheetSetLauncher.OpenForUser(path); }
                    catch (Exception openEx) { openNote = openEx.Message; }
                }

                MessageBox.Show(
                    this,
                    sheetCount == 0
                        ? $"Sheet set cleared (0 sheets).\n\nFile: {path}\n\n"
                          + (string.IsNullOrWhiteSpace(openNote)
                              ? "Tip: reopen Sheet Set Manager if the tree looks stale."
                              : openNote)
                        : $"Sheet set written successfully.\n\nSheets: {sheetCount}\nFile: {path}\n\n"
                          + "Table order and subsets were written to the DST.\n"
                          + "Save the DWG to persist its sheet-set association.\n\n"
                          + (string.IsNullOrWhiteSpace(openNote)
                              ? "Tip: reopen Sheet Set Manager if the tree looks stale."
                              : openNote),
                    "Create Sheet Set",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                SheetSetAutoLog.WriteException(
                    null, _dwgPath, "PLSHEETSET", "Create/Update DST failed", ex, _fixedDstPath);
                SetStatus(ex.Message, true);
                string logPath = SheetSetAutoLog.GetLogFilePath(_dwgPath, _fixedDstPath);
                MessageBox.Show(
                    this,
                    ex.Message
                    + "\n\nFull details written to:\n"
                    + logPath,
                    "Create Sheet Set",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                UseWaitCursor = false;
                ValidateUi();
            }
        }

        /// <summary>
        /// Full table order for DST write: subset headers + included sheets.
        /// </summary>
        private List<SheetSetEntry> GetOrderedEntriesForWrite()
        {
            return _entries
                .Where(x => x != null && (x.IsSubset || x.Layout != null))
                .OrderBy(x => x.Order)
                .ToList();
        }

        private void RequestExport()
        {
            CommitGrid();
            var selected = GetSelectedEntries();
            if (selected.Count == 0) return;
            ExportLayouts = selected
                .Where(x => !x.IsSubset && x.Layout != null)
                .Select(x => x.Layout)
                .ToList();
            if (ExportLayouts.Count == 0) return;
            SaveDefaults();
            DialogResult = DialogResult.OK;
            Close();
        }

        private List<SheetSetEntry> GetSelectedEntries()
        {
            return _entries
                .Where(x => x != null && !x.IsSubset && x.Layout != null)
                .OrderBy(x => x.Order)
                .ToList();
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
            int count = _entries.Count(x => !x.IsSubset && x.Layout != null);
            bool hasPath = !string.IsNullOrWhiteSpace(_pathBox.Text);
            // Allow Create/Update with 0 sheets to clear the DST; Export PDF still needs sheets.
            _createBtn.Enabled = hasPath;
            _exportBtn.Enabled = count > 0;
            if (_status != null && (_status.Text.StartsWith("Loaded") || _status.Text.StartsWith("No DST")
                || _status.Text.StartsWith("DST has") || string.IsNullOrWhiteSpace(_status.Text)
                || _status.Text.Contains("sheet(s)")))
            {
                if (string.IsNullOrWhiteSpace(_status.Text) || _status.Text.Contains("layout(s) selected")
                    || _status.Text.Contains("sheet(s) selected") || _status.Text.Contains("sheet(s)."))
                    _status.Text = $"{count} sheet(s).";
            }
        }

        private void ApplyFixedPath(string dstPath, string sheetSetName)
        {
            _pathBox.Text = dstPath ?? "";
            _nameBox.Text = sheetSetName ?? "";
            LoadOpenSsmPref();
        }

        private void LoadOpenSsmPref()
        {
            if (_openSsmChk == null) return;
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(RegKey))
                {
                    var raw = key?.GetValue("OpenSsmAfterCreate") as string;
                    // Default ON when unset.
                    _openSsmChk.Checked = !string.Equals(raw, "0", StringComparison.Ordinal);
                }
            }
            catch
            {
                _openSsmChk.Checked = true;
            }
        }

        private void SaveDefaults()
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(RegKey))
                {
                    key?.SetValue("LastDstPath", _pathBox.Text ?? "");
                    key?.SetValue("LastSheetSetName", _nameBox.Text ?? "");
                    if (_openSsmChk != null)
                        key?.SetValue("OpenSsmAfterCreate", _openSsmChk.Checked ? "1" : "0");
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
            var button = new Button
            {
                Text = text,
                Width = 105,
                Height = 30,
                Margin = new Padding(0, 0, 8, 0),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
            };
            button.FlatAppearance.BorderSize = 0;
            button.Click += (s, e) => action();
            return button;
        }

        private static Button StyleButton(Button button, Color back, Color fore)
        {
            if (button == null) return null;
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
            button.BackColor = back;
            button.ForeColor = fore;
            button.UseVisualStyleBackColor = false;
            button.Cursor = Cursors.Hand;
            // Slightly darker hover via FlatAppearance
            button.FlatAppearance.MouseOverBackColor = ControlPaint.Dark(back, 0.08f);
            button.FlatAppearance.MouseDownBackColor = ControlPaint.Dark(back, 0.15f);
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
