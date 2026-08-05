using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Win32;
using PrintLayoutAddin.Core;

namespace PrintLayoutAddin.UI
{
    // -----------------------------------------------------------------
    // Upgraded numbering dialog supporting three modes:
    //   - Simple   : Prefix + seq + Suffix, Skip, Padding
    //   - Advanced : Prefix + Main + Sep + Sub (with per-main exceptions)
    //   - Import   : load codes from XLSX or CSV/TXT file
    //
    // Caller passes in the expected frame count (number of frames the
    // guide polyline will hit). Codes are generated and validated before
    // the dialog lets the user click OK.
    // -----------------------------------------------------------------

    public enum NumberingMode
    {
        Simple = 0,
        Advanced = 1,
        Import = 2
    }

    public class SttOptionsDialog : Form
    {
        private const string RegKey = @"Software\PrintLayoutAddin";
        private const string NumbRegKey = @"Software\PrintLayoutAddin\Numbering";

        // Expected number of codes (== frames polyline will number).
        // Passed in by caller; may be 0 when we couldn't precompute.
        private readonly int _expectedCount;

        // Outputs
        public List<string> Codes { get; private set; }
        public NumberingMode Mode { get; private set; }
        public bool AllowCountMismatch { get; private set; }

        // ---- Widgets ----
        private TabControl _tabs;

        // Simple tab
        private TextBox _sPrefix, _sSuffix, _sSkip;
        private NumericUpDown _sStart, _sStep, _sPadding, _sCount;

        // Advanced tab
        private TextBox _aPrefix, _aSep, _aDefaultSub;
        private NumericUpDown _aStart, _aEnd, _aStep, _aPadding;
        private DataGridView _aExcGrid;

        // Import tab
        private Label _iStatus;
        private Button _iImportBtn, _iExportBtn;
        private List<string> _importedCodes;
        private string _importedSourcePath;
        private List<string> _importedNotes;

        // Preview + validation + buttons
        private DataGridView _grid;
        private Label _counter;
        private Label _validation;
        private Button _previewBtn, _okBtn, _cancelBtn;
        private CheckBox _allowMismatchChk;

        public SttOptionsDialog(int expectedFrameCount)
        {
            _expectedCount = System.Math.Max(0, expectedFrameCount);

            Text = "Number frames";
            Width = 880;
            Height = 840;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            MinimumSize = new Size(760, 720);

            BuildLayout();
            LoadFromRegistry();

            // Auto-preview once on open so the grid isn't empty.
            try { DoPreview(showErrors: false); } catch { /* best effort */ }
        }

        // =============================================================
        // Layout
        // =============================================================
        private void BuildLayout()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5,
                Padding = new Padding(8),
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 420)); // tabs
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));  // preview row
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));  // validation
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));  // grid
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));  // buttons

            _tabs = new TabControl { Dock = DockStyle.Fill };
            _tabs.Controls.Add(BuildSimpleTab());
            _tabs.Controls.Add(BuildAdvancedTab());
            _tabs.Controls.Add(BuildImportTab());
            root.Controls.Add(_tabs, 0, 0);

            // Preview row
            var previewPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(0, 6, 0, 0),
            };
            _previewBtn = new Button
            {
                Text = "Preview",
                Width = 110,
                Height = 30,
                Margin = new Padding(0, 0, 10, 0),
            };
            _previewBtn.Click += (s, e) => DoPreview(showErrors: true);
            _counter = new Label
            {
                AutoSize = true,
                Margin = new Padding(10, 8, 0, 0),
                Text = _expectedCount > 0
                    ? $"Expected frames: {_expectedCount}"
                    : "Expected frames: unknown"
            };
            _allowMismatchChk = new CheckBox
            {
                Text = "Allow count mismatch",
                AutoSize = true,
                Margin = new Padding(20, 8, 0, 0),
            };
            previewPanel.Controls.Add(_previewBtn);
            previewPanel.Controls.Add(_counter);
            previewPanel.Controls.Add(_allowMismatchChk);
            root.Controls.Add(previewPanel, 0, 1);

            // Validation label
            _validation = new Label
            {
                Dock = DockStyle.Fill,
                AutoEllipsis = false,
                ForeColor = Color.DarkRed,
                Padding = new Padding(2),
                Text = "",
            };
            root.Controls.Add(_validation, 0, 2);

            // Grid
            _grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                RowHeadersVisible = false,
            };
            _grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "#",
                Width = 60,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                Name = "IdxCol"
            });
            _grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Code",
                Name = "CodeCol",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                FillWeight = 60,
            });
            _grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Note",
                Name = "NoteCol",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                FillWeight = 40,
            });
            root.Controls.Add(_grid, 0, 3);

            // Buttons
            var btnPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(0, 8, 0, 0),
            };
            _cancelBtn = new Button { Text = "Cancel", Width = 100, Height = 30, DialogResult = DialogResult.Cancel };
            _okBtn = new Button { Text = "OK", Width = 100, Height = 30 };
            _okBtn.Click += (s, e) => Accept();
            btnPanel.Controls.Add(_cancelBtn);
            btnPanel.Controls.Add(_okBtn);
            root.Controls.Add(btnPanel, 0, 4);

            Controls.Add(root);
            AcceptButton = _okBtn;
            CancelButton = _cancelBtn;
        }

        private TabPage BuildSimpleTab()
        {
            var page = new TabPage("Simple");
            var tbl = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 5,
                Padding = new Padding(10),
            };
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            for (int r = 0; r < 4; r++)
                tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            tbl.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            _sPrefix = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(3, 6, 3, 3) };
            _sSuffix = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(3, 6, 3, 3) };
            _sStart = Num(1, -100000, 1000000);
            _sStep = Num(1, -10000, 10000);
            _sSkip = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(3, 6, 3, 3) };
            _sPadding = Num(2, 0, 10);
            _sCount = Num(_expectedCount > 0 ? _expectedCount : 10, 0, 100000);

            AddPair(tbl, 0, 0, "Prefix", _sPrefix);
            AddPair(tbl, 0, 2, "Suffix", _sSuffix);
            AddPair(tbl, 1, 0, "Start", _sStart);
            AddPair(tbl, 1, 2, "Step", _sStep);
            AddPair(tbl, 2, 0, "Skip", _sSkip);
            AddPair(tbl, 2, 2, "Padding", _sPadding);
            AddPair(tbl, 3, 0, "Count", _sCount);

            var hint = new Label
            {
                Text = "Skip: comma-separated numbers to omit (e.g. 3,7).\r\nCount defaults to the frame count detected on the polyline.",
                Dock = DockStyle.Fill,
                ForeColor = Color.Gray,
                AutoSize = false,
                Margin = new Padding(3, 8, 3, 3),
            };
            tbl.Controls.Add(hint, 0, 4);
            tbl.SetColumnSpan(hint, 4);

            page.Controls.Add(tbl);
            return page;
        }

        private TabPage BuildAdvancedTab()
        {
            var page = new TabPage("Advanced");
            var outer = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(10),
            };
            outer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            outer.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));    // header
            outer.RowStyles.Add(new RowStyle(SizeType.Absolute, 150));   // fields (4 rows × 34)
            outer.RowStyles.Add(new RowStyle(SizeType.Percent, 100));    // exceptions — fills rest

            outer.Controls.Add(new Label
            {
                Text = "Main number range & suffix. Use Exceptions to define multiple suffixes per main.",
                Dock = DockStyle.Fill,
                ForeColor = Color.Gray,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoSize = false,
            }, 0, 0);

            var fields = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 4,
            };
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            for (int r = 0; r < 4; r++)
                fields.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));

            _aPrefix = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(3, 6, 3, 3) };
            _aStart = Num(4200, -100000, 1000000);
            _aEnd = Num(4210, -100000, 1000000);
            _aStep = Num(1, -10000, 10000);
            _aPadding = Num(4, 0, 10);
            _aSep = new TextBox { Dock = DockStyle.Fill, Text = ".", Margin = new Padding(3, 6, 3, 3) };
            _aDefaultSub = new TextBox { Dock = DockStyle.Fill, Text = "A", Margin = new Padding(3, 6, 3, 3) };

            AddPair(fields, 0, 0, "Prefix", _aPrefix);
            AddPair(fields, 0, 2, "Separator", _aSep);
            AddPair(fields, 1, 0, "Start", _aStart);
            AddPair(fields, 1, 2, "End", _aEnd);
            AddPair(fields, 2, 0, "Step", _aStep);
            AddPair(fields, 2, 2, "Padding", _aPadding);
            AddPair(fields, 3, 0, "Default Sub", _aDefaultSub);
            outer.Controls.Add(fields, 0, 1);

            // Exceptions — a 2-column grid is far easier to read/edit than a
            // free-form multiline textbox once there are many rules.
            var excWrapper = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Margin = new Padding(0, 6, 0, 0),
            };
            excWrapper.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            excWrapper.RowStyles.Add(new RowStyle(SizeType.Absolute, 22)); // title
            excWrapper.RowStyles.Add(new RowStyle(SizeType.Absolute, 34)); // toolbar
            excWrapper.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // grid

            excWrapper.Controls.Add(new Label
            {
                Text = "Exceptions   (one row per main number that needs custom suffixes)",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoSize = false,
                Font = new Font(FontFamily.GenericSansSerif, 9f, FontStyle.Bold),
            }, 0, 0);

            var excToolbar = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0),
            };
            var addRowBtn = new Button { Text = "+ Add row", Width = 90, Height = 26, Margin = new Padding(0, 2, 6, 2) };
            var delRowBtn = new Button { Text = "− Remove", Width = 90, Height = 26, Margin = new Padding(0, 2, 6, 2) };
            var clearBtn = new Button { Text = "Clear all", Width = 80, Height = 26, Margin = new Padding(0, 2, 6, 2) };
            addRowBtn.Click += (s, e) => { _aExcGrid.Rows.Add(); };
            delRowBtn.Click += (s, e) => RemoveSelectedExceptionRows();
            clearBtn.Click += (s, e) =>
            {
                _aExcGrid.Rows.Clear();
                try { DoPreview(showErrors: false); } catch { }
            };
            excToolbar.Controls.Add(addRowBtn);
            excToolbar.Controls.Add(delRowBtn);
            excToolbar.Controls.Add(clearBtn);
            excToolbar.Controls.Add(new Label
            {
                Text = "Suffixes: comma-separated (A,B,C). Empty rows are ignored.",
                AutoSize = true,
                ForeColor = Color.Gray,
                Margin = new Padding(8, 7, 0, 0),
            });
            excWrapper.Controls.Add(excToolbar, 0, 1);

            _aExcGrid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = true,
                AllowUserToDeleteRows = true,
                AllowUserToResizeRows = true,
                RowHeadersVisible = true,
                RowHeadersWidth = 46,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                EditMode = DataGridViewEditMode.EditOnEnter,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ColumnHeadersHeight = 32,
                Font = new Font(FontFamily.GenericSansSerif, 9.75f),
                BorderStyle = BorderStyle.FixedSingle,
                GridColor = Color.Gainsboro,
            };
            // Roomy rows so each rule is easy to click and read (like the preview grid).
            _aExcGrid.RowTemplate.Height = 30;
            _aExcGrid.DefaultCellStyle.Padding = new Padding(4, 2, 4, 2);
            _aExcGrid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 247, 250);
            _aExcGrid.ColumnHeadersDefaultCellStyle.Font =
                new Font(FontFamily.GenericSansSerif, 9.75f, FontStyle.Bold);
            _aExcGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Main number",
                Name = "ExcMain",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                Width = 170,
                DefaultCellStyle = new DataGridViewCellStyle { Padding = new Padding(6, 2, 6, 2) },
            });
            _aExcGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Suffixes (comma-separated, e.g. A,B,C)",
                Name = "ExcSubs",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                FillWeight = 100,
                DefaultCellStyle = new DataGridViewCellStyle { Padding = new Padding(6, 2, 6, 2) },
            });
            excWrapper.Controls.Add(_aExcGrid, 0, 2);
            outer.Controls.Add(excWrapper, 0, 2);

            page.Controls.Add(outer);
            return page;
        }

        private void RemoveSelectedExceptionRows()
        {
            // Collect distinct, non-new rows from the current selection.
            var toRemove = _aExcGrid.SelectedCells
                .Cast<DataGridViewCell>()
                .Select(c => c.OwningRow)
                .Where(r => r != null && !r.IsNewRow)
                .Distinct()
                .ToList();
            foreach (var r in toRemove)
                _aExcGrid.Rows.Remove(r);
            try { DoPreview(showErrors: false); } catch { }
        }

        // Serialize the exceptions grid back into the "main=A,B" line format
        // that ParseExceptionRules / the registry both understand.
        private string BuildExceptionsText()
        {
            var sb = new System.Text.StringBuilder();
            foreach (DataGridViewRow row in _aExcGrid.Rows)
            {
                if (row.IsNewRow) continue;
                var main = Convert.ToString(row.Cells["ExcMain"].Value)?.Trim() ?? "";
                var subs = Convert.ToString(row.Cells["ExcSubs"].Value)?.Trim() ?? "";
                if (main.Length == 0 && subs.Length == 0) continue;
                sb.AppendLine(main + "=" + subs);
            }
            return sb.ToString();
        }

        // Populate the grid from the persisted "main=A,B" text format.
        private void LoadExceptionsFromText(string text)
        {
            _aExcGrid.Rows.Clear();
            if (string.IsNullOrWhiteSpace(text)) return;
            var lines = text.Replace("\r\n", "\n").Split('\n');
            foreach (var raw in lines)
            {
                var line = raw.Trim();
                if (line.Length == 0 || line.StartsWith("#")) continue;
                int eq = line.IndexOf('=');
                string main = eq < 0 ? line : line.Substring(0, eq).Trim();
                string subs = eq < 0 ? "" : line.Substring(eq + 1).Trim();
                _aExcGrid.Rows.Add(main, subs);
            }
        }

        private TabPage BuildImportTab()
        {
            var page = new TabPage("Import");
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(10),
            };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 80));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            panel.Controls.Add(new Label
            {
                Dock = DockStyle.Fill,
                Text =
                    "Import frame numbers from XLSX (.xlsx) or CSV/TXT (.csv, .txt).\r\n" +
                    "Required column: FrameNumber (or 'Code', 'STT', 'INNO-STT').\r\n" +
                    "Optional columns: Order (sort key) and Note (ignored).",
                ForeColor = Color.DimGray,
            }, 0, 0);

            var btnRow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
            };
            _iExportBtn = new Button { Text = "Export Template…", Width = 150, Height = 30 };
            _iImportBtn = new Button { Text = "Import File…", Width = 130, Height = 30 };
            _iExportBtn.Click += OnExportTemplate;
            _iImportBtn.Click += OnImportFile;
            btnRow.Controls.Add(_iExportBtn);
            btnRow.Controls.Add(_iImportBtn);
            panel.Controls.Add(btnRow, 0, 1);

            _iStatus = new Label
            {
                Dock = DockStyle.Fill,
                AutoEllipsis = true,
                ForeColor = Color.DimGray,
                Text = "No file loaded.",
            };
            panel.Controls.Add(_iStatus, 0, 2);

            page.Controls.Add(panel);
            return page;
        }

        private static NumericUpDown Num(int value, int min, int max)
        {
            return new NumericUpDown
            {
                Minimum = min,
                Maximum = max,
                Value = System.Math.Max(min, System.Math.Min(max, value)),
                Dock = DockStyle.Fill,
            };
        }

        private static void AddPair(TableLayoutPanel tbl, int row, int col, string label, Control input)
        {
            tbl.Controls.Add(new Label
            {
                Text = label,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoSize = false,
            }, col, row);
            tbl.Controls.Add(input, col + 1, row);
        }

        // =============================================================
        // Preview / validation
        // =============================================================
        private void DoPreview(bool showErrors)
        {
            List<string> codes = null;
            List<string> notes = null;
            var errors = new List<string>();
            var warnings = new List<string>();
            var mode = (NumberingMode)_tabs.SelectedIndex;

            try
            {
                switch (mode)
                {
                    case NumberingMode.Simple:
                        codes = GenerateSimpleFromUI(errors);
                        break;
                    case NumberingMode.Advanced:
                        codes = GenerateAdvancedFromUI(errors);
                        break;
                    case NumberingMode.Import:
                        codes = GetImportedCodes(errors, warnings, out notes);
                        break;
                }
            }
            catch (Exception ex)
            {
                errors.Add(ex.Message);
            }

            // Validate against expected count + duplicates
            if (codes != null)
            {
                var vr = NumberGenerator.ValidateNumberList(
                    codes,
                    expectedCount: _expectedCount > 0 ? (int?)_expectedCount : null,
                    allowDuplicates: false);
                // If user allows mismatch, treat count-mismatch as warning.
                if (_allowMismatchChk.Checked)
                {
                    var countErr = vr.Errors.FirstOrDefault(e => e.StartsWith("Count mismatch"));
                    if (countErr != null)
                    {
                        vr.Errors.Remove(countErr);
                        vr.Warnings.Add(countErr);
                    }
                }
                errors.AddRange(vr.Errors);
                warnings.AddRange(vr.Warnings);
            }

            // Paint grid
            _grid.SuspendLayout();
            _grid.Rows.Clear();
            if (codes != null)
            {
                for (int i = 0; i < codes.Count; i++)
                {
                    var note = (notes != null && i < notes.Count) ? notes[i] : "";
                    _grid.Rows.Add((i + 1).ToString(), codes[i] ?? "", note ?? "");
                }
            }
            _grid.ResumeLayout();

            _counter.Text = (_expectedCount > 0
                ? $"Expected frames: {_expectedCount}"
                : "Expected frames: unknown")
                + $"   |   Codes: {(codes?.Count ?? 0)}";

            if (errors.Count == 0 && warnings.Count == 0)
            {
                _validation.ForeColor = Color.DarkGreen;
                _validation.Text = codes != null && codes.Count > 0 ? "Validation OK." : "";
            }
            else
            {
                _validation.ForeColor = errors.Count > 0 ? Color.DarkRed : Color.DarkOrange;
                var text = "";
                if (errors.Count > 0) text += "ERROR: " + string.Join(" | ", errors);
                if (warnings.Count > 0)
                {
                    if (text.Length > 0) text += "   ";
                    text += "WARN: " + string.Join(" | ", warnings);
                }
                _validation.Text = text;
            }

            _okBtn.Enabled = codes != null && codes.Count > 0 && errors.Count == 0;
            // expose latest preview so OK can take it directly
            _cachedCodes = codes;
        }

        private List<string> _cachedCodes;

        private List<string> GenerateSimpleFromUI(List<string> errors)
        {
            var p = new SimpleGenerationParams
            {
                Prefix = _sPrefix.Text ?? "",
                Suffix = _sSuffix.Text ?? "",
                Start = (int)_sStart.Value,
                Step = (int)_sStep.Value,
                Padding = (int)_sPadding.Value,
                Skip = NumberGenerator.ParseSkipList(_sSkip.Text),
                Count = (int)_sCount.Value,
            };
            if (p.Step == 0) { errors.Add("Step must not be zero."); return null; }
            if (p.Count < 0) { errors.Add("Count must be >= 0."); return null; }
            return NumberGenerator.GenerateSimple(p);
        }

        private List<string> GenerateAdvancedFromUI(List<string> errors)
        {
            var p = new AdvancedGenerationParams
            {
                Prefix = _aPrefix.Text ?? "",
                Start = (int)_aStart.Value,
                End = (int)_aEnd.Value,
                Step = (int)_aStep.Value,
                Padding = (int)_aPadding.Value,
                Separator = _aSep.Text ?? ".",
                DefaultSubSuffix = _aDefaultSub.Text ?? "A",
            };
            if (p.Step == 0) { errors.Add("Step must not be zero."); return null; }

            // parse exceptions
            var excDict = NumberGenerator.ParseExceptionRules(BuildExceptionsText(), out var parseErrors);
            if (parseErrors != null && parseErrors.Count > 0)
            {
                errors.AddRange(parseErrors);
                return null;
            }
            p.Exceptions = excDict;
            return NumberGenerator.GenerateAdvanced(p);
        }

        private List<string> GetImportedCodes(List<string> errors, List<string> warnings, out List<string> notes)
        {
            notes = null;
            if (_importedCodes == null)
            {
                errors.Add("No file imported yet. Click 'Import File…' to load codes.");
                return null;
            }
            notes = _importedNotes;
            return new List<string>(_importedCodes);
        }

        private void OnExportTemplate(object sender, EventArgs e)
        {
            using (var sfd = new System.Windows.Forms.SaveFileDialog
            {
                Title = "Export numbering template",
                Filter = "Excel workbook (*.xlsx)|*.xlsx|CSV file (*.csv)|*.csv",
                FileName = "numbering-template",
            })
            {
                if (sfd.ShowDialog(this) != DialogResult.OK) return;
                int rowCount = _expectedCount > 0 ? _expectedCount : 20;
                try
                {
                    string ext = (Path.GetExtension(sfd.FileName) ?? "").ToLowerInvariant();
                    if (ext == ".xlsx") NumberImporter.ExportTemplateXlsx(sfd.FileName, rowCount);
                    else NumberImporter.ExportTemplateCsv(sfd.FileName, rowCount);
                    MessageBox.Show(this,
                        $"Template saved:\n{sfd.FileName}\n\n" +
                        $"Fill in the FrameNumber column ({rowCount} rows) then click Import File.",
                        "Export template", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Export failed: " + ex.Message,
                        "Export template", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void OnImportFile(object sender, EventArgs e)
        {
            string lastDir = null;
            try
            {
                using (var k = Registry.CurrentUser.OpenSubKey(NumbRegKey))
                    lastDir = k?.GetValue("LastImportDir") as string;
            }
            catch { }

            using (var ofd = new System.Windows.Forms.OpenFileDialog
            {
                Title = "Import frame numbers",
                Filter = "All supported (*.xlsx;*.csv;*.txt)|*.xlsx;*.csv;*.txt|Excel (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv|Text (*.txt)|*.txt",
                InitialDirectory = Directory.Exists(lastDir ?? "") ? lastDir : null,
            })
            {
                if (ofd.ShowDialog(this) != DialogResult.OK) return;

                try
                {
                    var import = NumberImporter.ImportFromFile(ofd.FileName);
                    if (import.Errors.Count > 0)
                    {
                        MessageBox.Show(this,
                            "Import failed:\n\n" + string.Join("\n", import.Errors),
                            "Import", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    _importedCodes = import.Rows.Select(r => r.FrameNumber).ToList();
                    _importedNotes = import.Rows.Select(r => r.Note ?? "").ToList();
                    _importedSourcePath = ofd.FileName;

                    var sb = new System.Text.StringBuilder();
                    sb.AppendLine($"File: {Path.GetFileName(ofd.FileName)}");
                    sb.Append($"Rows: {_importedCodes.Count}");
                    if (import.HasOrderColumn) sb.Append("  |  sorted by Order");
                    if (import.HasNoteColumn) sb.Append("  |  Note column detected");
                    if (import.Warnings.Count > 0)
                        sb.AppendLine().Append("Warnings: " + string.Join(" | ", import.Warnings));
                    _iStatus.Text = sb.ToString();
                    _iStatus.ForeColor = Color.DarkGreen;

                    try
                    {
                        using (var k = Registry.CurrentUser.CreateSubKey(NumbRegKey))
                            k?.SetValue("LastImportDir", Path.GetDirectoryName(ofd.FileName) ?? "");
                    }
                    catch { }

                    _tabs.SelectedIndex = (int)NumberingMode.Import;
                    DoPreview(showErrors: true);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Import failed: " + ex.Message,
                        "Import", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // =============================================================
        // Accept
        // =============================================================
        private void Accept()
        {
            // Final re-preview to ensure whatever is in the fields
            // matches what will be applied (user may have edited a box
            // without clicking Preview).
            DoPreview(showErrors: true);
            if (_cachedCodes == null || _cachedCodes.Count == 0)
            {
                MessageBox.Show(this, "No codes to apply. Fix validation errors first.",
                    "Number frames", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            // OK button is only enabled if validation passed, but double-check:
            if (!_okBtn.Enabled)
            {
                MessageBox.Show(this, "Please resolve validation errors before applying.",
                    "Number frames", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Codes = new List<string>(_cachedCodes);
            Mode = (NumberingMode)_tabs.SelectedIndex;
            AllowCountMismatch = _allowMismatchChk.Checked;
            SaveToRegistry();
            DialogResult = DialogResult.OK;
            Close();
        }

        // =============================================================
        // Registry persistence
        // =============================================================
        private void LoadFromRegistry()
        {
            try
            {
                // Legacy keys (pre-v1.4) for Simple tab
                using (var k = Registry.CurrentUser.OpenSubKey(RegKey))
                {
                    if (k != null)
                    {
                        _sPrefix.Text = (k.GetValue("Prefix") as string) ?? _sPrefix.Text;
                        _sSuffix.Text = (k.GetValue("Suffix") as string) ?? _sSuffix.Text;
                        _sSkip.Text = (k.GetValue("Skip") as string) ?? _sSkip.Text;
                        _sStart.Value = ToDecimal(k.GetValue("Start"), _sStart.Value);
                        _sStep.Value = ToDecimal(k.GetValue("Step"), _sStep.Value);
                        _sPadding.Value = ToDecimal(k.GetValue("Padding"), _sPadding.Value);
                    }
                }

                using (var k = Registry.CurrentUser.OpenSubKey(NumbRegKey))
                {
                    if (k == null) return;

                    // Mode
                    var modeS = k.GetValue("Mode") as string;
                    if (!string.IsNullOrEmpty(modeS) &&
                        Enum.TryParse<NumberingMode>(modeS, out var mode))
                    {
                        _tabs.SelectedIndex = (int)mode;
                    }

                    _allowMismatchChk.Checked = (k.GetValue("AllowMismatch") as string) == "1";

                    // Advanced
                    _aPrefix.Text = (k.GetValue("APrefix") as string) ?? _aPrefix.Text;
                    _aSep.Text = (k.GetValue("ASep") as string) ?? _aSep.Text;
                    _aDefaultSub.Text = (k.GetValue("ADefaultSub") as string) ?? _aDefaultSub.Text;
                    _aStart.Value = ToDecimal(k.GetValue("AStart"), _aStart.Value);
                    _aEnd.Value = ToDecimal(k.GetValue("AEnd"), _aEnd.Value);
                    _aStep.Value = ToDecimal(k.GetValue("AStep"), _aStep.Value);
                    _aPadding.Value = ToDecimal(k.GetValue("APadding"), _aPadding.Value);
                    LoadExceptionsFromText((k.GetValue("AExceptions") as string) ?? "");
                }
            }
            catch { }
        }

        private void SaveToRegistry()
        {
            try
            {
                // Legacy Simple keys (shared)
                using (var k = Registry.CurrentUser.CreateSubKey(RegKey))
                {
                    if (k != null)
                    {
                        k.SetValue("Prefix", _sPrefix.Text ?? "");
                        k.SetValue("Suffix", _sSuffix.Text ?? "");
                        k.SetValue("Skip", _sSkip.Text ?? "");
                        k.SetValue("Start", (int)_sStart.Value);
                        k.SetValue("Step", (int)_sStep.Value);
                        k.SetValue("Padding", (int)_sPadding.Value);
                    }
                }
                using (var k = Registry.CurrentUser.CreateSubKey(NumbRegKey))
                {
                    if (k == null) return;
                    k.SetValue("Mode", ((NumberingMode)_tabs.SelectedIndex).ToString());
                    k.SetValue("AllowMismatch", _allowMismatchChk.Checked ? "1" : "0");
                    k.SetValue("APrefix", _aPrefix.Text ?? "");
                    k.SetValue("ASep", _aSep.Text ?? "");
                    k.SetValue("ADefaultSub", _aDefaultSub.Text ?? "");
                    k.SetValue("AStart", (int)_aStart.Value);
                    k.SetValue("AEnd", (int)_aEnd.Value);
                    k.SetValue("AStep", (int)_aStep.Value);
                    k.SetValue("APadding", (int)_aPadding.Value);
                    k.SetValue("AExceptions", BuildExceptionsText());
                }
            }
            catch { }
        }

        private static decimal ToDecimal(object v, decimal def)
        {
            if (v == null) return def;
            if (decimal.TryParse(v.ToString(), out var d)) return d;
            return def;
        }
    }
}
