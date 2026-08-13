using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using PrintLayoutAddin.Core;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace PrintLayoutAddin.UI
{
    public class RevisionsDialog : Form
    {
        private readonly Database _db;
        private readonly List<string> _layoutNames;
        private readonly BindingList<RevisionItem> _items = new BindingList<RevisionItem>();
        private readonly BindingSource _source = new BindingSource();

        private ComboBox _layoutCombo;
        private DataGridView _grid;
        private Label _status;
        private Button _saveBtn;
        private bool _dirty;
        private string _loadedLayout;

        public string CurrentLayoutName => _layoutCombo?.SelectedItem as string ?? _loadedLayout;

        public RevisionsDialog(Database db, IEnumerable<string> layoutNames, string initialLayout)
        {
            _db = db;
            _layoutNames = (layoutNames ?? Enumerable.Empty<string>())
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                .ToList();

            Text = "Revisions";
            Width = 720;
            Height = 480;
            MinimumSize = new Size(560, 360);
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = true;

            BuildUi();

            foreach (var name in _layoutNames)
                _layoutCombo.Items.Add(name);

            string pick = initialLayout;
            if (string.IsNullOrWhiteSpace(pick) || !_layoutNames.Contains(pick, StringComparer.OrdinalIgnoreCase))
                pick = _layoutNames.FirstOrDefault();

            if (!string.IsNullOrWhiteSpace(pick))
            {
                SelectLayoutQuiet(pick);
                LoadLayout(pick, force: true);
            }
            else
            {
                SetStatus("No layouts available.", true);
                _saveBtn.Enabled = false;
            }
        }

        /// <summary>Latest revision summary after a successful save (for parent grid).</summary>
        public string LastSavedSummary { get; private set; }

        /// <summary>True if at least one Save succeeded during this dialog session.</summary>
        public bool SavedAny { get; private set; }

        /// <summary>Fired after a successful Save (layoutName, summary) so the parent can refresh immediately.</summary>
        public Action<string, string> OnSaved { get; set; }

        private void BuildUi()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 1,
                RowCount = 3,
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 56));

            var header = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
            };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));

            header.Controls.Add(new Label
            {
                Text = "Layout",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
            }, 0, 0);

            _layoutCombo = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
            };
            _layoutCombo.SelectedIndexChanged += LayoutCombo_SelectedIndexChanged;
            header.Controls.Add(_layoutCombo, 1, 0);

            var addBtn = MakeButton("Add", Color.FromArgb(22, 163, 74), () =>
            {
                EnsureFixedSlots();
                int emptyIdx = -1;
                for (int i = 0; i < _items.Count; i++)
                {
                    if (_items[i] == null || _items[i].IsEmpty)
                    {
                        emptyIdx = i;
                        break;
                    }
                }
                if (emptyIdx < 0)
                {
                    MessageBox.Show(
                        this,
                        $"All {RevisionTableService.DataRowSlots} revision slots are filled.\n"
                        + "Clear a row first, or raise revTableDataRows in config.json.",
                        "Revisions",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                var slot = _items[emptyIdx] ?? new RevisionItem();
                if (string.IsNullOrWhiteSpace(slot.Date))
                    slot.Date = DateTime.Now.ToString("dd/MM/yyyy");
                _items[emptyIdx] = slot;
                RebindGrid();
                if (emptyIdx < _grid.Rows.Count)
                {
                    _grid.ClearSelection();
                    _grid.Rows[emptyIdx].Selected = true;
                    _grid.CurrentCell = _grid.Rows[emptyIdx].Cells[0];
                    try { _grid.BeginEdit(true); } catch { }
                }
                _dirty = true;
            });
            var removeBtn = MakeButton("Clear", Color.FromArgb(185, 28, 28), () =>
            {
                if (_grid.CurrentRow == null || _grid.CurrentRow.IsNewRow) return;
                int i = _grid.CurrentRow.Index;
                if (i < 0 || i >= _items.Count) return;
                _items[i] = new RevisionItem();
                EnsureFixedSlots();
                RebindGrid();
                _dirty = true;
            });
            header.Controls.Add(addBtn, 2, 0);
            header.Controls.Add(removeBtn, 3, 0);
            root.Controls.Add(header, 0, 0);

            _grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
            };
            _grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Lần nộp",
                DataPropertyName = nameof(RevisionItem.RevNo),
                Width = 80,
            });
            _grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Nội dung điều chỉnh",
                DataPropertyName = nameof(RevisionItem.Description),
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            });
            _grid.Columns.Add(new CalendarColumn
            {
                HeaderText = "Ngày nộp",
                DataPropertyName = nameof(RevisionItem.Date),
                Width = 120,
            });
            _source.DataSource = _items;
            _grid.DataSource = _source;
            _grid.CellValueChanged += (s, e) => _dirty = true;
            // Do not CommitEdit on CurrentCellDirtyStateChanged — that ends text
            // editing after the first typed character.
            root.Controls.Add(_grid, 0, 1);

            var footer = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
            };
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            _status = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.DimGray,
                AutoEllipsis = true,
            };
            footer.Controls.Add(_status, 0, 0);

            var buttons = new FlowLayoutPanel
            {
                AutoSize = true,
                FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft,
                WrapContents = false,
            };
            var close = MakeButton("Close", Color.FromArgb(226, 232, 240), () =>
            {
                if (!ConfirmDiscardIfDirty()) return;
                DialogResult = SavedAny ? DialogResult.OK : DialogResult.Cancel;
                Close();
            });
            close.ForeColor = Color.FromArgb(51, 65, 85);
            var reload = MakeButton("Reload", Color.FromArgb(13, 148, 136), () =>
            {
                if (!ConfirmDiscardIfDirty()) return;
                LoadLayout(CurrentLayoutName, force: true);
            });
            _saveBtn = MakeButton("Save to layout", Color.FromArgb(180, 83, 9), SaveToLayout);
            _saveBtn.Width = 130;
            buttons.Controls.Add(close);
            buttons.Controls.Add(reload);
            buttons.Controls.Add(_saveBtn);
            footer.Controls.Add(buttons, 1, 0);
            root.Controls.Add(footer, 0, 2);

            Controls.Add(root);
        }

        private void LayoutCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            var next = _layoutCombo.SelectedItem as string;
            if (string.IsNullOrWhiteSpace(next)) return;
            if (string.Equals(next, _loadedLayout, StringComparison.OrdinalIgnoreCase)) return;
            if (!ConfirmDiscardIfDirty())
            {
                SelectLayoutQuiet(_loadedLayout);
                return;
            }
            LoadLayout(next, force: true);
        }

        private void SelectLayoutQuiet(string layoutName)
        {
            _layoutCombo.SelectedIndexChanged -= LayoutCombo_SelectedIndexChanged;
            try
            {
                for (int i = 0; i < _layoutCombo.Items.Count; i++)
                {
                    if (string.Equals(_layoutCombo.Items[i] as string, layoutName, StringComparison.OrdinalIgnoreCase))
                    {
                        _layoutCombo.SelectedIndex = i;
                        break;
                    }
                }
            }
            finally
            {
                _layoutCombo.SelectedIndexChanged += LayoutCombo_SelectedIndexChanged;
            }
        }

        private void LoadLayout(string layoutName, bool force)
        {
            if (string.IsNullOrWhiteSpace(layoutName)) return;
            if (!force && string.Equals(layoutName, _loadedLayout, StringComparison.OrdinalIgnoreCase))
                return;

            var read = RevisionTableService.ReadRevisionsFromLayout(_db, layoutName);
            _items.RaiseListChangedEvents = false;
            _items.Clear();
            foreach (var item in RevisionTableService.NormalizeToDataSlots(read.Items))
                _items.Add(item);
            EnsureFixedSlots();
            _items.RaiseListChangedEvents = true;
            RebindGrid();

            _loadedLayout = layoutName;
            _dirty = false;
            Text = "Revisions — " + layoutName;
            SetStatus(read.Message, read.Found && read.CadDataRowCount < RevisionTableService.DataRowSlots);
            _saveBtn.Enabled = true;
            LastSavedSummary = RevisionTableService.FormatSummary(_items);
        }

        /// <summary>Keep the grid locked to exactly DataRowSlots rows (including blanks).</summary>
        private void EnsureFixedSlots()
        {
            int n = RevisionTableService.DataRowSlots;
            while (_items.Count < n)
                _items.Add(new RevisionItem());
            while (_items.Count > n)
                _items.RemoveAt(_items.Count - 1);
        }

        private void RebindGrid()
        {
            _source.DataSource = null;
            _source.DataSource = _items;
            _grid.DataSource = null;
            _grid.DataSource = _source;
        }

        private void SaveToLayout()
        {
            try { _grid.EndEdit(); } catch { }
            _source.EndEdit();

            var layoutName = CurrentLayoutName;
            if (string.IsNullOrWhiteSpace(layoutName))
            {
                SetStatus("No layout selected.", true);
                return;
            }

            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                SetStatus("No active document.", true);
                return;
            }

            try
            {
                using (doc.LockDocument())
                {
                    if (!RevisionTableService.WriteRevisionsToLayout(
                            _db, layoutName, _items.ToList(), out var message))
                    {
                        SetStatus(message, true);
                        MessageBox.Show(this, message, "Revisions", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    try { doc.Editor.Regen(); } catch { }

                    _dirty = false;
                    SavedAny = true;
                    LastSavedSummary = RevisionTableService.FormatSummary(_items);
                    OnSaved?.Invoke(layoutName, LastSavedSummary ?? "");

                    // Re-read from CAD so the grid matches what was written (no manual Reload).
                    LoadLayout(layoutName, force: true);
                    SetStatus(message, false);
                }
            }
            catch (Exception ex)
            {
                SetStatus(ex.Message, true);
                MessageBox.Show(this, ex.Message, "Revisions", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool ConfirmDiscardIfDirty()
        {
            if (!_dirty) return true;
            var answer = MessageBox.Show(
                this,
                "Revision edits are not saved. Discard changes?",
                "Revisions",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            return answer == DialogResult.Yes;
        }

        private void SetStatus(string text, bool error)
        {
            _status.Text = text ?? "";
            _status.ForeColor = error ? Color.DarkRed : Color.DarkGreen;
        }

        private static Button MakeButton(string text, Color back, Action click)
        {
            var b = new Button
            {
                Text = text,
                Width = 100,
                Height = 30,
                Margin = new Padding(4, 0, 0, 0),
                FlatStyle = FlatStyle.Flat,
                BackColor = back,
                ForeColor = Color.White,
                UseVisualStyleBackColor = false,
                Cursor = Cursors.Hand,
            };
            b.FlatAppearance.BorderSize = 0;
            b.Click += (s, e) => click();
            return b;
        }
    }

    /// <summary>
    /// DataGridView date column: DateTimePicker (type + calendar), stores dd/MM/yyyy string.
    /// </summary>
    internal sealed class CalendarColumn : DataGridViewColumn
    {
        public CalendarColumn()
            : base(new CalendarCell())
        {
            SortMode = DataGridViewColumnSortMode.NotSortable;
        }

        public override DataGridViewCell CellTemplate
        {
            get => base.CellTemplate;
            set
            {
                if (value != null && !(value is CalendarCell))
                    throw new InvalidCastException("CalendarColumn cells must be CalendarCell.");
                base.CellTemplate = value;
            }
        }
    }

    internal sealed class CalendarCell : DataGridViewTextBoxCell
    {
        private static readonly string[] ParseFormats = { "dd/MM/yyyy", "dd-MM-yyyy", "d/M/yyyy", "d-M-yyyy" };

        public CalendarCell()
        {
            Style.Format = "dd/MM/yyyy";
        }

        public override Type EditType => typeof(CalendarEditingControl);
        public override Type ValueType => typeof(string);
        public override object DefaultNewRowValue => "";

        public override void InitializeEditingControl(
            int rowIndex, object initialFormattedValue, DataGridViewCellStyle dataGridViewCellStyle)
        {
            base.InitializeEditingControl(rowIndex, initialFormattedValue, dataGridViewCellStyle);
            if (!(DataGridView?.EditingControl is CalendarEditingControl ctl))
                return;

            string text = Convert.ToString(initialFormattedValue)?.Trim() ?? "";
            if (string.IsNullOrEmpty(text))
                text = Convert.ToString(Value)?.Trim() ?? "";

            if (TryParseDate(text, out var dt))
            {
                ctl.Value = Clamp(ctl, dt);
                ctl.Checked = true;
            }
            else
            {
                ctl.Value = Clamp(ctl, DateTime.Today);
                // Empty slot → unchecked so user can leave date blank, or check/pick.
                ctl.Checked = !string.IsNullOrWhiteSpace(text);
            }
        }

        internal static bool TryParseDate(string text, out DateTime dt)
        {
            return DateTime.TryParseExact(
                text ?? "",
                ParseFormats,
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None,
                out dt);
        }

        internal static string FormatDate(DateTime dt) => dt.ToString("dd/MM/yyyy");

        private static DateTime Clamp(DateTimePicker ctl, DateTime dt)
        {
            if (dt < ctl.MinDate) return ctl.MinDate;
            if (dt > ctl.MaxDate) return ctl.MaxDate;
            return dt;
        }
    }

    internal sealed class CalendarEditingControl : DateTimePicker, IDataGridViewEditingControl
    {
        private DataGridView _grid;
        private bool _valueChanged;
        private int _rowIndex;

        public CalendarEditingControl()
        {
            Format = DateTimePickerFormat.Custom;
            CustomFormat = "dd/MM/yyyy";
            ShowCheckBox = true; // unchecked ⇒ empty date string
        }

        public object EditingControlFormattedValue
        {
            get => Checked ? CalendarCell.FormatDate(Value) : "";
            set
            {
                var text = Convert.ToString(value)?.Trim() ?? "";
                if (CalendarCell.TryParseDate(text, out var dt))
                {
                    Value = dt;
                    Checked = true;
                }
                else if (string.IsNullOrWhiteSpace(text))
                {
                    Checked = false;
                }
            }
        }

        public object GetEditingControlFormattedValue(DataGridViewDataErrorContexts context)
            => EditingControlFormattedValue;

        public void ApplyCellStyleToEditingControl(DataGridViewCellStyle style)
        {
            Font = style.Font;
            CalendarForeColor = style.ForeColor;
            CalendarMonthBackground = style.BackColor;
        }

        public int EditingControlRowIndex
        {
            get => _rowIndex;
            set => _rowIndex = value;
        }

        public bool EditingControlWantsInputKey(Keys keyData, bool dataGridViewWantsInputKey)
        {
            switch (keyData & Keys.KeyCode)
            {
                case Keys.Left:
                case Keys.Up:
                case Keys.Down:
                case Keys.Right:
                case Keys.Home:
                case Keys.End:
                case Keys.PageDown:
                case Keys.PageUp:
                    return true;
                default:
                    return !dataGridViewWantsInputKey;
            }
        }

        public void PrepareEditingControlForEdit(bool selectAll) { }

        public bool RepositionEditingControlOnValueChange => false;

        public DataGridView EditingControlDataGridView
        {
            get => _grid;
            set => _grid = value;
        }

        public bool EditingControlValueChanged
        {
            get => _valueChanged;
            set => _valueChanged = value;
        }

        public Cursor EditingPanelCursor => base.Cursor;

        protected override void OnValueChanged(EventArgs eventargs)
        {
            _valueChanged = true;
            EditingControlDataGridView?.NotifyCurrentCellDirty(true);
            base.OnValueChanged(eventargs);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            // Toggling the checkbox also changes "has date".
            _valueChanged = true;
            EditingControlDataGridView?.NotifyCurrentCellDirty(true);
            base.OnMouseDown(e);
        }
    }
}
