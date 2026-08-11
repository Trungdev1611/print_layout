using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using PrintLayoutAddin.Core;

namespace PrintLayoutAddin.UI
{
    public class BlockPickerDialog : Form
    {
        private readonly List<BlockChoice> _all;
        private ListBox _list;
        private TextBox _filter;
        private Button _okBtn;
        private Button _cancelBtn;
        private readonly string _preselectName;

        public BlockChoice Selected { get; private set; }

        public BlockPickerDialog(IEnumerable<BlockChoice> choices, string preselectName = null)
            : this(choices, preselectName, "Select source block / xref")
        {
        }

        public BlockPickerDialog(
            IEnumerable<BlockChoice> choices,
            string preselectName,
            string title)
        {
            _all = (choices ?? Enumerable.Empty<BlockChoice>())
                .Where(c => c != null)
                .ToList();
            _preselectName = preselectName;

            Text = string.IsNullOrWhiteSpace(title) ? "Select source block / xref" : title;
            Width = 560;
            Height = 460;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MinimizeBox = false;
            MaximizeBox = false;

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(10),
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));

            var filterPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
            };
            filterPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 50));
            filterPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            filterPanel.Controls.Add(new Label
            {
                Text = "Filter",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
            }, 0, 0);
            _filter = new TextBox { Dock = DockStyle.Fill };
            _filter.TextChanged += (s, e) => ApplyFilter();
            filterPanel.Controls.Add(_filter, 1, 0);
            root.Controls.Add(filterPanel, 0, 0);

            _list = new ListBox
            {
                Dock = DockStyle.Fill,
                IntegralHeight = false,
                Font = new Font("Consolas", 9f),
            };
            _list.DoubleClick += (s, e) => TryAccept();
            root.Controls.Add(_list, 0, 1);

            _okBtn = new Button { Text = "OK", DialogResult = DialogResult.OK, Width = 90 };
            _cancelBtn = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Width = 90 };
            _okBtn.Click += (s, e) => TryAccept();

            var btnPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(0, 6, 0, 0),
            };
            btnPanel.Controls.Add(_cancelBtn);
            btnPanel.Controls.Add(_okBtn);
            root.Controls.Add(btnPanel, 0, 2);

            Controls.Add(root);
            AcceptButton = _okBtn;
            CancelButton = _cancelBtn;

            ApplyFilter();
        }

        private void ApplyFilter()
        {
            string q = (_filter?.Text ?? "").Trim();
            _list.BeginUpdate();
            try
            {
                _list.Items.Clear();
                IEnumerable<BlockChoice> view = _all;
                if (!string.IsNullOrEmpty(q))
                {
                    view = _all.Where(c =>
                        (c.Name ?? "").IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0);
                }
                foreach (var c in view)
                    _list.Items.Add(c);

                // Prefer preselect, else first item.
                int sel = -1;
                if (!string.IsNullOrEmpty(_preselectName))
                {
                    for (int i = 0; i < _list.Items.Count; i++)
                    {
                        if (string.Equals(
                                ((BlockChoice)_list.Items[i]).Name,
                                _preselectName,
                                StringComparison.OrdinalIgnoreCase))
                        {
                            sel = i;
                            break;
                        }
                    }
                }
                if (sel < 0 && _list.Items.Count > 0) sel = 0;
                _list.SelectedIndex = sel;
            }
            finally
            {
                _list.EndUpdate();
            }
        }

        private void TryAccept()
        {
            if (_list.SelectedItem is BlockChoice c)
            {
                Selected = c;
                DialogResult = DialogResult.OK;
                Close();
            }
        }
    }
}
