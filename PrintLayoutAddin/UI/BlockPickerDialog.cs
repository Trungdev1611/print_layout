using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using PrintLayoutAddin.Core;

namespace PrintLayoutAddin.UI
{
    public class BlockPickerDialog : Form
    {
        private ListBox _list;
        private Button _okBtn;
        private Button _cancelBtn;
        public BlockChoice Selected { get; private set; }

        public BlockPickerDialog(IEnumerable<BlockChoice> choices, string preselectName = null)
        {
            Text = "Select source block / xref";
            Width = 460;
            Height = 420;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MinimizeBox = false;
            MaximizeBox = false;

            _list = new ListBox
            {
                Dock = DockStyle.Fill,
                IntegralHeight = false,
                Font = new Font("Consolas", 9f)
            };
            foreach (var c in choices) _list.Items.Add(c);

            if (!string.IsNullOrEmpty(preselectName))
            {
                for (int i = 0; i < _list.Items.Count; i++)
                {
                    if (((BlockChoice)_list.Items[i]).Name == preselectName)
                    {
                        _list.SelectedIndex = i;
                        break;
                    }
                }
            }
            if (_list.SelectedIndex < 0 && _list.Items.Count > 0) _list.SelectedIndex = 0;

            _list.DoubleClick += (s, e) => TryAccept();

            _okBtn = new Button { Text = "OK", DialogResult = DialogResult.OK, Width = 90 };
            _cancelBtn = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Width = 90 };
            _okBtn.Click += (s, e) => TryAccept();

            var btnPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft,
                Height = 40,
                Padding = new Padding(6)
            };
            btnPanel.Controls.Add(_cancelBtn);
            btnPanel.Controls.Add(_okBtn);

            Controls.Add(_list);
            Controls.Add(btnPanel);

            AcceptButton = _okBtn;
            CancelButton = _cancelBtn;
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
