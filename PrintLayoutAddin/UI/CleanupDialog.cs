using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using PrintLayoutAddin.Core;

namespace PrintLayoutAddin.UI
{
    public class CleanupDialog : Form
    {
        private CheckedListBox _list;
        private CheckBox _alsoLayouts;
        private Button _okBtn;
        private Button _cancelBtn;

        public List<CleanupCandidate> SelectedBlocks { get; private set; } = new List<CleanupCandidate>();
        public bool AlsoDeleteLayouts { get; private set; }

        public CleanupDialog(IEnumerable<CleanupCandidate> candidates)
        {
            Text = "Cleanup PLAUTO-generated frames";
            Width = 500;
            Height = 440;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MinimizeBox = false;
            MaximizeBox = false;

            var info = new Label
            {
                Text = "Select blocks to remove. All instances in ModelSpace will be erased and the block definition purged.",
                Dock = DockStyle.Top,
                Height = 40,
                Padding = new Padding(12, 12, 12, 4)
            };

            _list = new CheckedListBox
            {
                Dock = DockStyle.Fill,
                CheckOnClick = true,
                IntegralHeight = false,
                Font = new Font("Consolas", 9f)
            };
            foreach (var c in candidates) _list.Items.Add(c, true); // pre-check all

            _alsoLayouts = new CheckBox
            {
                Text = "Also delete layouts whose name matches the INNO-STT values on these frames",
                Dock = DockStyle.Bottom,
                Height = 28,
                Padding = new Padding(12, 4, 12, 0),
                Checked = false
            };

            _okBtn = new Button { Text = "Delete", DialogResult = DialogResult.OK, Width = 90 };
            _cancelBtn = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Width = 90 };
            _okBtn.Click += (s, e) => TryAccept();

            var btnPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft,
                Height = 42,
                Padding = new Padding(8)
            };
            btnPanel.Controls.Add(_cancelBtn);
            btnPanel.Controls.Add(_okBtn);

            Controls.Add(_list);
            Controls.Add(_alsoLayouts);
            Controls.Add(btnPanel);
            Controls.Add(info);

            AcceptButton = _okBtn;
            CancelButton = _cancelBtn;
        }

        private void TryAccept()
        {
            SelectedBlocks = _list.CheckedItems.Cast<CleanupCandidate>().ToList();
            AlsoDeleteLayouts = _alsoLayouts.Checked;
            if (SelectedBlocks.Count == 0)
            {
                MessageBox.Show(this, "No block selected.", "Cleanup",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                DialogResult = DialogResult.None;
                return;
            }
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
