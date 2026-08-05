using System.Drawing;
using System.Windows.Forms;

namespace PrintLayoutAddin.UI
{
    public enum DuplicateAction { Skip, Overwrite, Rename, Abort }

    public class DuplicateLayoutDialog : Form
    {
        public DuplicateAction Action { get; private set; } = DuplicateAction.Abort;
        public bool ApplyToAll { get; private set; }

        public DuplicateLayoutDialog(string layoutName)
        {
            Text = "Layout already exists";
            Width = 440;
            Height = 230;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;

            var label = new Label
            {
                Text = $"Layout \"{layoutName}\" already exists. What would you like to do?",
                Dock = DockStyle.Top,
                Height = 40,
                Padding = new Padding(12, 12, 12, 0)
            };

            var applyAll = new CheckBox
            {
                Text = "Apply to all remaining duplicate layouts",
                Dock = DockStyle.Top,
                Height = 28,
                Padding = new Padding(12, 0, 12, 0)
            };

            var skipBtn = MakeBtn("Skip", DuplicateAction.Skip, applyAll);
            var overwriteBtn = MakeBtn("Overwrite", DuplicateAction.Overwrite, applyAll);
            var renameBtn = MakeBtn("Rename (_1, _2...)", DuplicateAction.Rename, applyAll);
            var abortBtn = MakeBtn("Abort", DuplicateAction.Abort, applyAll);

            var panel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.LeftToRight,
                Height = 50,
                Padding = new Padding(10)
            };
            panel.Controls.Add(skipBtn);
            panel.Controls.Add(overwriteBtn);
            panel.Controls.Add(renameBtn);
            panel.Controls.Add(abortBtn);

            Controls.Add(panel);
            Controls.Add(applyAll);
            Controls.Add(label);
        }

        private Button MakeBtn(string text, DuplicateAction action, CheckBox applyAll)
        {
            var btn = new Button { Text = text, Width = 90, Height = 30 };
            btn.Click += (s, e) =>
            {
                Action = action;
                ApplyToAll = applyAll.Checked;
                DialogResult = DialogResult.OK;
                Close();
            };
            return btn;
        }
    }
}
