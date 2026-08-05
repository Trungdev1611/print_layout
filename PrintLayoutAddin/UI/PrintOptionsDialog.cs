using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.AutoCAD.DatabaseServices;
using Microsoft.Win32;
using PrintLayoutAddin.Core;

namespace PrintLayoutAddin.UI
{
    public class PrintOptionsDialog : Form
    {
        private const string RegKey = @"Software\PrintLayoutAddin";
        private const string UseLayoutStyle = "(use layout setting)";

        private readonly Database _db;
        private readonly List<PrintableLayout> _layouts;
        private readonly string _templateLayoutName;
        private readonly string _defaultPdfPath;

        private CheckedListBox _layoutList;
        private ComboBox _deviceCombo;
        private ComboBox _mediaCombo;
        private ComboBox _styleCombo;
        private ComboBox _orientationCombo;
        private ComboBox _outputModeCombo;
        private ComboBox _plotAreaCombo;
        private NumericUpDown _copiesUpDown;
        private CheckBox _plotToFileChk;
        private CheckBox _scaleToFitChk;
        private CheckBox _openAfterChk;
        private Label _pathLabel;
        private TextBox _outputPathBox;
        private Button _okBtn;

        public PrintJobOptions Options { get; private set; }

        public PrintOptionsDialog(
            Database db,
            IEnumerable<PrintableLayout> layouts,
            string templateLayoutName,
            string defaultPdfPath)
        {
            _db = db;
            _layouts = layouts?.ToList() ?? new List<PrintableLayout>();
            _templateLayoutName = templateLayoutName ?? "";
            _defaultPdfPath = defaultPdfPath;

            Text = "Print / Export PDF";
            Width = 760;
            Height = 610;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MinimizeBox = false;
            MaximizeBox = false;

            BuildUi();
            LoadValues();
            ValidateUi();
        }

        private void BuildUi()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 2,
                Padding = new Padding(12)
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 260));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));

            var left = BuildLayoutPanel();
            var right = BuildSettingsPanel();
            var buttons = BuildButtonsPanel();

            root.Controls.Add(left, 0, 0);
            root.Controls.Add(right, 1, 0);
            root.SetColumnSpan(buttons, 2);
            root.Controls.Add(buttons, 0, 1);
            Controls.Add(root);
        }

        private Control BuildLayoutPanel()
        {
            var panel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 0, 10, 0) };
            var title = new Label
            {
                Text = "Layouts",
                Dock = DockStyle.Top,
                Height = 24,
                Font = new System.Drawing.Font(Font, FontStyle.Bold)
            };

            _layoutList = new CheckedListBox
            {
                Dock = DockStyle.Fill,
                CheckOnClick = false,
                IntegralHeight = false,
                Font = new System.Drawing.Font("Consolas", 9f)
            };
            _layoutList.ItemCheck += (s, e) =>
            {
                if (_okBtn == null) return;
                if (IsHandleCreated) BeginInvoke((Action)ValidateUi);
                else ValidateUi();
            };

            var savedOrder = ReadMultiString("PrintLayoutOrder");
            var savedChecked = ReadMultiString("PrintLayoutChecked");
            bool anyChecked = false;
            foreach (var layout in ApplySavedOrder(_layouts, savedOrder))
            {
                bool check = savedChecked != null
                    ? savedChecked.Contains(layout.Name, StringComparer.OrdinalIgnoreCase)
                    : !string.Equals(layout.Name, _templateLayoutName, StringComparison.OrdinalIgnoreCase);
                _layoutList.Items.Add(layout, check);
                anyChecked |= check;
            }

            // Remembered selection no longer matches any current layout -> fall back to the default.
            if (savedChecked != null && !anyChecked)
            {
                for (int i = 0; i < _layoutList.Items.Count; i++)
                {
                    var l = (PrintableLayout)_layoutList.Items[i];
                    _layoutList.SetItemChecked(i, !string.Equals(l.Name, _templateLayoutName, StringComparison.OrdinalIgnoreCase));
                }
            }

            var row = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                Height = 38,
                Padding = new Padding(0, 6, 0, 0)
            };
            var allBtn = new Button { Text = "All", Width = 52, Height = 28 };
            var noneBtn = new Button { Text = "None", Width = 56, Height = 28 };
            var upBtn = new Button { Text = "Up", Width = 48, Height = 28 };
            var downBtn = new Button { Text = "Down", Width = 58, Height = 28 };
            allBtn.Click += (s, e) => SetAllLayouts(true);
            noneBtn.Click += (s, e) => SetAllLayouts(false);
            upBtn.Click += (s, e) => MoveSelectedLayout(-1);
            downBtn.Click += (s, e) => MoveSelectedLayout(1);
            row.Controls.Add(allBtn);
            row.Controls.Add(noneBtn);
            row.Controls.Add(upBtn);
            row.Controls.Add(downBtn);

            panel.Controls.Add(_layoutList);
            panel.Controls.Add(row);
            panel.Controls.Add(title);
            return panel;
        }

        private Control BuildSettingsPanel()
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 12
            };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 88));

            AddLabel(panel, "Printer / PDF", 0);
            _deviceCombo = MakeCombo();
            _deviceCombo.SelectedIndexChanged += (s, e) =>
            {
                UpdateMediaList(null);
                if (SelectedDeviceLooksLikePdf()) _plotToFileChk.Checked = true;
                ValidateUi();
            };
            panel.Controls.Add(_deviceCombo, 1, 0);
            panel.SetColumnSpan(_deviceCombo, 2);

            AddLabel(panel, "Paper size", 1);
            _mediaCombo = MakeCombo();
            _mediaCombo.SelectedIndexChanged += (s, e) => ValidateUi();
            panel.Controls.Add(_mediaCombo, 1, 1);
            panel.SetColumnSpan(_mediaCombo, 2);

            AddLabel(panel, "Plot style", 2);
            _styleCombo = MakeCombo();
            panel.Controls.Add(_styleCombo, 1, 2);
            panel.SetColumnSpan(_styleCombo, 2);

            AddLabel(panel, "Orientation", 3);
            _orientationCombo = MakeCombo();
            _orientationCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _orientationCombo.Items.Add("Use layout setting");
            _orientationCombo.Items.Add("Portrait");
            _orientationCombo.Items.Add("Landscape");
            _orientationCombo.SelectedIndex = 0;
            panel.Controls.Add(_orientationCombo, 1, 3);
            panel.SetColumnSpan(_orientationCombo, 2);

            AddLabel(panel, "Plot area", 4);
            _plotAreaCombo = MakeCombo();
            _plotAreaCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _plotAreaCombo.Items.Add("Layout page");
            _plotAreaCombo.Items.Add("Extents centered");
            _plotAreaCombo.SelectedIndex = 0;
            _plotAreaCombo.SelectedIndexChanged += (s, e) =>
            {
                UpdateScaleToFitState();
            };
            panel.Controls.Add(_plotAreaCombo, 1, 4);
            panel.SetColumnSpan(_plotAreaCombo, 2);

            AddLabel(panel, "Copies", 5);
            _copiesUpDown = new NumericUpDown
            {
                Dock = DockStyle.Left,
                Width = 90,
                Minimum = 1,
                Maximum = 99,
                Value = 1
            };
            panel.Controls.Add(_copiesUpDown, 1, 5);

            AddLabel(panel, "Output", 6);
            _plotToFileChk = new CheckBox
            {
                Text = "Plot to PDF/file",
                Dock = DockStyle.Fill,
                Checked = true
            };
            _plotToFileChk.CheckedChanged += (s, e) =>
            {
                _outputPathBox.Enabled = _plotToFileChk.Checked;
                _outputModeCombo.Enabled = _plotToFileChk.Checked;
                UpdatePathLabel();
                UpdateCopiesState();
                ValidateUi();
            };
            panel.Controls.Add(_plotToFileChk, 1, 6);
            panel.SetColumnSpan(_plotToFileChk, 2);

            AddLabel(panel, "PDF output", 7);
            _outputModeCombo = MakeCombo();
            _outputModeCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _outputModeCombo.Items.Add("Combine into one PDF");
            _outputModeCombo.Items.Add("Separate PDFs for selected layouts");
            _outputModeCombo.SelectedIndex = 0;
            _outputModeCombo.SelectedIndexChanged += (s, e) =>
            {
                UpdatePathForOutputMode();
                UpdateCopiesState();
                ValidateUi();
            };
            panel.Controls.Add(_outputModeCombo, 1, 7);
            panel.SetColumnSpan(_outputModeCombo, 2);

            _pathLabel = new Label
            {
                Text = "File path",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            panel.Controls.Add(_pathLabel, 0, 8);
            _outputPathBox = new TextBox { Dock = DockStyle.Fill };
            _outputPathBox.TextChanged += (s, e) => ValidateUi();
            var browseBtn = new Button { Text = "Browse...", Width = 78, Height = 26 };
            browseBtn.Click += (s, e) => BrowseOutputPath();
            panel.Controls.Add(_outputPathBox, 1, 8);
            panel.Controls.Add(browseBtn, 2, 8);

            var spacer = new Panel { Dock = DockStyle.Fill };
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            panel.Controls.Add(spacer, 0, 9);
            panel.SetColumnSpan(spacer, 3);

            _scaleToFitChk = new CheckBox
            {
                Text = "Scale to fit paper (use only when the selected paper differs from the template)",
                Dock = DockStyle.Fill,
                Checked = false
            };
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            panel.Controls.Add(_scaleToFitChk, 0, 10);
            panel.SetColumnSpan(_scaleToFitChk, 3);
            UpdateScaleToFitState();

            _openAfterChk = new CheckBox
            {
                Text = "Open the PDF / output folder when finished",
                Dock = DockStyle.Fill,
                Checked = true
            };
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            panel.Controls.Add(_openAfterChk, 0, 11);
            panel.SetColumnSpan(_openAfterChk, 3);

            return panel;
        }

        private Control BuildButtonsPanel()
        {
            var panel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft,
                Padding = new Padding(0, 9, 0, 0)
            };

            var cancelBtn = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Width = 90, Height = 30 };
            _okBtn = new Button { Text = "Print", Width = 90, Height = 30 };
            _okBtn.Click += (s, e) => TryAccept();
            panel.Controls.Add(cancelBtn);
            panel.Controls.Add(_okBtn);

            AcceptButton = _okBtn;
            CancelButton = cancelBtn;
            return panel;
        }

        private static ComboBox MakeCombo()
        {
            return new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
        }

        private static void AddLabel(TableLayoutPanel panel, string text, int row)
        {
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            panel.Controls.Add(new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            }, 0, row);
        }

        private void LoadValues()
        {
            string lastDevice = null;
            string lastMedia = null;
            string lastStyle = null;
            string lastOutput = null;
            string lastOutputMode = null;
            string lastPlotArea = null;
            int lastOrientation = -1;
            int lastCopies = -1;
            int lastScale = -1;
            int lastToFile = -1;
            int lastOpenAfter = -1;
            try
            {
                using (var k = Registry.CurrentUser.OpenSubKey(RegKey))
                {
                    lastDevice = k?.GetValue("PrintDevice") as string;
                    lastMedia = k?.GetValue("PrintMedia") as string;
                    lastStyle = k?.GetValue("PrintStyle") as string;
                    lastOutput = k?.GetValue("PrintOutputPath") as string;
                    lastOutputMode = k?.GetValue("PrintOutputMode") as string;
                    lastPlotArea = k?.GetValue("PrintPlotArea") as string;
                    lastOrientation = ReadInt(k, "PrintOrientation", -1);
                    lastCopies = ReadInt(k, "PrintCopies", -1);
                    lastScale = ReadInt(k, "PrintScaleToFit", -1);
                    lastToFile = ReadInt(k, "PrintToFile", -1);
                    lastOpenAfter = ReadInt(k, "PrintOpenAfter", -1);
                }
            }
            catch { }

            var devices = LayoutPlotter.GetPlotDevices();
            foreach (var d in devices) _deviceCombo.Items.Add(d);
            SelectComboText(_deviceCombo, lastDevice);
            if (_deviceCombo.SelectedIndex < 0 && _deviceCombo.Items.Count > 0) _deviceCombo.SelectedIndex = 0;

            UpdateMediaList(lastMedia);

            _styleCombo.Items.Add(UseLayoutStyle);
            foreach (var s in LayoutPlotter.GetPlotStyleSheets()) _styleCombo.Items.Add(s);
            SelectComboText(_styleCombo, string.IsNullOrWhiteSpace(lastStyle) ? "monochrome.ctb" : lastStyle);
            if (_styleCombo.SelectedIndex < 0) _styleCombo.SelectedIndex = 0;

            // Default to a single combined PDF; only use per-layout PDFs when the user explicitly chose it before.
            _outputModeCombo.SelectedIndex = string.Equals(lastOutputMode, "SeparatePerLayout", StringComparison.OrdinalIgnoreCase)
                ? 1
                : 0;
            _plotAreaCombo.SelectedIndex = string.Equals(lastPlotArea, "Extents", StringComparison.OrdinalIgnoreCase)
                ? 1
                : 0;
            _outputPathBox.Text = !string.IsNullOrWhiteSpace(lastOutput) ? lastOutput : _defaultPdfPath;
            UpdatePathForOutputMode();

            if (lastOrientation >= 0 && lastOrientation < _orientationCombo.Items.Count)
                _orientationCombo.SelectedIndex = lastOrientation;
            if (lastCopies >= _copiesUpDown.Minimum && lastCopies <= _copiesUpDown.Maximum)
                _copiesUpDown.Value = lastCopies;
            if (lastToFile == 0 || lastToFile == 1)
                _plotToFileChk.Checked = lastToFile == 1;
            if ((lastScale == 0 || lastScale == 1) && _plotAreaCombo.SelectedIndex != 1)
                _scaleToFitChk.Checked = lastScale == 1;
            if (lastOpenAfter == 0 || lastOpenAfter == 1)
                _openAfterChk.Checked = lastOpenAfter == 1;

            UpdateCopiesState();
        }

        private static int ReadInt(RegistryKey k, string name, int fallback)
        {
            try
            {
                var v = k?.GetValue(name);
                if (v != null && int.TryParse(v.ToString(), out var n)) return n;
            }
            catch { }
            return fallback;
        }

        private static string[] ReadMultiString(string name)
        {
            try
            {
                using (var k = Registry.CurrentUser.OpenSubKey(RegKey))
                    return k?.GetValue(name) as string[];
            }
            catch { return null; }
        }

        private static List<PrintableLayout> ApplySavedOrder(List<PrintableLayout> layouts, string[] savedOrder)
        {
            if (savedOrder == null || savedOrder.Length == 0) return layouts;

            var byName = new Dictionary<string, PrintableLayout>(StringComparer.OrdinalIgnoreCase);
            foreach (var l in layouts) byName[l.Name] = l;

            var result = new List<PrintableLayout>();
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var name in savedOrder)
            {
                if (byName.TryGetValue(name, out var l) && used.Add(name))
                    result.Add(l);
            }
            // Append layouts that did not exist when the order was saved (new layouts), keeping their order.
            foreach (var l in layouts)
                if (!used.Contains(l.Name))
                    result.Add(l);
            return result;
        }

        private void UpdateMediaList(string preferredCanonicalName)
        {
            var selected = preferredCanonicalName;
            if (string.IsNullOrWhiteSpace(selected) && _mediaCombo.SelectedItem is PlotMediaOption current)
                selected = current.CanonicalName;

            _mediaCombo.Items.Clear();
            foreach (var media in LayoutPlotter.GetMediaOptions(_deviceCombo.SelectedItem as string))
                _mediaCombo.Items.Add(media);

            if (!string.IsNullOrWhiteSpace(selected))
            {
                for (int i = 0; i < _mediaCombo.Items.Count; i++)
                {
                    if (_mediaCombo.Items[i] is PlotMediaOption m &&
                        string.Equals(m.CanonicalName, selected, StringComparison.OrdinalIgnoreCase))
                    {
                        _mediaCombo.SelectedIndex = i;
                        return;
                    }
                }
            }

            if (_mediaCombo.Items.Count > 0) _mediaCombo.SelectedIndex = 0;
        }

        private static void SelectComboText(ComboBox combo, string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return;
            for (int i = 0; i < combo.Items.Count; i++)
            {
                if (string.Equals(combo.Items[i]?.ToString(), value, StringComparison.OrdinalIgnoreCase))
                {
                    combo.SelectedIndex = i;
                    return;
                }
            }
        }

        private bool SelectedDeviceLooksLikePdf()
        {
            var d = _deviceCombo.SelectedItem as string;
            return !string.IsNullOrWhiteSpace(d)
                && d.IndexOf("PDF", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private void BrowseOutputPath()
        {
            if (IsSeparatePdfMode())
            {
                using (var fbd = new FolderBrowserDialog())
                {
                    fbd.Description = "Select output folder for per-layout PDFs";
                    if (Directory.Exists(_outputPathBox.Text))
                        fbd.SelectedPath = _outputPathBox.Text;
                    else
                    {
                        var dir = Path.GetDirectoryName(_outputPathBox.Text);
                        if (!string.IsNullOrWhiteSpace(dir) && Directory.Exists(dir))
                            fbd.SelectedPath = dir;
                    }

                    if (fbd.ShowDialog(this) == DialogResult.OK)
                        _outputPathBox.Text = fbd.SelectedPath;
                }
            }
            else
            {
                using (var sfd = new System.Windows.Forms.SaveFileDialog())
                {
                    sfd.Title = "Output PDF/file";
                    sfd.Filter = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*";
                    sfd.FileName = Path.GetFileName(_outputPathBox.Text);
                    var dir = Path.GetDirectoryName(_outputPathBox.Text);
                    if (!string.IsNullOrWhiteSpace(dir) && Directory.Exists(dir))
                        sfd.InitialDirectory = dir;
                    sfd.OverwritePrompt = true;

                    if (sfd.ShowDialog(this) == DialogResult.OK)
                        _outputPathBox.Text = sfd.FileName;
                }
            }
        }

        private bool IsSeparatePdfMode()
        {
            return _plotToFileChk.Checked && _outputModeCombo.SelectedIndex == 1;
        }

        private void UpdatePathForOutputMode()
        {
            if (_outputPathBox == null) return;
            if (IsSeparatePdfMode())
            {
                var current = _outputPathBox.Text;
                _outputPathBox.Text = GetSeparateOutputFolder(current);
            }
            else
            {
                var current = _outputPathBox.Text;
                if (!string.IsNullOrWhiteSpace(current) && Directory.Exists(current))
                {
                    var fileName = Path.GetFileName(_defaultPdfPath);
                    _outputPathBox.Text = Path.Combine(current, fileName);
                }
            }
            UpdatePathLabel();
        }

        private void UpdateScaleToFitState()
        {
            if (_scaleToFitChk == null || _plotAreaCombo == null) return;
            if (_plotAreaCombo.SelectedIndex == 1)
            {
                _scaleToFitChk.Checked = true;
                _scaleToFitChk.Enabled = false;
            }
            else
            {
                _scaleToFitChk.Enabled = true;
            }
        }

        private void UpdateCopiesState()
        {
            if (_copiesUpDown == null || _plotToFileChk == null) return;
            // Copies apply only to a physical printer. PDF/file output always produces one file.
            _copiesUpDown.Enabled = !_plotToFileChk.Checked;
        }

        private void UpdatePathLabel()
        {
            if (_pathLabel == null) return;
            _pathLabel.Text = IsSeparatePdfMode() ? "Folder" : "File path";
        }

        private void SetAllLayouts(bool value)
        {
            for (int i = 0; i < _layoutList.Items.Count; i++)
                _layoutList.SetItemChecked(i, value);
            ValidateUi();
        }

        private void MoveSelectedLayout(int delta)
        {
            int index = _layoutList.SelectedIndex;
            if (index < 0) return;

            int target = index + delta;
            if (target < 0 || target >= _layoutList.Items.Count) return;

            var item = _layoutList.Items[index];
            bool isChecked = _layoutList.GetItemChecked(index);

            _layoutList.Items.RemoveAt(index);
            _layoutList.Items.Insert(target, item);
            _layoutList.SetItemChecked(target, isChecked);
            _layoutList.SelectedIndex = target;
            ValidateUi();
        }

        private void ValidateUi()
        {
            if (_okBtn == null) return;
            bool ok = _layoutList.CheckedItems.Count > 0
                && _deviceCombo.SelectedItem != null
                && _mediaCombo.SelectedItem is PlotMediaOption
                && (!_plotToFileChk.Checked || !string.IsNullOrWhiteSpace(_outputPathBox.Text));
            _okBtn.Enabled = ok;
        }

        private void TryAccept()
        {
            var selected = _layoutList.CheckedItems.Cast<PrintableLayout>().ToList();
            if (selected.Count == 0)
            {
                MessageBox.Show(this, "Select at least one layout.", "Print",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                DialogResult = DialogResult.None;
                return;
            }
            if (!(_mediaCombo.SelectedItem is PlotMediaOption media))
            {
                MessageBox.Show(this, "Select a paper size.", "Print",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                DialogResult = DialogResult.None;
                return;
            }
            string style = _styleCombo.SelectedItem?.ToString();
            if (style == UseLayoutStyle) style = null;

            var outputMode = ParseOutputMode();
            var outputPath = _outputPathBox.Text.Trim();
            if (_plotToFileChk.Checked && outputMode == PdfOutputMode.SeparatePerLayout)
            {
                outputPath = GetSeparateOutputFolder(outputPath);
                _outputPathBox.Text = outputPath;
            }

            Options = new PrintJobOptions
            {
                DeviceName = _deviceCombo.SelectedItem as string,
                MediaName = media.CanonicalName,
                StyleSheet = style,
                Copies = (int)_copiesUpDown.Value,
                PlotToFile = _plotToFileChk.Checked,
                OutputPath = outputPath,
                PdfOutputMode = outputMode,
                PlotArea = ParsePlotArea(),
                ScaleToFit = _scaleToFitChk.Checked,
                Orientation = ParseOrientation(),
                OpenAfterExport = _openAfterChk.Checked,
                Layouts = selected
            };

            SaveValues(media.CanonicalName, style, outputPath, outputMode, ParsePlotArea());
            DialogResult = DialogResult.OK;
            Close();
        }

        private string GetSeparateOutputFolder(string path)
        {
            var drawingFolder = DefaultSeparateOutputFolder();
            if (string.IsNullOrWhiteSpace(path)) return drawingFolder;
            if (Path.HasExtension(path)) return drawingFolder;

            var drawingName = Path.GetFileName(drawingFolder);
            var selectedName = Path.GetFileName(path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
            if (string.Equals(selectedName, drawingName, StringComparison.OrdinalIgnoreCase))
                return path;

            return Path.Combine(path, drawingName);
        }

        private string DefaultSeparateOutputFolder()
        {
            var dir = Path.GetDirectoryName(_defaultPdfPath);
            var name = Path.GetFileNameWithoutExtension(_defaultPdfPath);
            if (string.IsNullOrWhiteSpace(name)) return dir;
            if (name.EndsWith("_layout", StringComparison.OrdinalIgnoreCase))
                name = name.Substring(0, name.Length - "_layout".Length);
            return Path.Combine(dir, name);
        }

        private PlotOrientationChoice ParseOrientation()
        {
            switch (_orientationCombo.SelectedIndex)
            {
                case 1: return PlotOrientationChoice.Portrait;
                case 2: return PlotOrientationChoice.Landscape;
                default: return PlotOrientationChoice.LayoutSetting;
            }
        }

        private PdfOutputMode ParseOutputMode()
        {
            return _outputModeCombo.SelectedIndex == 1
                ? PdfOutputMode.SeparatePerLayout
                : PdfOutputMode.Combined;
        }

        private PlotAreaChoice ParsePlotArea()
        {
            return _plotAreaCombo.SelectedIndex == 1
                ? PlotAreaChoice.Extents
                : PlotAreaChoice.Layout;
        }

        private void SaveValues(
            string mediaName,
            string style,
            string outputPath,
            PdfOutputMode outputMode,
            PlotAreaChoice plotArea)
        {
            try
            {
                using (var k = Registry.CurrentUser.CreateSubKey(RegKey))
                {
                    k?.SetValue("PrintDevice", _deviceCombo.SelectedItem as string ?? "");
                    k?.SetValue("PrintMedia", mediaName ?? "");
                    k?.SetValue("PrintStyle", style ?? "");
                    k?.SetValue("PrintOutputPath", outputPath ?? "");
                    k?.SetValue("PrintOutputMode", outputMode.ToString());
                    k?.SetValue("PrintPlotArea", plotArea.ToString());
                    k?.SetValue("PrintOrientation", _orientationCombo.SelectedIndex);
                    k?.SetValue("PrintCopies", (int)_copiesUpDown.Value);
                    k?.SetValue("PrintScaleToFit", _scaleToFitChk.Checked ? 1 : 0);
                    k?.SetValue("PrintToFile", _plotToFileChk.Checked ? 1 : 0);
                    k?.SetValue("PrintOpenAfter", _openAfterChk.Checked ? 1 : 0);

                    var order = _layoutList.Items.Cast<PrintableLayout>().Select(l => l.Name).ToArray();
                    var checkedNames = _layoutList.CheckedItems.Cast<PrintableLayout>().Select(l => l.Name).ToArray();
                    k?.SetValue("PrintLayoutOrder", order, RegistryValueKind.MultiString);
                    k?.SetValue("PrintLayoutChecked", checkedNames, RegistryValueKind.MultiString);
                }
            }
            catch { }
        }
    }
}
