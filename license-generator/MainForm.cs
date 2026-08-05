using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PrintLayoutAddin.LicenseGenerator
{
    // GUI with two modes (tabs):
    //   Mode 1 "Cap 1 key"      -> type user / machine id / expiry, get one key.
    //   Mode 2 "Cap nhieu key"  -> export a CSV template, fill it, import it, get many keys.
    public class MainForm : Form
    {
        private static readonly Color AccentBlue = Color.FromArgb(0, 120, 215);
        private static readonly Color OkBg = Color.FromArgb(232, 245, 233);
        private static readonly Color ErrBg = Color.FromArgb(253, 232, 232);
        private static readonly Color WarnBg = Color.FromArgb(255, 248, 225);
        private static readonly Color OkFg = Color.FromArgb(30, 138, 60);
        private static readonly Color ErrFg = Color.FromArgb(196, 60, 60);
        private static readonly Color HintGray = Color.FromArgb(96, 102, 110);

        // Mode 1 controls
        private TextBox _m1Note, _m1Mid, _m1Key;
        private DateTimePicker _m1Exp;
        private Button _m1Copy;

        // Mode 2 controls
        private DataGridView _grid;
        private DataGridViewColumn _colNote, _colMid, _colExp, _colKey, _colStatus;

        private Label _status;

        public MainForm()
        {
            BuildUI();
        }

        private void BuildUI()
        {
            Text = "PrintLayoutAddin — Cấp License";
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Segoe UI", 9F);
            BackColor = Color.White;
            ClientSize = new Size(900, 560);
            MinimumSize = new Size(680, 440);
            AutoScaleMode = AutoScaleMode.Font;

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(10),
                BackColor = Color.White,
            };
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var tabs = new TabControl { Dock = DockStyle.Fill, Padding = new Point(14, 6) };
            tabs.TabPages.Add(BuildModeOneTab());
            tabs.TabPages.Add(BuildModeTwoTab());
            root.Controls.Add(tabs, 0, 0);

            _status = new Label
            {
                Dock = DockStyle.Fill,
                Height = 24,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = HintGray,
                Margin = new Padding(2, 6, 2, 0),
                Text = "Mode 1: gõ tay 1 máy.   Mode 2: xuất file mẫu, điền rồi nhập CSV để cấp nhiều key.",
            };
            root.Controls.Add(_status, 0, 1);

            Controls.Add(root);
        }

        // ---------------- Mode 1: single key ----------------

        private TabPage BuildModeOneTab()
        {
            var tab = new TabPage("  Cấp 1 key  ") { BackColor = Color.White, Padding = new Padding(16) };

            var t = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                AutoSize = false,
            };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            int row = 0;
            void AddPair(string label, Control c)
            {
                t.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                t.Controls.Add(new Label
                {
                    Text = label,
                    AutoSize = true,
                    Anchor = AnchorStyles.Left,
                    Margin = new Padding(0, 8, 8, 8),
                    Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold),
                }, 0, row);
                t.Controls.Add(c, 1, row);
                row++;
            }

            _m1Note = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(0, 5, 0, 5) };
            _m1Mid = new TextBox
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 5, 0, 5),
                Font = new Font("Consolas", 10F),
            };
            _m1Exp = new DateTimePicker
            {
                Format = DateTimePickerFormat.Custom,
                CustomFormat = "yyyy-MM-dd",
                Width = 140,
                Anchor = AnchorStyles.Left,
                Value = DateTime.Today.AddYears(1),
                Margin = new Padding(0, 5, 0, 5),
            };

            AddPair("Tên user / Ghi chú:", _m1Note);
            AddPair("Machine ID:", _m1Mid);
            AddPair("Ngày hết hạn:", _m1Exp);

            var btnGen = MakeButton("Tạo key", OnGenerateOne, primary: true);
            btnGen.Margin = new Padding(0, 6, 0, 10);
            t.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            t.Controls.Add(btnGen, 1, row++);

            t.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            t.Controls.Add(new Label
            {
                Text = "License Key:",
                AutoSize = true,
                Margin = new Padding(0, 4, 0, 2),
                Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold),
            }, 0, row);
            _m1Key = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                Height = 70,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Consolas", 9F),
                BackColor = Color.FromArgb(250, 251, 252),
                Margin = new Padding(0, 4, 0, 6),
            };
            t.Controls.Add(_m1Key, 1, row++);

            _m1Copy = MakeButton("Copy key", OnCopyOne, primary: false);
            _m1Copy.Enabled = false;
            t.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            t.Controls.Add(_m1Copy, 1, row++);

            // filler row so controls stay at the top
            t.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            tab.Controls.Add(t);
            return tab;
        }

        private void OnGenerateOne(object sender, EventArgs e)
        {
            string mid = _m1Mid.Text.Trim();
            string note = _m1Note.Text.Trim();
            if (string.IsNullOrWhiteSpace(mid))
            {
                Error("Hãy nhập Machine ID (khách lấy từ dialog PLLICENSE).");
                _m1Mid.Focus();
                return;
            }
            DateTime exp = _m1Exp.Value.Date;
            string payload;
            string key = LicenseCore.MakeKey(mid, exp, note, out payload);
            LicenseCore.LogIssued(mid, exp, note, key);

            _m1Key.Text = key;
            _m1Copy.Enabled = true;
            TrySetClipboard(key);
            SetStatus("Đã tạo key cho " + mid + " (hết hạn " + exp.ToString("yyyy-MM-dd") +
                      ") — đã copy sẵn vào clipboard.", OkFg);
        }

        private void OnCopyOne(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_m1Key.Text)) return;
            if (TrySetClipboard(_m1Key.Text))
                SetStatus("Đã copy key. Gửi cho khách dán vào dialog PLLICENSE.", OkFg);
        }

        // ---------------- Mode 2: batch via CSV ----------------

        private TabPage BuildModeTwoTab()
        {
            var tab = new TabPage("  Cấp nhiều key (CSV)  ") { BackColor = Color.White, Padding = new Padding(12) };

            var t = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
            };
            t.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // step 1 buttons
            t.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // hint
            t.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // grid
            t.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // step 2 buttons

            var top = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true };
            top.Controls.Add(MakeButton("1. Xuất file mẫu (template)…", OnExportTemplate, primary: false));
            top.Controls.Add(MakeButton("2. Nhập CSV…", OnImportCsv, primary: false));
            t.Controls.Add(top, 0, 0);

            t.Controls.Add(new Label
            {
                AutoSize = true,
                ForeColor = HintGray,
                Margin = new Padding(2, 4, 2, 8),
                Font = new Font("Segoe UI", 8.5F, FontStyle.Italic),
                Text = "Xuất file mẫu → mở bằng Excel/Sheet, điền machine_id + expire (+note) → lưu lại → Nhập CSV.",
            }, 0, 1);

            t.Controls.Add(BuildGrid(), 0, 2);

            var bottom = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true, Margin = new Padding(0, 8, 0, 0) };
            bottom.Controls.Add(MakeButton("3. Tạo tất cả key", OnGenerateAll, primary: true));
            bottom.Controls.Add(MakeButton("Copy cột key", OnCopyKeys, primary: false));
            bottom.Controls.Add(MakeButton("Lưu CSV kết quả…", OnSaveCsv, primary: false));
            bottom.Controls.Add(MakeButton("Xóa bảng", OnClearAll, primary: false));
            t.Controls.Add(bottom, 0, 3);

            tab.Controls.Add(t);
            return tab;
        }

        private Control BuildGrid()
        {
            _grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = true,
                AllowUserToResizeRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                RowHeadersWidth = 30,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2,
            };
            _grid.EnableHeadersVisualStyles = false;
            _grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold);
            _grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(245, 247, 250);

            _colNote = new DataGridViewTextBoxColumn { HeaderText = "Người dùng / Ghi chú", Width = 200 };
            _colMid = new DataGridViewTextBoxColumn
            {
                HeaderText = "Machine ID",
                Width = 185,
                DefaultCellStyle = new DataGridViewCellStyle { Font = new Font("Consolas", 9.5F) },
            };
            _colExp = new DataGridViewTextBoxColumn { HeaderText = "Hết hạn (YYYY-MM-DD)", Width = 150 };
            _colKey = new DataGridViewTextBoxColumn
            {
                HeaderText = "License Key",
                Width = 220,
                ReadOnly = true,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    Font = new Font("Consolas", 8.5F),
                    BackColor = Color.FromArgb(250, 251, 252),
                },
            };
            _colStatus = new DataGridViewTextBoxColumn { HeaderText = "Trạng thái", Width = 110, ReadOnly = true };
            _grid.Columns.AddRange(_colNote, _colMid, _colExp, _colKey, _colStatus);

            _grid.CellDoubleClick += (s, e) =>
            {
                if (e.RowIndex >= 0 && e.ColumnIndex == _colKey.Index)
                {
                    var v = _grid.Rows[e.RowIndex].Cells[_colKey.Index].Value as string;
                    if (!string.IsNullOrEmpty(v)) { TrySetClipboard(v); SetStatus("Đã copy 1 key của dòng đang chọn.", HintGray); }
                }
            };
            return _grid;
        }

        private void OnExportTemplate(object sender, EventArgs e)
        {
            using (var dlg = new SaveFileDialog
            {
                Title = "Lưu file mẫu",
                Filter = "CSV (*.csv)|*.csv",
                FileName = "license_template.csv",
            })
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                var sb = new StringBuilder();
                sb.Append("user,machine_id,expire,note\r\n");
                sb.Append("Nguyen Van A,ABCD-EFGH-JKLM-NPQR,2027-12-31,Cong ty A (xoa dong vi du nay)\r\n");
                try
                {
                    File.WriteAllText(dlg.FileName, sb.ToString(), new UTF8Encoding(true));
                    SetStatus("Đã xuất file mẫu: " + dlg.FileName + ". Điền dữ liệu rồi bấm \"Nhập CSV\".", OkFg);
                }
                catch (Exception ex) { Error("Không lưu được file mẫu: " + ex.Message); }
            }
        }

        private void OnImportCsv(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog
            {
                Title = "Chọn file CSV đã điền",
                Filter = "CSV / text (*.csv;*.txt)|*.csv;*.txt|Tất cả (*.*)|*.*",
            })
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                List<string[]> rows;
                try { rows = LicenseCore.ParseTable(File.ReadAllText(dlg.FileName, Encoding.UTF8)); }
                catch (Exception ex) { Error("Không đọc được file: " + ex.Message); return; }

                if (rows.Count < 2) { Error("File cần có dòng tiêu đề + ít nhất 1 dòng dữ liệu."); return; }

                string[] header = rows[0];
                int ciMid = LicenseCore.FindColumn(header, LicenseCore.MidAliases);
                int ciExp = LicenseCore.FindColumn(header, LicenseCore.ExpAliases);
                int ciUser = LicenseCore.FindColumn(header, LicenseCore.UserAliases);
                int ciNote = LicenseCore.FindColumn(header, LicenseCore.NoteAliases);
                if (ciMid < 0 || ciExp < 0)
                {
                    Error("Không nhận diện được cột machine_id và/hoặc expire ở dòng tiêu đề.\r\n" +
                          "Tiêu đề đọc được: " + string.Join(" | ", header) + "\r\n" +
                          "Hãy dùng đúng file mẫu (cột: user, machine_id, expire, note).");
                    return;
                }

                _grid.Rows.Clear();
                int added = 0;
                for (int r = 1; r < rows.Count; r++)
                {
                    string[] row = rows[r];
                    string mid = Cell(row, ciMid);
                    string exp = Cell(row, ciExp);
                    string note = LicenseCore.CombineNote(
                        ciUser >= 0 ? Cell(row, ciUser) : "",
                        ciNote >= 0 ? Cell(row, ciNote) : "");
                    if (string.IsNullOrWhiteSpace(mid) && string.IsNullOrWhiteSpace(exp) && string.IsNullOrWhiteSpace(note))
                        continue;
                    DateTime d;
                    if (LicenseCore.TryParseDate(exp, out d)) exp = d.ToString("yyyy-MM-dd");
                    _grid.Rows.Add(note, mid, exp, "", "");
                    added++;
                }
                SetStatus("Đã nhập " + added + " dòng từ " + Path.GetFileName(dlg.FileName) +
                          ". Bấm \"Tạo tất cả key\".", HintGray);
            }
        }

        private void OnGenerateAll(object sender, EventArgs e)
        {
            if (_grid.IsCurrentCellInEditMode) _grid.EndEdit();

            int ok = 0, fail = 0;
            foreach (DataGridViewRow row in _grid.Rows)
            {
                if (row.IsNewRow) continue;
                string mid = ((row.Cells[_colMid.Index].Value as string) ?? "").Trim();
                string expStr = ((row.Cells[_colExp.Index].Value as string) ?? "").Trim();
                string note = ((row.Cells[_colNote.Index].Value as string) ?? "").Trim();

                if (string.IsNullOrWhiteSpace(mid) && string.IsNullOrWhiteSpace(expStr) && string.IsNullOrWhiteSpace(note))
                    continue;

                if (string.IsNullOrWhiteSpace(mid)) { SetRowResult(row, "", "Thiếu Machine ID", ErrBg); fail++; continue; }
                DateTime exp;
                if (!LicenseCore.TryParseDate(expStr, out exp)) { SetRowResult(row, "", "Ngày sai định dạng", ErrBg); fail++; continue; }
                row.Cells[_colExp.Index].Value = exp.ToString("yyyy-MM-dd");

                string payload;
                string key = LicenseCore.MakeKey(mid, exp, note, out payload);
                LicenseCore.LogIssued(mid, exp, note, key);
                bool expired = exp.Date < DateTime.UtcNow.Date;
                SetRowResult(row, key, expired ? "OK (đã hết hạn)" : "OK", expired ? WarnBg : OkBg);
                ok++;
            }

            if (ok == 0 && fail == 0)
                SetStatus("Chưa có dòng nào. Nhập CSV hoặc gõ trực tiếp vào bảng.", HintGray);
            else
                SetStatus("Đã tạo " + ok + " key, lỗi " + fail + ". " +
                          (fail > 0 ? "Sửa dòng đỏ rồi tạo lại. " : "") +
                          "Dùng \"Copy cột key\" hoặc \"Lưu CSV kết quả\".",
                          fail > 0 ? ErrFg : OkFg);
        }

        private void OnCopyKeys(object sender, EventArgs e)
        {
            var keys = new List<string>();
            foreach (DataGridViewRow row in _grid.Rows)
            {
                if (row.IsNewRow) continue;
                string k = (row.Cells[_colKey.Index].Value as string) ?? "";
                if (!string.IsNullOrWhiteSpace(k)) keys.Add(k);
            }
            if (keys.Count == 0) { Error("Chưa có key nào. Bấm \"Tạo tất cả key\" trước."); return; }
            if (TrySetClipboard(string.Join("\r\n", keys)))
                SetStatus("Đã copy " + keys.Count + " key (mỗi dòng 1 key). Dán vào cột key trong Google Sheet.", OkFg);
        }

        private void OnSaveCsv(object sender, EventArgs e)
        {
            if (_grid.IsCurrentCellInEditMode) _grid.EndEdit();
            using (var dlg = new SaveFileDialog
            {
                Title = "Lưu kết quả CSV",
                Filter = "CSV (*.csv)|*.csv",
                FileName = "license_keys_" + DateTime.Now.ToString("yyyyMMdd") + ".csv",
            })
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                var sb = new StringBuilder();
                sb.Append("note,machine_id,expire,license_key,status\r\n");
                int n = 0;
                foreach (DataGridViewRow row in _grid.Rows)
                {
                    if (row.IsNewRow) continue;
                    string note = (row.Cells[_colNote.Index].Value as string) ?? "";
                    string mid = (row.Cells[_colMid.Index].Value as string) ?? "";
                    string exp = (row.Cells[_colExp.Index].Value as string) ?? "";
                    string key = (row.Cells[_colKey.Index].Value as string) ?? "";
                    string st = (row.Cells[_colStatus.Index].Value as string) ?? "";
                    if (string.IsNullOrWhiteSpace(note) && string.IsNullOrWhiteSpace(mid) &&
                        string.IsNullOrWhiteSpace(exp) && string.IsNullOrWhiteSpace(key))
                        continue;
                    sb.Append(string.Join(",", new[] { note, mid, exp, key, st }.Select(LicenseCore.CsvEscape)));
                    sb.Append("\r\n");
                    n++;
                }
                try
                {
                    File.WriteAllText(dlg.FileName, sb.ToString(), new UTF8Encoding(true));
                    SetStatus("Đã lưu " + n + " dòng vào " + dlg.FileName, OkFg);
                }
                catch (Exception ex) { Error("Không lưu được: " + ex.Message); }
            }
        }

        private void OnClearAll(object sender, EventArgs e)
        {
            if (_grid.Rows.Count == 0) return;
            if (MessageBox.Show(this, "Xóa toàn bộ dòng trong bảng?", "Xác nhận",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                _grid.Rows.Clear();
                SetStatus("Đã xóa bảng.", HintGray);
            }
        }

        // ---------------- shared helpers ----------------

        private Button MakeButton(string text, EventHandler onClick, bool primary)
        {
            var b = new Button
            {
                Text = text,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                MinimumSize = new Size(96, 32),
                Padding = new Padding(10, 4, 10, 4),
                Margin = new Padding(0, 0, 8, 0),
                FlatStyle = primary ? FlatStyle.Flat : FlatStyle.System,
                Cursor = Cursors.Hand,
            };
            if (primary)
            {
                b.BackColor = AccentBlue;
                b.ForeColor = Color.White;
                b.Font = new Font("Segoe UI Semibold", 9.5F, FontStyle.Bold);
                b.FlatAppearance.BorderSize = 0;
            }
            b.Click += onClick;
            return b;
        }

        private void SetRowResult(DataGridViewRow row, string key, string status, Color bg)
        {
            row.Cells[_colKey.Index].Value = key;
            row.Cells[_colStatus.Index].Value = status;
            row.Cells[_colStatus.Index].Style.BackColor = bg;
            row.Cells[_colKey.Index].Style.BackColor = bg;
        }

        private static string Cell(string[] row, int idx)
        {
            return (idx >= 0 && idx < row.Length) ? (row[idx] ?? "").Trim() : "";
        }

        private bool TrySetClipboard(string text)
        {
            try { Clipboard.SetText(text); return true; }
            catch (Exception ex) { Error("Không copy được: " + ex.Message); return false; }
        }

        private void SetStatus(string text, Color color)
        {
            _status.ForeColor = color;
            _status.Text = text;
        }

        private void Error(string text)
        {
            MessageBox.Show(this, text, "Cấp License", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}
