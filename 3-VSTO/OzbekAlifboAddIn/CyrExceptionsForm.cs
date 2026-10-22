using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace OzbekAlifboAddIn
{
    /// <summary>
    /// Kirilldan o'girish istisnolari: "so'z" — o'girilmaydi, "o'zak=Lotin" — maxsus shakl
    /// (Майкрософтнинг → Microsoftning). Fayl: %APPDATA%\OzbekAlifbo\kirill-istisnolar.txt
    /// </summary>
    internal sealed class CyrExceptionsForm : Form
    {
        readonly Word.Application _app;
        readonly List<string> _users = new List<string>();
        readonly List<string> _off = new List<string>();
        readonly TextBox _txtKey = new TextBox();
        readonly TextBox _txtLat = new TextBox();
        readonly ListView _lstUser = new ListView();
        readonly CheckedListBox _lstDef = new CheckedListBox();
        readonly Label _lblMsg = new Label();
        readonly TextBox _txtTry = new TextBox();
        readonly Label _lblTry = new Label();
        bool _loading, _dirty;

        public CyrExceptionsForm(Word.Application app)
        {
            _app = app;
            Text = "Özbekça — kirill istisnolari";
            Font = new Font("Segoe UI", 9f);
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(600, 540);
            MinimumSize = new Size(520, 460);
            BuildUi();
            ExceptionStore.Split(ExceptionStore.ReadLines(Converter.CyrExceptionsPath), _users, _off);
            Render();
        }

        void BuildUi()
        {
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(12), ColumnCount = 2, RowCount = 6 };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));
            for (int i = 0; i < 3; i++) root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(root);

            var help = new Label
            {
                AutoSize = true, MaximumSize = new Size(570, 0), ForeColor = SystemColors.GrayText,
                Text = "Faqat kirill so'z — o'girilmaydi (masalan, ruscha ism).  Lotin shakli bilan — shu shaklda yoziladi,\n" +
                       "qo'shimchalari odatdagidek o'giriladi: Майкрософтнинг → Microsoftning.  So'z oxirida * — shu bilan boshlanganlar."
            };
            root.Controls.Add(help, 0, 0); root.SetColumnSpan(help, 2);

            var addRow = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true, Margin = new Padding(0, 6, 0, 0) };
            _txtKey.Width = 150; _txtLat.Width = 150;
            KeyEventHandler enter = (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; Add(); } };
            _txtKey.KeyDown += enter; _txtLat.KeyDown += enter;
            var btnAdd = new Button { Text = "Qo'shish", AutoSize = true };
            btnAdd.Click += (s, e) => Add();
            var btnSel = new Button { Text = "Belgilangan so'zni olish", AutoSize = true };
            btnSel.Click += (s, e) => FromSelection();
            addRow.Controls.AddRange(new Control[] {
                new Label { Text = "Kirill:", AutoSize = true, Margin = new Padding(0, 6, 2, 0) }, _txtKey,
                new Label { Text = "→ Lotin:", AutoSize = true, Margin = new Padding(6, 6, 2, 0) }, _txtLat, btnAdd, btnSel });
            root.Controls.Add(addRow, 0, 1); root.SetColumnSpan(addRow, 2);

            _lblMsg.AutoSize = true; _lblMsg.ForeColor = Color.FromArgb(30, 110, 90); _lblMsg.Margin = new Padding(3, 4, 3, 6);
            root.Controls.Add(_lblMsg, 0, 2); root.SetColumnSpan(_lblMsg, 2);

            // mening yozuvlarim
            var left = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, ColumnCount = 1, Margin = new Padding(0, 0, 6, 0) };
            left.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            left.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            left.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            left.Controls.Add(new Label { Text = "Mening yozuvlarim", AutoSize = true, Font = new Font(Font, FontStyle.Bold) }, 0, 0);
            _lstUser.Dock = DockStyle.Fill; _lstUser.View = View.Details; _lstUser.FullRowSelect = true;
            _lstUser.HideSelection = false; _lstUser.MultiSelect = true;
            _lstUser.Columns.Add("Kirill", 140); _lstUser.Columns.Add("Lotin", 150);
            _lstUser.KeyDown += (s, e) => { if (e.KeyCode == Keys.Delete) RemoveSelected(); };
            _lstUser.DoubleClick += (s, e) =>
            {
                if (_lstUser.SelectedItems.Count == 0) return;
                var it = _lstUser.SelectedItems[0];
                _txtKey.Text = it.Text; _txtLat.Text = it.SubItems[1].Text == "(o'girilmaydi)" ? "" : it.SubItems[1].Text;
            };
            left.Controls.Add(_lstUser, 0, 1);
            var btnDel = new Button { Text = "Tanlanganni o'chirish (Del)", AutoSize = true, Dock = DockStyle.Fill };
            btnDel.Click += (s, e) => RemoveSelected();
            left.Controls.Add(btnDel, 0, 2);
            root.Controls.Add(left, 0, 3);

            // standartlar
            var right = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, Margin = new Padding(0) };
            right.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            right.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            right.Controls.Add(new Label { Text = "Standart yozuvlar", AutoSize = true, Font = new Font(Font, FontStyle.Bold) }, 0, 0);
            _lstDef.Dock = DockStyle.Fill; _lstDef.CheckOnClick = true;
            _lstDef.ItemCheck += OnDefCheck;
            right.Controls.Add(_lstDef, 0, 1);
            root.Controls.Add(right, 1, 3);

            // sinab ko'rish
            var tryRow = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, Margin = new Padding(0, 8, 0, 0) };
            _txtTry.Width = 220;
            _txtTry.TextChanged += (s, e) => Try();
            _lblTry.AutoSize = true; _lblTry.Margin = new Padding(6, 6, 0, 0); _lblTry.ForeColor = Color.FromArgb(30, 110, 90);
            tryRow.Controls.AddRange(new Control[] { new Label { Text = "Sinab ko'rish:", AutoSize = true, Margin = new Padding(0, 6, 4, 0) }, _txtTry, _lblTry });
            root.Controls.Add(tryRow, 0, 4); root.SetColumnSpan(tryRow, 2);

            var bottom = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.RightToLeft, Margin = new Padding(0, 8, 0, 0) };
            var btnCancel = new Button { Text = "Bekor qilish", AutoSize = true, DialogResult = DialogResult.Cancel };
            var btnSave = new Button { Text = "Saqlash", AutoSize = true, DialogResult = DialogResult.OK };
            btnSave.Click += (s, e) => { ExceptionStore.Save(_users, _off, Converter.CyrExceptionsPath); _dirty = false; };
            bottom.Controls.AddRange(new Control[] { btnCancel, btnSave });
            root.Controls.Add(bottom, 0, 5); root.SetColumnSpan(bottom, 2);
            CancelButton = btnCancel;
            FormClosing += (s, e) =>
            {
                if (DialogResult == DialogResult.Cancel && _dirty &&
                    MessageBox.Show(this, "O'zgarishlar saqlanmagan. Baribir yopilsinmi?", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    e.Cancel = true;
            };
        }

        static string EntryText(string e)
        {
            int i = e.IndexOf('=');
            return i > 0 ? e.Substring(0, i) + "  →  " + e.Substring(i + 1) : e + "  (o'girilmaydi)";
        }

        bool IsOff(string d) => _off.Any(o => Converter.ExceptionKey(o) == Converter.ExceptionKey(d));

        void Render()
        {
            _loading = true;
            _lstUser.BeginUpdate(); _lstUser.Items.Clear();
            foreach (var u in _users)
            {
                int i = u.IndexOf('=');
                var it = new ListViewItem(i > 0 ? u.Substring(0, i) : u);
                it.SubItems.Add(i > 0 ? u.Substring(i + 1) : "(o'girilmaydi)");
                it.Tag = u;
                _lstUser.Items.Add(it);
            }
            _lstUser.EndUpdate();
            _lstDef.BeginUpdate(); _lstDef.Items.Clear();
            foreach (var d in Converter.DefaultCyrExceptions) _lstDef.Items.Add(EntryText(d), !IsOff(d));
            _lstDef.EndUpdate();
            _loading = false;
            Try();
        }

        void OnDefCheck(object sender, ItemCheckEventArgs e)
        {
            if (_loading) return;
            var d = Converter.DefaultCyrExceptions[e.Index];
            _off.RemoveAll(o => Converter.ExceptionKey(o) == Converter.ExceptionKey(d));
            if (e.NewValue == CheckState.Unchecked) _off.Add(d.Split('=')[0]);
            _dirty = true;
            BeginInvoke((Action)Try);
        }

        void Add()
        {
            var k = _txtKey.Text.Trim(); var l = _txtLat.Text.Trim();
            if (k.Length == 0) { _lblMsg.Text = "Kirill so'zni kiriting."; return; }
            if (k.Any(char.IsWhiteSpace) || !k.Any(c => c >= 'Ѐ' && c <= 'ӿ')) { _lblMsg.Text = "Bitta kirill so'z (bo'shliqsiz) kiriting."; return; }
            if (l.Any(c => c >= 'Ѐ' && c <= 'ӿ')) { _lblMsg.Text = "Lotin shaklida kirill harf bo'lmasin."; return; }
            var entry = l.Length > 0 ? k + "=" + l : k;
            _users.RemoveAll(u => Converter.ExceptionKey(u) == Converter.ExceptionKey(entry));
            _users.Add(entry);
            _dirty = true;
            _txtKey.Clear(); _txtLat.Clear(); _txtKey.Focus();
            _lblMsg.Text = "Qo'shildi: " + EntryText(entry);
            Render();
        }

        void RemoveSelected()
        {
            var sel = _lstUser.SelectedItems.Cast<ListViewItem>().Select(i => (string)i.Tag).ToList();
            if (sel.Count == 0) { _lblMsg.Text = "O'chirish uchun ro'yxatdan yozuv tanlang."; return; }
            foreach (var u in sel) _users.Remove(u);
            _dirty = true;
            _lblMsg.Text = sel.Count + " ta yozuv o'chirildi.";
            Render();
        }

        internal void PrefillFromSelection() => FromSelection();

        void FromSelection()
        {
            try
            {
                if (_app.Documents.Count == 0) return;
                var t = (_app.Selection.Text ?? "").Trim();
                t = new string(t.SkipWhile(c => !(c >= 'Ѐ' && c <= 'ӿ')).Reverse()
                                .SkipWhile(c => !(c >= 'Ѐ' && c <= 'ӿ')).Reverse().ToArray());
                if (t.Length == 0 || t.Any(char.IsWhiteSpace)) { _lblMsg.Text = "Hujjatda bitta kirill so'zni belgilang."; return; }
                _txtKey.Text = t; _txtLat.Focus();
                _lblMsg.Text = "Lotin shaklini yozing (yoki bo'sh qoldiring) va «Qo'shish» ni bosing.";
            }
            catch (Exception ex) { _lblMsg.Text = ex.Message; }
        }

        void Try()
        {
            var t = _txtTry.Text;
            if (t.Length == 0) { _lblTry.Text = ""; return; }
            var lines = _off.Select(o => "!" + o).Concat(_users);
            _lblTry.Text = "→  " + Converter.CyrToLatin(t, false, Converter.MergeCyrExceptions(lines));
        }
    }
}
