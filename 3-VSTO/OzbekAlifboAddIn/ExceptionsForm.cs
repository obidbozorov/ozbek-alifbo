using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace OzbekAlifboAddIn
{
    /// <summary>
    /// Istisno so'zlarni ko'rish va boshqarish oynasi.
    /// O'zgarishlar "Saqlash" bosilganda %APPDATA%\OzbekAlifbo\istisnolar.txt ga yoziladi.
    /// </summary>
    internal sealed class ExceptionsForm : Form
    {
        readonly Word.Application _app;
        readonly List<string> _users = new List<string>();
        readonly List<string> _off = new List<string>();
        bool _dirty, _loading;

        readonly TextBox _txtNew = new TextBox();
        readonly TextBox _txtSearch = new TextBox();
        readonly TextBox _txtTest = new TextBox();
        readonly ListBox _lstUser = new ListBox();
        readonly CheckedListBox _lstDef = new CheckedListBox();
        readonly Label _lblTest = new Label();
        readonly Label _lblMsg = new Label();
        readonly Label _lblUser = new Label();
        readonly Label _lblDef = new Label();

        public ExceptionsForm(Word.Application app)
        {
            _app = app;
            Text = "Özbekça — istisno so'zlar";
            Font = new Font("Segoe UI", 9f);
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(560, 560);
            MinimumSize = new Size(480, 480);
            KeyPreview = true;
            BuildUi();
            ExceptionStore.Split(ExceptionStore.ReadLines(), _users, _off);
            Render();
        }

        // ------------------------------------------------------------ UI
        void BuildUi()
        {
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(12), ColumnCount = 2, RowCount = 7 };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));     // qo'shish
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));     // xabar
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));     // qidirish
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));     // sarlavhalar
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // ro'yxatlar
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));     // tekshirish
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));     // tugmalar
            Controls.Add(root);

            // 1. qo'shish
            var addRow = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = false };
            _txtNew.Width = 200;
            _txtNew.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; AddWord(_txtNew.Text, true); } };
            var btnAdd = new Button { Text = "Qo'shish", AutoSize = true };
            btnAdd.Click += (s, e) => AddWord(_txtNew.Text, true);
            var btnSel = new Button { Text = "Hujjatda belgilangan so'zni qo'shish", AutoSize = true };
            btnSel.Click += (s, e) => { Msg(AddSelectedWord(_app, this)); };
            addRow.Controls.AddRange(new Control[] { new Label { Text = "Yangi so'z:", AutoSize = true, Margin = new Padding(0, 6, 4, 0) }, _txtNew, btnAdd, btnSel });
            root.Controls.Add(addRow, 0, 0); root.SetColumnSpan(addRow, 2);

            // 2. xabar
            _lblMsg.AutoSize = true; _lblMsg.ForeColor = Color.FromArgb(30, 110, 90); _lblMsg.Margin = new Padding(3, 2, 3, 6);
            root.Controls.Add(_lblMsg, 0, 1); root.SetColumnSpan(_lblMsg, 2);

            // 3. qidirish
            var searchRow = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = false };
            _txtSearch.Width = 260;
            _txtSearch.TextChanged += (s, e) => Render();
            searchRow.Controls.AddRange(new Control[] { new Label { Text = "Qidirish:", AutoSize = true, Margin = new Padding(0, 6, 4, 0) }, _txtSearch });
            root.Controls.Add(searchRow, 0, 2); root.SetColumnSpan(searchRow, 2);

            // 4-5. ro'yxatlar
            _lblUser.AutoSize = true; _lblUser.Font = new Font(Font, FontStyle.Bold);
            _lblDef.AutoSize = true; _lblDef.Font = new Font(Font, FontStyle.Bold);
            root.Controls.Add(_lblUser, 0, 3);
            root.Controls.Add(_lblDef, 1, 3);

            var userPanel = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, Margin = new Padding(0) };
            userPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            userPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _lstUser.Dock = DockStyle.Fill;
            _lstUser.SelectionMode = SelectionMode.MultiExtended;
            _lstUser.Font = new Font("Consolas", 10f);
            _lstUser.KeyDown += (s, e) => { if (e.KeyCode == Keys.Delete) RemoveSelected(); };
            _lstUser.DoubleClick += (s, e) => { if (_lstUser.SelectedItem != null) _txtTest.Text = (string)_lstUser.SelectedItem; };
            var btnDel = new Button { Text = "Tanlanganni o'chirish (Del)", AutoSize = true, Dock = DockStyle.Fill };
            btnDel.Click += (s, e) => RemoveSelected();
            userPanel.Controls.Add(_lstUser, 0, 0);
            userPanel.Controls.Add(btnDel, 0, 1);
            root.Controls.Add(userPanel, 0, 4);

            var defPanel = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, Margin = new Padding(0) };
            defPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            defPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _lstDef.Dock = DockStyle.Fill;
            _lstDef.CheckOnClick = true;
            _lstDef.Font = new Font("Consolas", 10f);
            _lstDef.ItemCheck += OnDefaultCheck;
            defPanel.Controls.Add(_lstDef, 0, 0);
            defPanel.Controls.Add(new Label { Text = "Belgini olib tashlasangiz, bu so'z ham o'giriladi.", AutoSize = true, ForeColor = SystemColors.GrayText }, 0, 1);
            root.Controls.Add(defPanel, 1, 4);

            // 6. tekshirish
            var testRow = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true, Margin = new Padding(0, 8, 0, 0) };
            _txtTest.Width = 180;
            _txtTest.TextChanged += (s, e) => TestWord();
            _lblTest.AutoSize = true; _lblTest.Margin = new Padding(6, 6, 0, 0);
            testRow.Controls.AddRange(new Control[] { new Label { Text = "So'zni tekshirish:", AutoSize = true, Margin = new Padding(0, 6, 4, 0) }, _txtTest, _lblTest });
            root.Controls.Add(testRow, 0, 5); root.SetColumnSpan(testRow, 2);

            // 7. pastki tugmalar
            var bottom = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 10, 0, 0) };
            var btnImport = new Button { Text = "Import...", AutoSize = true };
            btnImport.Click += (s, e) => Import();
            var btnExport = new Button { Text = "Eksport...", AutoSize = true };
            btnExport.Click += (s, e) => Export();
            var btnHelp = new Button { Text = "Shablonlar?", AutoSize = true };
            btnHelp.Click += (s, e) => MessageBox.Show(this,
                "Oddiy so'z:  Schnorr\n" +
                "  kichik harfli yozuv registrga sezgir emas, katta harf bo'lsa - sezgir.\n\n" +
                "phish*  - shu bilan boshlanadigan har qanday so'z (phishing, phisher)\n" +
                "SHA#    - keyin faqat raqam, '-' yoki '/' (SHA-256, SHA3), lekin SHAHAR emas\n\n" +
                "Fayl: " + Converter.ExceptionsPath,
                "Shablonlar", MessageBoxButtons.OK, MessageBoxIcon.Information);
            var btnSave = new Button { Text = "Saqlash", AutoSize = true, DialogResult = DialogResult.OK };
            btnSave.Click += (s, e) => { Save(); };
            var btnCancel = new Button { Text = "Bekor qilish", AutoSize = true, DialogResult = DialogResult.Cancel };
            bottom.Controls.AddRange(new Control[] { btnImport, btnExport, btnHelp, new Label { Width = 40 }, btnSave, btnCancel });
            root.Controls.Add(bottom, 0, 6); root.SetColumnSpan(bottom, 2);

            AcceptButton = null;   // Enter so'z qo'shish uchun
            CancelButton = btnCancel;
            FormClosing += (s, e) =>
            {
                if (DialogResult == DialogResult.Cancel && _dirty &&
                    MessageBox.Show(this, "O'zgarishlar saqlanmagan. Baribir yopilsinmi?", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    e.Cancel = true;
            };
        }

        void Msg(string text) => _lblMsg.Text = text ?? "";

        // ------------------------------------------------------------ holat
        void Render()
        {
            _loading = true;
            string q = _txtSearch.Text.Trim();
            Func<string, bool> hit = w => q.Length == 0 || w.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0;

            _lstUser.BeginUpdate();
            _lstUser.Items.Clear();
            foreach (var w in _users.Where(hit)) _lstUser.Items.Add(w);
            _lstUser.EndUpdate();

            _lstDef.BeginUpdate();
            _lstDef.Items.Clear();
            foreach (var d in Converter.DefaultExceptions.Where(hit))
                _lstDef.Items.Add(d, !IsOff(d));
            _lstDef.EndUpdate();

            int active = Converter.DefaultExceptions.Count(d => !IsOff(d));
            _lblUser.Text = "Mening so'zlarim (" + _users.Count + ")";
            _lblDef.Text = "Standart istisnolar (" + active + " / " + Converter.DefaultExceptions.Length + ")";
            _loading = false;
            TestWord();
        }

        bool IsOff(string d) => _off.Any(o => string.Equals(o, d, StringComparison.OrdinalIgnoreCase));

        void OnDefaultCheck(object sender, ItemCheckEventArgs e)
        {
            if (_loading) return;
            var d = (string)_lstDef.Items[e.Index];
            _off.RemoveAll(o => string.Equals(o, d, StringComparison.OrdinalIgnoreCase));
            if (e.NewValue == CheckState.Unchecked) _off.Add(d);
            _dirty = true;
            Msg(e.NewValue == CheckState.Unchecked ? "«" + d + "» o'chirildi — endi o'giriladi." : "«" + d + "» yoqildi.");
            BeginInvoke((Action)Render);   // ItemCheck ichida ro'yxatni qayta qurib bo'lmaydi
        }

        bool AddWord(string raw, bool clearBox)
        {
            var w = (raw ?? "").Trim();
            if (w.Length == 0) { Msg("So'z kiriting."); return false; }
            if (w.Any(char.IsWhiteSpace)) { Msg("Faqat bitta so'z (bo'shliqsiz)."); return false; }
            if (!w.EndsWith("*") && !w.EndsWith("#")) w = Converter.WordCore(w);
            var def = Converter.DefaultExceptions.FirstOrDefault(d => string.Equals(d, w, StringComparison.OrdinalIgnoreCase));
            if (def != null)
            {
                if (IsOff(def)) { _off.RemoveAll(o => string.Equals(o, def, StringComparison.OrdinalIgnoreCase)); _dirty = true; Msg("«" + def + "» standart ro'yxatda bor — qayta yoqildi."); }
                else Msg("«" + def + "» standart ro'yxatda allaqachon bor.");
            }
            else if (_users.Contains(w)) { Msg("«" + w + "» ro'yxatda allaqachon bor."); return false; }
            else { _users.Add(w); _dirty = true; Msg("«" + w + "» qo'shildi."); }
            if (clearBox) _txtNew.Clear();
            Render();
            return true;
        }

        void RemoveSelected()
        {
            var sel = _lstUser.SelectedItems.Cast<string>().ToList();
            if (sel.Count == 0) { Msg("O'chirish uchun ro'yxatdan so'z tanlang."); return; }
            foreach (var w in sel) _users.Remove(w);
            _dirty = true;
            Msg(sel.Count == 1 ? "«" + sel[0] + "» o'chirildi." : sel.Count + " ta so'z o'chirildi.");
            Render();
        }

        void TestWord()
        {
            var w = _txtTest.Text.Trim();
            if (w.Length == 0) { _lblTest.Text = ""; return; }
            var opts = Settings.Options();
            opts.Exceptions = Converter.MergeExceptions(_off.Select(o => "!" + o).Concat(_users));
            var r = Converter.Explain(w, opts);
            if (r.Item2 != null) { _lblTest.Text = "O'girilmaydi — " + r.Item2; _lblTest.ForeColor = Color.DarkGoldenrod; }
            else if (r.Item1 != w) { _lblTest.Text = w + "  →  " + r.Item1; _lblTest.ForeColor = Color.FromArgb(30, 110, 90); }
            else { _lblTest.Text = "O'zgarmaydi (o'giriladigan harf yo'q)"; _lblTest.ForeColor = SystemColors.GrayText; }
        }

        void Save()
        {
            ExceptionStore.Save(_users, _off);
            _dirty = false;
        }

        void Import()
        {
            using (var dlg = new OpenFileDialog { Filter = "Matn fayli (*.txt)|*.txt|Barcha fayllar|*.*", Title = "Istisnolarni import qilish" })
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                int before = _users.Count;
                ExceptionStore.Split(File.ReadAllLines(dlg.FileName, Encoding.UTF8), _users, _off);
                _dirty = true;
                Msg((_users.Count - before) + " ta yangi so'z import qilindi.");
                Render();
            }
        }

        void Export()
        {
            using (var dlg = new SaveFileDialog { Filter = "Matn fayli (*.txt)|*.txt", FileName = "istisnolar.txt", Title = "Istisnolarni eksport qilish" })
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                var lines = new List<string> { "; Yangi alifbo - istisno so'zlar" };
                lines.AddRange(_off.Select(o => "!" + o));
                lines.AddRange(_users);
                File.WriteAllLines(dlg.FileName, lines, new UTF8Encoding(false));
                Msg("Eksport qilindi: " + dlg.FileName);
            }
        }

        // ------------------------------------------------------------ hujjatdagi so'z
        /// <summary>
        /// Kursor turgan yoki belgilangan so'zni istisnoga qo'shadi. So'z allaqachon
        /// o'girilgan bo'lsa (Şamir), hujjatda asl holiga (Shamir) qaytaradi.
        /// Forma ochiq bo'lsa (form != null) — faylga emas, formaning ro'yxatiga qo'shadi.
        /// </summary>
        internal static string AddSelectedWord(Word.Application app, ExceptionsForm form)
        {
            if (app.Documents.Count == 0) return "Hujjat ochilmagan.";
            var sel = app.Selection;
            Word.Range r = sel.Type == Word.WdSelectionType.wdSelectionIP ? sel.Words[1] : sel.Range.Duplicate;
            string w = r.Text ?? "";
            int lead = 0, trail = 0;
            while (lead < w.Length && !char.IsLetterOrDigit(w[lead])) lead++;
            while (trail < w.Length - lead && !char.IsLetterOrDigit(w[w.Length - 1 - trail])) trail++;
            if (lead + trail >= w.Length) return "Bitta so'zni belgilang yoki kursorni so'z ustiga qo'ying.";
            r.MoveStart(Word.WdUnits.wdCharacter, lead);
            r.MoveEnd(Word.WdUnits.wdCharacter, -trail);
            w = r.Text;
            if (w.Any(char.IsWhiteSpace)) return "Faqat bitta so'zni belgilang.";

            string orig = Converter.RevertWord(w);
            if (orig != w) r.Text = orig;
            string note = orig != w ? "  Hujjatda asl holiga qaytarildi: " + w + " → " + orig : "";

            if (form != null)
            {
                form.AddWord(orig, false);
                return form._lblMsg.Text + note;
            }
            var core = Converter.WordCore(orig);
            if (Converter.DefaultExceptions.Any(d => string.Equals(d, core, StringComparison.OrdinalIgnoreCase)))
            {
                var users = new List<string>(); var off = new List<string>();
                ExceptionStore.Split(ExceptionStore.ReadLines(), users, off);
                if (off.RemoveAll(o => string.Equals(o, core, StringComparison.OrdinalIgnoreCase)) > 0)
                {
                    ExceptionStore.Save(users, off);
                    return "«" + core + "» standart istisno qayta yoqildi." + note;
                }
                return "«" + core + "» standart ro'yxatda allaqachon bor." + note;
            }
            return (ExceptionStore.Add(core) ? "«" + core + "» istisnolarga qo'shildi." : "«" + core + "» ro'yxatda allaqachon bor.") + note;
        }
    }
}
