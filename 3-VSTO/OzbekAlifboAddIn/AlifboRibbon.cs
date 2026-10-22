using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace OzbekAlifboAddIn
{
    [ComVisible(true)]
    public class AlifboRibbon : Office.IRibbonExtensibility
    {
        Office.IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonID)
        {
            var asm = Assembly.GetExecutingAssembly();
            foreach (var name in asm.GetManifestResourceNames())
                if (name.EndsWith("AlifboRibbon.xml", StringComparison.OrdinalIgnoreCase))
                    using (var r = new StreamReader(asm.GetManifestResourceStream(name)))
                        return r.ReadToEnd();
            return null;
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI) => _ribbon = ribbonUI;

        // ---------------- tugmalar
        public void OnHujjat(Office.IRibbonControl c) => Convert(false);
        public void OnTanlangan(Office.IRibbonControl c) => Convert(true);
        public void OnKirillHujjat(Office.IRibbonControl c) => Convert(false, ConvertMode.Cyrillic);
        public void OnKirillTanlangan(Office.IRibbonControl c) => Convert(true, ConvertMode.Cyrillic);
        public void OnLotinKirillHujjat(Office.IRibbonControl c) => Convert(false, ConvertMode.LatinToCyrillic);
        public void OnLotinKirillTanlangan(Office.IRibbonControl c) => Convert(true, ConvertMode.LatinToCyrillic);
        public void OnEskiHujjat(Office.IRibbonControl c) => Convert(false, ConvertMode.NewToOld);
        public void OnEskiTanlangan(Office.IRibbonControl c) => Convert(true, ConvertMode.NewToOld);

        /// <summary>Lenta ikonkalari: icons/*.png (Embedded Resource).</summary>
        public System.Drawing.Bitmap LoadImage(string imageId)
        {
            var asm = Assembly.GetExecutingAssembly();
            var name = asm.GetManifestResourceNames().FirstOrDefault(n => n.EndsWith("." + imageId + ".png", StringComparison.OrdinalIgnoreCase));
            if (name == null) return null;
            using (var st = asm.GetManifestResourceStream(name)) return new System.Drawing.Bitmap(st);
        }

        void Convert(bool onlySelection, ConvertMode mode = ConvertMode.Latin)
        {
            var app = Globals.ThisAddIn.Application;
            if (app.Documents.Count == 0) return;
            try
            {
                var res = DocumentConverter.Run(app, onlySelection, mode);
                MessageBox.Show(
                    res.Changed + (mode == ConvertMode.Cyrillic ? " ta so'z kirilldan o'girildi."
                        : mode == ConvertMode.LatinToCyrillic ? " ta so'z kirillga o'girildi."
                        : mode == ConvertMode.NewToOld ? " ta so'z eski lotinga o'girildi." : " ta harf o'girildi.") +
                    (res.Skipped > 0 ? "\n" + res.Skipped + " ta joy o'tkazib yuborildi (maydon yoki yashirin matn)." : "") +
                    "\nBekor qilish: Ctrl+Z",
                    "Özbekça", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Özbekça", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // ---------------- yozish paytida
        public bool GetLivePressed(Office.IRibbonControl c) => Globals.ThisAddIn.Live.Enabled;

        // ---------------- tezkor tugmalar uchun amallar
        /// <summary>Matn belgilangan bo'lsa — faqat uni, aks holda butun hujjatni o'giradi.</summary>
        internal void ConvertAuto(ConvertMode mode)
        {
            var app = Globals.ThisAddIn.Application;
            if (app.Documents.Count == 0) return;
            Convert(app.Selection.Type != Word.WdSelectionType.wdSelectionIP, mode);
        }

        internal void ToggleLive()
        {
            var live = Globals.ThisAddIn.Live;
            OnLive(null, !live.Enabled);
            _ribbon?.InvalidateControl("tglYozishda");
            Globals.ThisAddIn.Application.StatusBar = "Özbekça: yozishda o'girish " + (live.Enabled ? "yoqildi" : "o'chirildi");
        }

        /// <summary>Kursor turgan so'z kirillcha bo'lsa — kirill istisnolari oynasi, aks holda lotin istisnosi.</summary>
        internal void AddExceptionAuto()
        {
            var app = Globals.ThisAddIn.Application;
            if (app.Documents.Count == 0) return;
            var sel = app.Selection;
            string w = sel.Type == Word.WdSelectionType.wdSelectionIP ? sel.Words[1].Text : sel.Text;
            if ((w ?? "").Any(ch => ch >= '\u0400' && ch <= '\u04FF'))
            {
                using (var f = new CyrExceptionsForm(app)) { f.PrefillFromSelection(); f.ShowDialog(); }
            }
            else OnIstisnoQoshish(null);
        }

        internal void OpenExceptions() => OnIstisnolar(null);

        internal static void ShowHelp()
        {
            MessageBox.Show(
                "Tezkor tugmalar:\n\n" +
                "   Alt+A — yangi alifboga (belgilangan matn yoki butun hujjat)\n" +
                "   Alt+E — eski (amaldagi) lotinga\n" +
                "   Alt+K — kirilldan yangi lotinga\n" +
                "   Alt+C — lotindan kirillga\n" +
                "   Alt+Z — yozishda o'girishni yoqish/o'chirish\n" +
                "   Alt+I — kursor turgan so'zni istisnoga qo'shish\n" +
                "   Alt+O — istisnolar oynasi\n\n" +
                "Istisnolar: " + Converter.ExceptionsPath + "\n" +
                "Kirill istisnolari: " + Converter.CyrExceptionsPath,
                "Özbekça — yordam", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void OnYordam(Office.IRibbonControl c) => ShowHelp();
        public bool GetHotkeysPressed(Office.IRibbonControl c) => Settings.Hotkeys;
        public void OnHotkeys(Office.IRibbonControl c, bool pressed)
        {
            Settings.Hotkeys = pressed;
            if (pressed) Globals.ThisAddIn.Keys.Enable(); else Globals.ThisAddIn.Keys.Disable();
        }

        public void OnLive(Office.IRibbonControl c, bool pressed)
        {
            if (pressed) Globals.ThisAddIn.Live.Enable(); else Globals.ThisAddIn.Live.Disable();
            Settings.LiveTyping = pressed;
        }

        // ---------------- sozlamalar
        public bool GetEnglishPressed(Office.IRibbonControl c) => Settings.SkipEnglish;
        public void OnEnglish(Office.IRibbonControl c, bool pressed) => Settings.SkipEnglish = pressed;

        public bool GetRussianPressed(Office.IRibbonControl c) => Settings.SkipRussian;
        public void OnRussian(Office.IRibbonControl c, bool pressed) => Settings.SkipRussian = pressed;

        public bool GetTutuqPressed(Office.IRibbonControl c) => Settings.Tutuq;
        public void OnTutuq(Office.IRibbonControl c, bool pressed) => Settings.Tutuq = pressed;

        // ---------------- istisnolar
        public void OnIstisnolar(Office.IRibbonControl c)
        {
            using (var f = new ExceptionsForm(Globals.ThisAddIn.Application))
                f.ShowDialog();
        }

        public void OnKirillIstisnolar(Office.IRibbonControl c)
        {
            using (var f = new CyrExceptionsForm(Globals.ThisAddIn.Application))
                f.ShowDialog();
        }

        public void OnIstisnoQoshish(Office.IRibbonControl c)
        {
            try
            {
                var msg = ExceptionsForm.AddSelectedWord(Globals.ThisAddIn.Application, null);
                Globals.ThisAddIn.Application.StatusBar = msg;
                MessageBox.Show(msg, "Özbekça", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Özbekça", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
