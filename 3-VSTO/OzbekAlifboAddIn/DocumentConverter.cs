using System;
using Word = Microsoft.Office.Interop.Word;

namespace OzbekAlifboAddIn
{
    internal enum ConvertMode { Latin, Cyrillic, LatinToCyrillic, NewToOld }

    internal sealed class ConvertResult
    {
        public int Changed;
        public int Skipped;
    }

    /// <summary>Hujjatni yoki belgilangan matnni yangi alifboga o'giradi.</summary>
    internal static class DocumentConverter
    {
        public static ConvertResult Run(Word.Application app, bool onlySelection, ConvertMode mode = ConvertMode.Latin)
        {
            var res = new ConvertResult();
            var doc = app.ActiveDocument;
            if (doc.ProtectionType != Word.WdProtectionType.wdNoProtection)
                throw new InvalidOperationException("Hujjat himoyalangan - avval himoyani olib tashlang.");

            var opts = Settings.Options();
            var undo = app.UndoRecord;
            app.ScreenUpdating = false;
            undo.StartCustomRecord(mode == ConvertMode.Cyrillic ? "Kirilldan yangi lotinga"
                : mode == ConvertMode.LatinToCyrillic ? "Lotindan kirillga"
                : mode == ConvertMode.NewToOld ? "Eski lotinga" : "Yangi alifboga o'girish");
            try
            {
                if (onlySelection)
                {
                    var sel = app.Selection;
                    if (sel.Type == Word.WdSelectionType.wdSelectionIP)
                        throw new InvalidOperationException("Avval o'giriladigan matnni belgilang.");
                    var limit = sel.Range.Duplicate;
                    ProcessRange(sel.Range, limit, opts, res, app, mode);
                }
                else
                {
                    foreach (Word.Range story in doc.StoryRanges)
                    {
                        Word.Range r = story;
                        while (r != null)
                        {
                            ProcessRange(r, null, opts, res, app, mode);
                            r = r.NextStoryRange;
                        }
                    }
                }
            }
            finally
            {
                undo.EndCustomRecord();
                app.ScreenUpdating = true;
                app.StatusBar = "";
            }
            return res;
        }

        static void ProcessRange(Word.Range rng, Word.Range limit, ConverterOptions opts, ConvertResult res, Word.Application app, ConvertMode mode)
        {
            int np = 0;
            foreach (Word.Paragraph p in rng.Paragraphs)
            {
                if (++np % 100 == 0) app.StatusBar = "Özbekça: " + np + "-abzas...";

                Word.Range pr = p.Range;
                string text = pr.Text;
                int ps = pr.Start;
                string font = pr.Font.Name;
                bool mixed = string.IsNullOrEmpty(font);
                if (Converter.IsMonoFont(font)) continue;

                var edits = mode == ConvertMode.Cyrillic
                    ? Converter.ComputeCyrillicEdits(text, Settings.SkipRussian, opts.CyrExceptions)
                    : mode == ConvertMode.LatinToCyrillic
                        ? Converter.ComputeLatCyrEdits(text, opts)
                        : mode == ConvertMode.NewToOld
                            ? Converter.ComputeNewToOldEdits(text)
                            : Converter.ComputeEdits(text, opts);
                for (int k = edits.Count - 1; k >= 0; k--)
                {
                    var e = edits[k];
                    int a = ps + e.Pos, b = a + e.Len;
                    if (limit != null && (a < limit.Start || b > limit.End)) continue;

                    Word.Range r = p.Range.Duplicate;
                    r.SetRange(a, b);
                    if (r.Text != e.From) { res.Skipped++; continue; }   // maydon / yashirin matn
                    if (mixed && Converter.IsMonoFont(r.Font.Name)) continue;
                    r.Text = e.To;
                    res.Changed++;
                }
            }
        }
    }
}
