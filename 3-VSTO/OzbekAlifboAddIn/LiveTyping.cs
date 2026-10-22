using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace OzbekAlifboAddIn
{
    /// <summary>
    /// Yozish paytida o'girish. Word oqimiga WH_GETMESSAGE hook o'rnatiladi:
    /// h/H yoki apostrof (WM_CHAR) kelganda, Word uni joylagandan keyin
    /// kursordan oldingi 2 belgi tekshiriladi:  sh→ş, ch→ç, o'→ö, g'→ğ.
    /// Qaytarish: tugmani yana bosing (ş + h → sh, ö + ' → o') yoki Ctrl+Z.
    /// Klaviatura tilidan qat'i nazar ishlaydi, chunki haqiqiy belgiga qaraydi.
    /// </summary>
    internal sealed class LiveTyping : IDisposable
    {
        delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr SetWindowsHookEx(int idHook, HookProc fn, IntPtr hMod, uint threadId);
        [DllImport("user32.dll")]
        static extern bool UnhookWindowsHookEx(IntPtr hhk);
        [DllImport("user32.dll")]
        static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("kernel32.dll")]
        static extern uint GetCurrentThreadId();

        [StructLayout(LayoutKind.Sequential)]
        struct MSG
        {
            public IntPtr hwnd;
            public uint message;
            public IntPtr wParam;
            public IntPtr lParam;
            public uint time;
            public int ptX, ptY;
        }

        const int WH_GETMESSAGE = 3;
        const int HC_ACTION = 0;
        const int PM_REMOVE = 1;
        const uint WM_CHAR = 0x0102;

        readonly Word.Application _app;
        readonly Control _marshal;     // UI oqimiga qaytish uchun
        HookProc _proc;                // GC yig'ib olmasligi uchun saqlanadi
        IntPtr _hook = IntPtr.Zero;
        bool _pending;

        public LiveTyping(Word.Application app)
        {
            _app = app;
            _marshal = new Control();
            _marshal.CreateControl();
            var _ = _marshal.Handle;   // oyna dastagini yaratish
        }

        public bool Enabled => _hook != IntPtr.Zero;

        public void Enable()
        {
            if (Enabled) return;
            _proc = HookCallback;
            _hook = SetWindowsHookEx(WH_GETMESSAGE, _proc, IntPtr.Zero, GetCurrentThreadId());
        }

        public void Disable()
        {
            if (!Enabled) return;
            UnhookWindowsHookEx(_hook);
            _hook = IntPtr.Zero;
        }

        IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode == HC_ACTION && wParam.ToInt64() == PM_REMOVE)
            {
                var msg = (MSG)Marshal.PtrToStructure(lParam, typeof(MSG));
                if (msg.message == WM_CHAR)
                {
                    char ch = (char)msg.wParam.ToInt64();
                    if ((ch == 'h' || ch == 'H' || Converter.IsApos(ch)) && !_pending)
                    {
                        _pending = true;
                        // WM_CHAR qayta ishlangach bajariladi (post qilingan xabar navbatda keyin keladi)
                        _marshal.BeginInvoke((Action)Check);
                    }
                }
            }
            return CallNextHookEx(_hook, nCode, wParam, lParam);
        }

        void Check()
        {
            _pending = false;
            try
            {
                var sel = _app.Selection;
                if (sel == null || sel.Type != Word.WdSelectionType.wdSelectionIP) return;
                int pos = sel.Start;
                if (pos < 2) return;
                Word.Range r = sel.Range.Duplicate;
                r.SetRange(pos - 2, pos);
                string s = r.Text;
                if (s == null || s.Length != 2) return;
                if (Converter.IsMonoFont(r.Font.Name)) return;
                string rep = Rule(s[0], s[1]);
                if (rep == null) return;
                r.Text = rep;
                // kursor yangi harfdan keyin tursin (aks holda keyingi belgidan oshib ketadi)
                int np = pos - 2 + rep.Length;
                sel.SetRange(np, np);
            }
            catch { /* hujjat yopilgan yoki tahrirlab bo'lmaydi */ }
        }

        /// <summary>prev+last juftligi uchun almashtirish (yoki null).</summary>
        internal static string Rule(char prev, char last)
        {
            if (last == 'h' || last == 'H')
            {
                switch (prev)
                {
                    case 's': return "ş";
                    case 'S': return "Ş";
                    case 'c': return "ç";
                    case 'C': return "Ç";
                    case 'ş': return "s" + last;   // qaytarish
                    case 'Ş': return "S" + last;
                    case 'ç': return "c" + last;
                    case 'Ç': return "C" + last;
                }
            }
            else if (Converter.IsApos(last))
            {
                switch (prev)
                {
                    case 'o': return "ö";
                    case 'O': return "Ö";
                    case 'g': return "ğ";
                    case 'G': return "Ğ";
                    case 'ö': return "o" + last;   // qaytarish
                    case 'Ö': return "O" + last;
                    case 'ğ': return "g" + last;
                    case 'Ğ': return "G" + last;
                }
            }
            return null;
        }

        public void Dispose()
        {
            Disable();
            _marshal.Dispose();
        }
    }
}
