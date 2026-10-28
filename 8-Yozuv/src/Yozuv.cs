// Özbekça Yozuv — Windows'ning istalgan dasturida yozish paytida yangi alifboga o'girish.
// Soat yonida (tray) ikonka bo'lib turadi. Alt+Shift+Z — yoqish/o'chirish.
// Hech qanday tugma faylga yozilmaydi va hech qayerga yuborilmaydi: dastur faqat oxirgi so'zni
// xotirada ushlab turadi (o'girish va istisnoni tekshirish uchun).
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("Özbekça Yozuv 1.0")]
[assembly: AssemblyDescription("Istalgan dasturda yozişda yangi özbek alifbosiga ögiriş")]
[assembly: AssemblyCompany("Mirzo Uluğbek nomidagi Özbekiston Milliy universiteti")]
[assembly: AssemblyProduct("Özbekça Yozuv")]
[assembly: AssemblyCopyright("© 2026 Bozorov Obidjon Norqobilovich")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]
[assembly: AssemblyInformationalVersion("1.0")]

namespace Ozbekca
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            bool yangi;
            using (var m = new Mutex(true, "Local\\Ozbekca_Yozuv_yagona", out yangi))
            {
                if (!yangi)
                {
                    MessageBox.Show("Özbekça Yozuv allaqaçon işlamoqda.\nSoat yonidagi Ö ikonkasiga qarang.",
                        App.Nomi, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                using (var app = new App()) Application.Run();
            }
        }
    }

    sealed class App : IDisposable
    {
        public const string Nomi = "Özbekça Yozuv";
        const string Versiya = "1.0";
        const string RegKey = @"Software\Ozbekca\Yozuv";
        const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
        const string RunName = "OzbekcaYozuv";
        static readonly UIntPtr MARK = new UIntPtr(0x0E2BECA);

        readonly Engine engine = new Engine();
        readonly NotifyIcon tray = new NotifyIcon();
        readonly ToolStripMenuItem miYoq, miAvto;
        readonly Icon icoOn, icoOff;
        readonly HotkeyWindow hk;
        volatile bool yoqilgan = true;

        // ilgaklar (hook) — delegatlar GC tomonidan yig'ilmasligi uchun maydonda saqlanadi
        readonly W.LowLevelProc kbProc, msProc;
        readonly W.WinEventProc evProc;
        IntPtr kbHook, msHook, evFg, evFocus;
        Thread hookThread;
        int hookErr;

        // joriy oyna holati
        IntPtr curFg = IntPtr.Zero;
        volatile bool curSkip;            // istisno dastur / administrator oynasi
        volatile bool parolMaydoni;       // fokusdagi maydon — parol
        readonly Dictionary<uint, bool> pidSkip = new Dictionary<uint, bool>();
        HashSet<string> skipApps = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        DateTime appsMtime = DateTime.MinValue, excMtime = DateTime.MinValue;

        static string Dir { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OzbekAlifbo"); } }
        static string AppsFile { get { return Path.Combine(Dir, "yozuv-dasturlar.txt"); } }
        static string ExcFile { get { return Path.Combine(Dir, "istisnolar.txt"); } }

        public App()
        {
            icoOn = LoadIcon();
            icoOff = Grey(icoOn);

            var menu = new ContextMenuStrip();
            miYoq = new ToolStripMenuItem("Yoqilgan  (Alt+Shift+Z)", null, (s, e) => SetOn(!yoqilgan));
            miYoq.Font = new Font(miYoq.Font, FontStyle.Bold);
            miAvto = new ToolStripMenuItem("Windows bilan birga işga tuşsin", null, (s, e) => SetAutostart(!miAvto.Checked));
            menu.Items.Add(miYoq);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Istisno sözlar…", null, (s, e) => OpenFile(ExcFile, ExcTemplate()));
            menu.Items.Add("Istisno dasturlar…", null, (s, e) => OpenFile(AppsFile, AppsTemplate()));
            menu.Items.Add(miAvto);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Yordam", null, (s, e) => Help());
            menu.Items.Add("Çiqiş", null, (s, e) => Application.Exit());
            tray.ContextMenuStrip = menu;
            tray.MouseClick += (s, e) => { if (e.Button == MouseButtons.Left) SetOn(!yoqilgan); };

            ReloadFiles(true);

            // sozlamalar
            bool first;
            using (var k = Registry.CurrentUser.CreateSubKey(RegKey))
            {
                first = k.GetValue("Yoqilgan") == null;
                yoqilgan = (k.GetValue("Yoqilgan", "1") as string) != "0";
                if (first) k.SetValue("Yoqilgan", "1");
            }
            if (first) SetAutostart(true);
            else RefreshAutostart();

            kbProc = KbHook; msProc = MsHook; evProc = WinEvent;
            // Klaviatura ilgagi alohida oqimda: menyu, MSAA so'rovlari va oynalar uni sekinlashtirmasin
            // (aks holda Windows sekin ilgakni o'zi uzib qo'yadi).
            var ready = new ManualResetEvent(false);
            hookThread = new Thread(() =>
            {
                IntPtr hMod = W.GetModuleHandle(null);
                kbHook = W.SetWindowsHookEx(W.WH_KEYBOARD_LL, kbProc, hMod, 0);
                hookErr = Marshal.GetLastWin32Error();
                msHook = W.SetWindowsHookEx(W.WH_MOUSE_LL, msProc, hMod, 0);
                ready.Set();
                Application.Run();
                if (kbHook != IntPtr.Zero) W.UnhookWindowsHookEx(kbHook);
                if (msHook != IntPtr.Zero) W.UnhookWindowsHookEx(msHook);
            });
            hookThread.IsBackground = true;
            hookThread.SetApartmentState(ApartmentState.STA);
            hookThread.Start();
            ready.WaitOne(5000);
            evFg = W.SetWinEventHook(W.EVENT_SYSTEM_FOREGROUND, W.EVENT_SYSTEM_FOREGROUND, IntPtr.Zero, evProc, 0, 0, W.WINEVENT_OUTOFCONTEXT | W.WINEVENT_SKIPOWNPROCESS);
            evFocus = W.SetWinEventHook(W.EVENT_OBJECT_FOCUS, W.EVENT_OBJECT_FOCUS, IntPtr.Zero, evProc, 0, 0, W.WINEVENT_OUTOFCONTEXT | W.WINEVENT_SKIPOWNPROCESS);

            hk = new HotkeyWindow(() => SetOn(!yoqilgan));
            bool hkOk = hk.Register(W.MOD_ALT | W.MOD_SHIFT | W.MOD_NOREPEAT, (uint)Keys.Z);

            UpdateUi();
            tray.Visible = true;

            if (kbHook == IntPtr.Zero)
            {
                MessageBox.Show("Klaviatura bilan boğlanib bölmadi (xato " + hookErr + ").\nDasturni qayta işga tuşiring.",
                    Nomi, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            if (first)
            {
                tray.ShowBalloonTip(6000, Nomi,
                    "Endi istalgan dasturda o' → ö, g' → ğ, sh → ş, ch → ç özi ögiriladi.\n" +
                    (hkOk ? "Alt+Shift+Z yoki ikonkani bosish — yoqiş/öçiriş." : "Ikonkani bosish — yoqiş/öçiriş."),
                    ToolTipIcon.Info);
            }
        }

        // ---------------------------------------------------------------- UI
        void SetOn(bool on)
        {
            yoqilgan = on;
            lock (engine) engine.Reset();
            try { using (var k = Registry.CurrentUser.CreateSubKey(RegKey)) k.SetValue("Yoqilgan", on ? "1" : "0"); } catch { }
            UpdateUi();
            tray.ShowBalloonTip(1500, Nomi, on ? "Yoqildi — yozişda yangi alifboga ögiriladi." : "Öçirildi — harflar ögirilmaydi.", ToolTipIcon.None);
        }

        void UpdateUi()
        {
            tray.Icon = yoqilgan ? icoOn : icoOff;
            tray.Text = Nomi + (yoqilgan ? " — yoqilgan" : " — öçirilgan") + "\nAlt+Shift+Z — almaştiriş";
            miYoq.Checked = yoqilgan;
        }

        static string ExePath { get { return Application.ExecutablePath; } }

        void SetAutostart(bool on)
        {
            try
            {
                using (var k = Registry.CurrentUser.CreateSubKey(RunKey))
                {
                    if (on) k.SetValue(RunName, "\"" + ExePath + "\"");
                    else if (k.GetValue(RunName) != null) k.DeleteValue(RunName);
                }
            }
            catch { }
            RefreshAutostart();
        }

        void RefreshAutostart()
        {
            try
            {
                using (var k = Registry.CurrentUser.OpenSubKey(RunKey, true))
                {
                    var v = k == null ? null : k.GetValue(RunName) as string;
                    // dastur boshqa joyga ko'chirilgan bo'lsa — yo'lni yangilaymiz
                    if (v != null && k != null && v.Trim('"') != ExePath) k.SetValue(RunName, "\"" + ExePath + "\"");
                    miAvto.Checked = v != null;
                }
            }
            catch { miAvto.Checked = false; }
        }

        void Help()
        {
            MessageBox.Show(
                Nomi + " " + Versiya + "\n\n" +
                "Istalgan dasturda (Telegram, brauzer, Bloknot, …) odatdagidek yozing:\n" +
                "    o' → ö      g' → ğ      sh → ş      ch → ç\n\n" +
                "• Ögirilmasin desangiz — harfni yana bir bor bosing:\n" +
                "    s h h → sh,     o ' ' → o'\n" +
                "• Istisno sözlar (Photoshop, Chevrolet, English, …) söz tugagaç asl holiga qaytadi.\n" +
                "  Öz sözlaringizni «Istisno sözlar…» orqali qöşing (Word plagini bilan umumiy).\n" +
                "• Parol maydonlarida va administrator oynalarida işlamaydi.\n" +
                "• Word, Excel, PowerPoint'da Özbekça plaginining öz «Yozişda» rejimi bor,\n" +
                "  şuning uçun ular «Istisno dasturlar» röyxatida.\n\n" +
                "Alt+Shift+Z yoki ikonkani bosish — yoqiş / öçiriş.\n\n" +
                "Dastur bosilgan tugmalarni hech qayerga yozmaydi va yubormaydi.\n\n" +
                "Muallif: Bozorov Obidjon Norqobilovich  ·  o.bozorov@nuu.uz\n" +
                "Mirzo Uluğbek nomidagi ÖzMU, Axborot xavfsizligi kafedrasi\n" +
                "github.com/obidbozorov/ozbek-alifbo",
                Nomi, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        static string AppsTemplate()
        {
            return "; Özbekça Yozuv — bu dasturlarda harflar ögirilmaydi.\r\n" +
                   "; Har qatorda bitta dastur nomi (.exe siz). Saqlagaç darhol kuçga kiradi.\r\n" +
                   "; Dastur nomini bilish: Vazifalar dispetçeri (Ctrl+Shift+Esc) → Tafsilotlar.\r\n" +
                   "\r\n" +
                   "; Özbekça plaginining öz «Yozişda» rejimi bor\r\n" +
                   "winword\r\nexcel\r\npowerpnt\r\n\r\n" +
                   "; dasturlaş va buyruq satri\r\n" +
                   "cmd\r\npowershell\r\npwsh\r\nWindowsTerminal\r\nconhost\r\nCode\r\ndevenv\r\nidea64\r\npycharm64\r\nsublime_text\r\n\r\n" +
                   "; parol menejerlari\r\n" +
                   "KeePass\r\nKeePassXC\r\n1Password\r\nBitwarden\r\n";
        }

        static string ExcTemplate()
        {
            return "; Istisno sözlar — ögirilmaydi (Word, Excel, PowerPoint plagini va Özbekça Yozuv uçun umumiy).\r\n" +
                   "; Har qatorda bitta söz. Katta harf bölsa — harf kattaligi hisobga olinadi.\r\n" +
                   ";   söz*   — şu bilan boşlanadigan barça sözlar\r\n" +
                   ";   !söz   — standart istisnoni öçiriş\r\n" +
                   "\r\n";
        }

        void OpenFile(string path, string template)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                if (!File.Exists(path)) File.WriteAllText(path, template, new UTF8Encoding(true));
                Process.Start("notepad.exe", "\"" + path + "\"");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Faylni oçib bölmadi:\n" + ex.Message, Nomi, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        void ReloadFiles(bool force)
        {
            try
            {
                if (!File.Exists(AppsFile))
                {
                    Directory.CreateDirectory(Dir);
                    File.WriteAllText(AppsFile, AppsTemplate(), new UTF8Encoding(true));
                }
                var t = File.GetLastWriteTimeUtc(AppsFile);
                if (force || t != appsMtime)
                {
                    appsMtime = t;
                    var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var raw in File.ReadAllLines(AppsFile, Encoding.UTF8))
                    {
                        var s = raw.Trim().TrimStart('﻿');
                        if (s.Length == 0 || s[0] == ';') continue;
                        if (s.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)) s = s.Substring(0, s.Length - 4);
                        set.Add(s);
                    }
                    skipApps = set;
                    lock (pidSkip) pidSkip.Clear();
                }
            }
            catch { }
            try
            {
                var t = File.Exists(ExcFile) ? File.GetLastWriteTimeUtc(ExcFile) : DateTime.MinValue;
                if (force || t != excMtime) { excMtime = t; lock (engine) engine.IstisnolarniYukla(ExcFile); }
            }
            catch { }
        }

        static Icon LoadIcon()
        {
            try
            {
                using (var s = Assembly.GetExecutingAssembly().GetManifestResourceStream("tray.ico"))
                    if (s != null) return new Icon(s, SystemInformation.SmallIconSize);
            }
            catch { }
            return SystemIcons.Application;
        }

        static Icon Grey(Icon src)
        {
            try
            {
                using (var b = src.ToBitmap())
                using (var g = new Bitmap(b.Width, b.Height))
                {
                    using (var gr = Graphics.FromImage(g))
                    {
                        var cm = new ColorMatrix(new[] {
                            new float[] { .3f, .3f, .3f, 0, 0 }, new float[] { .59f, .59f, .59f, 0, 0 },
                            new float[] { .11f, .11f, .11f, 0, 0 }, new float[] { 0, 0, 0, .55f, 0 }, new float[] { 0, 0, 0, 0, 1 } });
                        var ia = new ImageAttributes();
                        ia.SetColorMatrix(cm);
                        gr.DrawImage(b, new Rectangle(0, 0, b.Width, b.Height), 0, 0, b.Width, b.Height, GraphicsUnit.Pixel, ia);
                    }
                    return Icon.FromHandle(g.GetHicon());
                }
            }
            catch { return src; }
        }

        // ---------------------------------------------------------------- oyna holati
        void WinEvent(IntPtr hook, uint ev, IntPtr hwnd, int idObject, int idChild, uint thread, uint time)
        {
            try
            {
                lock (engine) engine.Reset();
                if (ev == W.EVENT_SYSTEM_FOREGROUND)
                {
                    ReloadFiles(false);
                    parolMaydoni = false;
                }
                else if (ev == W.EVENT_OBJECT_FOCUS)
                {
                    bool p = !curSkip && yoqilgan && IsPassword(hwnd, idObject, idChild);
                    parolMaydoni = p;
                    lock (engine) engine.Reset();
                }
            }
            catch { }
        }

        bool ShouldSkip(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return true;
            uint pid;
            W.GetWindowThreadProcessId(hwnd, out pid);
            lock (pidSkip)
            {
                bool v;
                if (pidSkip.TryGetValue(pid, out v)) return v;
                v = ComputeSkip(pid);
                if (pidSkip.Count > 500) pidSkip.Clear();
                pidSkip[pid] = v;
                return v;
            }
        }

        bool ComputeSkip(uint pid)
        {
            if (pid == (uint)Process.GetCurrentProcess().Id) return true;
            try
            {
                string name;
                using (var p = Process.GetProcessById((int)pid)) name = p.ProcessName;
                if (skipApps.Contains(name)) return true;
            }
            catch { return true; }
            // administrator huquqidagi oynaga tugma yuborib bo'lmaydi (UIPI) — harf yo'qolmasligi uchun tegmaymiz
            return !W.SameOrLowerIntegrity(pid);
        }

        static bool IsPassword(IntPtr hwnd, int idObject, int idChild)
        {
            try
            {
                var cls = new StringBuilder(64);
                W.GetClassName(hwnd, cls, cls.Capacity);
                if (cls.ToString().IndexOf("Edit", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    (W.GetWindowLong(hwnd, W.GWL_STYLE) & W.ES_PASSWORD) != 0) return true;

                Accessibility.IAccessible acc;
                object child;
                if (W.AccessibleObjectFromEvent(hwnd, idObject, idChild, out acc, out child) != 0 || acc == null) return false;
                object st = acc.get_accState(child);
                return st is int && (((int)st) & W.STATE_SYSTEM_PROTECTED) != 0;
            }
            catch { return false; }
        }

        // ---------------------------------------------------------------- ilgaklar
        IntPtr MsHook(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                int m = wParam.ToInt32();
                if (m == W.WM_LBUTTONDOWN || m == W.WM_RBUTTONDOWN || m == W.WM_MBUTTONDOWN || m == W.WM_MOUSEWHEEL)
                    lock (engine) engine.Reset();
            }
            return W.CallNextHookEx(msHook, nCode, wParam, lParam);
        }

        readonly byte[] ks = new byte[256];
        readonly StringBuilder ub = new StringBuilder(8);

        IntPtr KbHook(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (nCode >= 0 && yoqilgan)
                {
                    int msg = wParam.ToInt32();
                    if (msg == W.WM_KEYDOWN || msg == W.WM_SYSKEYDOWN)
                    {
                        var k = (W.KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(W.KBDLLHOOKSTRUCT));
                        if (k.dwExtraInfo != MARK) { lock (engine) { if (Handle(k)) return (IntPtr)1; } }
                    }
                }
            }
            catch { lock (engine) engine.Reset(); }
            return W.CallNextHookEx(kbHook, nCode, wParam, lParam);
        }

        static bool Down(int vk) { return (W.GetAsyncKeyState(vk) & 0x8000) != 0; }

        // true — asl tugma yutildi (o'rniga o'zimiz yozdik)
        bool Handle(W.KBDLLHOOKSTRUCT k)
        {
            int vk = (int)k.vkCode;
            switch (vk)
            {
                case 0x10: case 0x11: case 0x12: case 0xA0: case 0xA1: case 0xA2: case 0xA3: case 0xA4: case 0xA5:
                case 0x14: case 0x90: case 0x91: // Shift, Ctrl, Alt, CapsLock, NumLock, ScrollLock
                    return false;
            }

            IntPtr fg = W.GetForegroundWindow();
            if (fg != curFg) { curFg = fg; curSkip = ShouldSkip(fg); parolMaydoni = false; engine.Reset(); }
            if (curSkip || parolMaydoni) return false;

            bool ctrl = Down(0x11), alt = Down(0x12), win = Down(0x5B) || Down(0x5C);
            if (win || (ctrl != alt)) { engine.Reset(); return false; }   // Ctrl+C, Alt+Tab, Win+… — tegmaymiz
            bool altGr = ctrl && alt;

            if (vk == 0x08) { if (altGr) engine.Reset(); else engine.Backspace(); return false; }

            Array.Clear(ks, 0, 256);
            if (Down(0x10)) ks[0x10] = 0x80;
            if ((W.GetKeyState(0x14) & 1) != 0) ks[0x14] = 0x01;
            if (ctrl) ks[0x11] = 0x80;
            if (alt) ks[0x12] = 0x80;

            uint pid;
            uint tid = W.GetWindowThreadProcessId(fg, out pid);
            IntPtr hkl = W.GetKeyboardLayout(tid);
            ub.Length = 0;
            int n = W.ToUnicodeEx(k.vkCode, k.scanCode, ks, ub, ub.Capacity, 4 /* holatni o'zgartirmaslik */, hkl);
            if (n != 1) { engine.Reset(); return false; }   // o'lik tugma, strelka, F1… — so'z uzildi
            char c = ub[0];

            if (c == '\r' || c == '\t' || c == ' ') { } // so'z chegarasi — dvijokda tekshiriladi
            else if (c < 0x20 || c == 0x7F) { engine.Reset(); return false; }

            if (altGr) { engine.Literal(c); return false; }

            var r = engine.OnChar(c);
            if (!r.Yutish) return false;
            if (!Send(r.Ochirish, r.Yozish)) { engine.Reset(); return false; }
            return true;
        }

        static bool Send(int backspaces, string text)
        {
            var list = new List<W.INPUT>();
            for (int i = 0; i < backspaces; i++) { list.Add(W.Key(0x08, false, MARK)); list.Add(W.Key(0x08, true, MARK)); }
            foreach (char ch in text)
            {
                if (ch == '\r') { list.Add(W.Key(0x0D, false, MARK)); list.Add(W.Key(0x0D, true, MARK)); }
                else if (ch == '\t') { list.Add(W.Key(0x09, false, MARK)); list.Add(W.Key(0x09, true, MARK)); }
                else { list.Add(W.Uni(ch, false, MARK)); list.Add(W.Uni(ch, true, MARK)); }
            }
            var arr = list.ToArray();
            return W.SendInput((uint)arr.Length, arr, Marshal.SizeOf(typeof(W.INPUT))) == arr.Length;
        }

        public void Dispose()
        {
            if (evFg != IntPtr.Zero) W.UnhookWinEvent(evFg);
            if (evFocus != IntPtr.Zero) W.UnhookWinEvent(evFocus);
            hk.Dispose();
            tray.Visible = false;
            tray.Dispose();
        }
    }

    sealed class HotkeyWindow : NativeWindow, IDisposable
    {
        readonly Action action;
        bool registered;
        public HotkeyWindow(Action a) { action = a; CreateHandle(new CreateParams()); }
        public bool Register(uint mods, uint vk) { registered = W.RegisterHotKey(Handle, 1, mods, vk); return registered; }
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0312) action();
            base.WndProc(ref m);
        }
        public void Dispose()
        {
            if (registered) W.UnregisterHotKey(Handle, 1);
            DestroyHandle();
        }
    }

    static class W
    {
        public const int WH_KEYBOARD_LL = 13, WH_MOUSE_LL = 14;
        public const int WM_KEYDOWN = 0x100, WM_SYSKEYDOWN = 0x104;
        public const int WM_LBUTTONDOWN = 0x201, WM_RBUTTONDOWN = 0x204, WM_MBUTTONDOWN = 0x207, WM_MOUSEWHEEL = 0x20A;
        public const uint EVENT_SYSTEM_FOREGROUND = 0x0003, EVENT_OBJECT_FOCUS = 0x8005;
        public const uint WINEVENT_OUTOFCONTEXT = 0, WINEVENT_SKIPOWNPROCESS = 2;
        public const uint MOD_ALT = 1, MOD_SHIFT = 4, MOD_NOREPEAT = 0x4000;
        public const int GWL_STYLE = -16, ES_PASSWORD = 0x20, STATE_SYSTEM_PROTECTED = 0x20000000;

        public delegate IntPtr LowLevelProc(int nCode, IntPtr wParam, IntPtr lParam);
        public delegate void WinEventProc(IntPtr hook, uint ev, IntPtr hwnd, int idObject, int idChild, uint thread, uint time);

        [StructLayout(LayoutKind.Sequential)]
        public struct KBDLLHOOKSTRUCT { public uint vkCode, scanCode, flags, time; public UIntPtr dwExtraInfo; }

        [StructLayout(LayoutKind.Sequential)]
        public struct MOUSEINPUT { public int dx, dy; public uint mouseData, dwFlags, time; public UIntPtr dwExtraInfo; }
        [StructLayout(LayoutKind.Sequential)]
        public struct KEYBDINPUT { public ushort wVk, wScan; public uint dwFlags, time; public UIntPtr dwExtraInfo; }
        [StructLayout(LayoutKind.Explicit)]
        public struct InputUnion { [FieldOffset(0)] public MOUSEINPUT mi; [FieldOffset(0)] public KEYBDINPUT ki; }
        [StructLayout(LayoutKind.Sequential)]
        public struct INPUT { public uint type; public InputUnion u; }

        public static INPUT Key(ushort vk, bool up, UIntPtr mark)
        {
            var i = new INPUT { type = 1 };
            i.u.ki = new KEYBDINPUT { wVk = vk, dwFlags = up ? 2u : 0u, dwExtraInfo = mark };
            return i;
        }
        public static INPUT Uni(char ch, bool up, UIntPtr mark)
        {
            var i = new INPUT { type = 1 };
            i.u.ki = new KEYBDINPUT { wScan = ch, dwFlags = 4u | (up ? 2u : 0u), dwExtraInfo = mark };
            return i;
        }

        [DllImport("user32.dll", SetLastError = true)] public static extern IntPtr SetWindowsHookEx(int id, LowLevelProc fn, IntPtr hMod, uint tid);
        [DllImport("user32.dll")] public static extern bool UnhookWindowsHookEx(IntPtr h);
        [DllImport("user32.dll")] public static extern IntPtr CallNextHookEx(IntPtr h, int n, IntPtr w, IntPtr l);
        [DllImport("user32.dll")] public static extern IntPtr SetWinEventHook(uint min, uint max, IntPtr hmod, WinEventProc fn, uint pid, uint tid, uint flags);
        [DllImport("user32.dll")] public static extern bool UnhookWinEvent(IntPtr h);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr GetModuleHandle(string name);
        [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
        [DllImport("user32.dll")] public static extern IntPtr GetKeyboardLayout(uint tid);
        [DllImport("user32.dll")] public static extern short GetAsyncKeyState(int vk);
        [DllImport("user32.dll")] public static extern short GetKeyState(int vk);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int ToUnicodeEx(uint vk, uint scan, byte[] state, StringBuilder buf, int cch, uint flags, IntPtr hkl);
        [DllImport("user32.dll", SetLastError = true)] public static extern uint SendInput(uint n, INPUT[] inputs, int size);
        [DllImport("user32.dll")] public static extern bool RegisterHotKey(IntPtr hwnd, int id, uint mods, uint vk);
        [DllImport("user32.dll")] public static extern bool UnregisterHotKey(IntPtr hwnd, int id);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr hwnd, StringBuilder sb, int max);
        [DllImport("user32.dll")] public static extern int GetWindowLong(IntPtr hwnd, int idx);
        [DllImport("oleacc.dll")]
        public static extern int AccessibleObjectFromEvent(IntPtr hwnd, int idObject, int idChild,
            [MarshalAs(UnmanagedType.Interface)] out Accessibility.IAccessible acc, out object child);

        // ---- jarayon yaxlitlik darajasi (integrity level)
        [DllImport("kernel32.dll")] static extern IntPtr OpenProcess(uint access, bool inherit, uint pid);
        [DllImport("kernel32.dll")] static extern IntPtr GetCurrentProcess();
        [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);
        [DllImport("advapi32.dll")] static extern bool OpenProcessToken(IntPtr p, uint access, out IntPtr tok);
        [DllImport("advapi32.dll")] static extern bool GetTokenInformation(IntPtr tok, int cls, IntPtr buf, int len, out int ret);
        [DllImport("advapi32.dll")] static extern IntPtr GetSidSubAuthorityCount(IntPtr sid);
        [DllImport("advapi32.dll")] static extern IntPtr GetSidSubAuthority(IntPtr sid, uint idx);

        static int Integrity(IntPtr proc)
        {
            IntPtr tok;
            if (!OpenProcessToken(proc, 8 /*TOKEN_QUERY*/, out tok)) return -1;
            try
            {
                int len;
                GetTokenInformation(tok, 25 /*TokenIntegrityLevel*/, IntPtr.Zero, 0, out len);
                if (len <= 0) return -1;
                IntPtr buf = Marshal.AllocHGlobal(len);
                try
                {
                    if (!GetTokenInformation(tok, 25, buf, len, out len)) return -1;
                    IntPtr sid = Marshal.ReadIntPtr(buf);
                    int cnt = Marshal.ReadByte(GetSidSubAuthorityCount(sid));
                    return Marshal.ReadInt32(GetSidSubAuthority(sid, (uint)(cnt - 1)));
                }
                finally { Marshal.FreeHGlobal(buf); }
            }
            finally { CloseHandle(tok); }
        }

        static int myIntegrity = int.MinValue;

        public static bool SameOrLowerIntegrity(uint pid)
        {
            if (myIntegrity == int.MinValue) myIntegrity = Integrity(GetCurrentProcess());
            IntPtr h = OpenProcess(0x1000 /*QUERY_LIMITED_INFORMATION*/, false, pid);
            if (h == IntPtr.Zero) return false;
            try
            {
                int lvl = Integrity(h);
                return lvl >= 0 && myIntegrity >= 0 && lvl <= myIntegrity;
            }
            finally { CloseHandle(h); }
        }
    }
}
