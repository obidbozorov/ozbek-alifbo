using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OzbekAlifboAddIn
{
    /// <summary>
    /// Tezkor tugmalar: Alt + harf. Word oqimiga WH_GETMESSAGE hook o'rnatiladi;
    /// mos tugma bosilganda xabar Wordga yetib bormaydi (WM_NULL), amal esa UI oqimida bajariladi.
    /// </summary>
    internal sealed class Hotkeys : IDisposable
    {
        delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")] static extern IntPtr SetWindowsHookEx(int idHook, HookProc fn, IntPtr hMod, uint threadId);
        [DllImport("user32.dll")] static extern bool UnhookWindowsHookEx(IntPtr hhk);
        [DllImport("user32.dll")] static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")] static extern short GetKeyState(int vk);
        [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();

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

        const int WH_GETMESSAGE = 3, HC_ACTION = 0, PM_REMOVE = 1;
        const uint WM_KEYDOWN = 0x0100, WM_SYSKEYDOWN = 0x0104, WM_NULL = 0x0000;
        const int VK_SHIFT = 0x10, VK_CONTROL = 0x11, VK_MENU = 0x12;

        readonly Dictionary<Keys, Action> _actions = new Dictionary<Keys, Action>();
        readonly Control _marshal = new Control();
        HookProc _proc;
        IntPtr _hook = IntPtr.Zero;

        public Hotkeys()
        {
            _marshal.CreateControl();
            var _ = _marshal.Handle;
        }

        /// <summary>Alt + key uchun amal.</summary>
        public void Register(Keys key, Action action) => _actions[key] = action;

        public void Enable()
        {
            if (_hook != IntPtr.Zero) return;
            _proc = Callback;
            _hook = SetWindowsHookEx(WH_GETMESSAGE, _proc, IntPtr.Zero, GetCurrentThreadId());
        }

        public void Disable()
        {
            if (_hook == IntPtr.Zero) return;
            UnhookWindowsHookEx(_hook);
            _hook = IntPtr.Zero;
        }

        static bool Down(int vk) => GetKeyState(vk) < 0;

        IntPtr Callback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode == HC_ACTION && wParam.ToInt64() == PM_REMOVE)
            {
                var msg = (MSG)Marshal.PtrToStructure(lParam, typeof(MSG));
                if ((msg.message == WM_KEYDOWN || msg.message == WM_SYSKEYDOWN)
                    && Down(VK_MENU) && !Down(VK_CONTROL) && !Down(VK_SHIFT)
                    && _actions.TryGetValue((Keys)(int)msg.wParam.ToInt64(), out var action))
                {
                    msg.message = WM_NULL;                       // Word bu tugmani ko'rmaydi (lenta yorliqlari ham ochilmaydi)
                    Marshal.StructureToPtr(msg, lParam, false);
                    _marshal.BeginInvoke(action);
                }
            }
            return CallNextHookEx(_hook, nCode, wParam, lParam);
        }

        public void Dispose()
        {
            Disable();
            _marshal.Dispose();
        }
    }
}
