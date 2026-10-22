using Microsoft.Win32;

namespace OzbekAlifboAddIn
{
    /// <summary>Sozlamalar HKCU\Software\OzbekAlifbo da saqlanadi.</summary>
    internal static class Settings
    {
        const string Key = @"Software\OzbekAlifbo";

        public static bool Tutuq
        {
            get => Get("Tutuq", false);
            set => Set("Tutuq", value);
        }

        public static bool SkipEnglish
        {
            get => Get("SkipEnglish", true);
            set => Set("SkipEnglish", value);
        }

        public static bool SkipRussian
        {
            get => Get("SkipRussian", true);
            set => Set("SkipRussian", value);
        }

        public static bool Hotkeys
        {
            get => Get("Hotkeys", true);
            set => Set("Hotkeys", value);
        }

        public static bool LiveTyping
        {
            get => Get("LiveTyping", false);
            set => Set("LiveTyping", value);
        }

        public static ConverterOptions Options()
        {
            var o = ConverterOptions.Load();
            o.Tutuq = Tutuq;
            o.SkipEnglish = SkipEnglish;
            return o;
        }

        static bool Get(string name, bool def)
        {
            using (var k = Registry.CurrentUser.OpenSubKey(Key))
            {
                var v = k?.GetValue(name);
                return v is int i ? i != 0 : def;
            }
        }

        static void Set(string name, bool val)
        {
            using (var k = Registry.CurrentUser.CreateSubKey(Key))
                k.SetValue(name, val ? 1 : 0, RegistryValueKind.DWord);
        }
    }
}
