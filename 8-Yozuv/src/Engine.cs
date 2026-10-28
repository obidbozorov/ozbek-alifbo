// Özbekça Yozuv — yozish paytida o'girish mantig'i (Windows'ga bog'liq emas, Linux'da ham sinaladi).
// Qoidalar Word plaginidagi "Yozişda" rejimi bilan bir xil:
//   s/c + h  -> ş/ç        ş/ç + h  -> sh/ch (qaytarish)
//   o/g + '  -> ö/ğ        ö/ğ + '  -> o'/g' (qaytarish)
// So'z tugaganda (bo'sh joy, tinish belgisi) istisno so'zlar (Photoshop, Chevrolet, ...) va
// chet so'zlar (w yoki h siz c bor) asl holiga qaytariladi.
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Ozbekca
{
    public struct Natija
    {
        public bool Yutish;        // asl tugmani o'tkazmaslik
        public int Ochirish;       // nechta Backspace yuborish
        public string Yozish;      // keyin yoziladigan matn
        public static readonly Natija Otkaz = new Natija();
    }

    public class Engine
    {
        public const string APOS = "'‘’ʻʼ`´′ʹʽʾʿˈˊՙ՚׳῾‛ꞌ";

        class Seg
        {
            public string Out, Orig;
            public Seg(string o, string g) { Out = o; Orig = g; }
        }

        readonly List<Seg> segs = new List<Seg>();
        List<string> istisnolar = new List<string>(DEFAULT_EXCEPTIONS);

        public static readonly string[] DEFAULT_EXCEPTIONS = {
            "ishoq*", "mushaf*", "ashob*", "ishol*", "SHA#", "SSH#", "SSHv2", "OpenSSH", "SHAKE#", "SHACAL#",
            "ChaCha#", "CHACHA#", "XChaCha#", "hash", "hashes", "hashing", "hashed", "hashcash", "hashdump",
            "HashiCorp", "shellcode", "Shellshock", "phishing", "smishing", "vishing", "spearphishing",
            "Schnorr", "Shamir", "Shannon", "Schneier", "Chaum", "Shoup", "Shacham", "Bleichenbacher",
            "Threefish", "Shodan", "Shamoon", "SharpHound", "Dashlane", "Chainlink", "Chainalysis", "bash",
            "zsh", "shell", "powershell", "push", "flash", "fetch", "match", "search", "research", "branch",
            "approach", "chain", "change", "channel", "machine", "dashboard", "shift", "mesh", "refresh",
            "publish", "flush", "splash", "hashtag", "threshold", "Scheme", "Apache", "PyTorch", "Torch",
            "chrome", "chromium", "chromebook", "chatgpt", "sharepoint", "photoshop", "shopify", "Shotcut",
            "Chevrolet", "Chery", "Changan", "Bosch", "Porsche", "english", "Shakespeare", "Michigan",
            "Massachusetts", "Munich", "Zurich", "Utrecht", "O'Brien", "O'Connor", "O'Connell", "O'Neill",
            "O'Neil", "O'Reilly", "O'Hara", "O'Donnell", "O'Sullivan", "O'Keefe", "O'Leary", "O'Malley",
            "O'Toole", "O'Higgins", "o'clock", "g'day",
            // yozişda ko'p uchraydigan inglizcha so'zlar
            "she", "shop", "shopping", "short", "show", "shower", "share", "shared", "sharing", "ship",
            "shipping", "shoes", "shot", "should", "fish", "wish", "english", "cash", "crash", "fresh",
            "finish", "finished", "much", "such", "each", "which", "rich", "teach", "teacher", "watch",
            "check", "checked", "checking", "chef", "chip", "chips", "choice", "choose",
            "school", "chemistry", "chart", "charts", "cheap", "check-in", "checklist", "touch", "lunch",
            "tech", "techno", "technology", "speech", "beach", "bench", "punch", "rush", "brush", "smash",
            "Telegram", "WhatsApp", "Shazam", "Chanel", "Cheshire", "Michael", "Richard", "Charles", "Charlie",
            "Chicago", "China", "Chile", "Czech", "Microsoft", "TikTok", "YouTube", "Twitch", "Huawei"
        };

        // ---------------- istisnolar ----------------
        // istisnolar.txt (Word plagini bilan umumiy): oddiy qator — so'z, "!so'z" — standartni o'chirish, ";" — izoh
        public void IstisnolarniYukla(string path)
        {
            var list = new List<string>(DEFAULT_EXCEPTIONS);
            try
            {
                if (path != null && File.Exists(path))
                {
                    foreach (var raw in File.ReadAllLines(path, Encoding.UTF8))
                    {
                        var s = raw.Trim().TrimStart('﻿');
                        if (s.Length == 0 || s[0] == ';') continue;
                        if (s[0] == '!')
                        {
                            var w = s.Substring(1).Trim().ToLowerInvariant();
                            list.RemoveAll(e => e.ToLowerInvariant() == w);
                        }
                        else list.Add(s);
                    }
                }
            }
            catch { }
            istisnolar = list;
        }

        static bool MatchException(string word, string entry)
        {
            if (string.IsNullOrEmpty(entry)) return false;
            bool cs = entry.ToLowerInvariant() != entry;
            string a = cs ? word : word.ToLowerInvariant();
            string e = cs ? entry : entry.ToLowerInvariant();
            a = NormApos(a); e = NormApos(e);
            char last = e[e.Length - 1];
            if (last == '*') return a.StartsWith(e.Substring(0, e.Length - 1), StringComparison.Ordinal);
            if (last == '#')
            {
                var p = e.Substring(0, e.Length - 1);
                if (!a.StartsWith(p, StringComparison.Ordinal)) return false;
                foreach (char ch in a.Substring(p.Length)) if (!(char.IsDigit(ch) || ch == '-' || ch == '/')) return false;
                return true;
            }
            return a == e;
        }

        static string NormApos(string s)
        {
            var sb = new StringBuilder(s.Length);
            foreach (char c in s) sb.Append(IsApos(c) ? '\'' : c);
            return sb.ToString();
        }

        bool IsException(string core)
        {
            foreach (var e in istisnolar) if (MatchException(core, e)) return true;
            return false;
        }

        // chet so'z belgisi: w yoki ortidan h kelmagan c
        static bool LooksForeign(string core)
        {
            for (int i = 0; i < core.Length; i++)
            {
                char c = core[i];
                if (c == 'w' || c == 'W') return true;
                if (c == 'c' || c == 'C')
                {
                    char n = i + 1 < core.Length ? core[i + 1] : '\0';
                    if (n != 'h' && n != 'H') return true;
                }
            }
            return false;
        }

        public static bool IsApos(char c) { return APOS.IndexOf(c) >= 0; }

        static bool IsBoundary(char c)
        {
            if (IsApos(c)) return false;
            if (char.IsLetterOrDigit(c)) return false;
            return true;
        }

        // ---------------- holat ----------------
        public void Reset() { segs.Clear(); }

        public void Backspace()
        {
            if (segs.Count > 0) segs.RemoveAt(segs.Count - 1);
        }

        // Ctrl/Alt bilan kiritilgan belgi: faqat eslab qolinadi, o'girilmaydi
        public void Literal(char c)
        {
            Add(c.ToString(), c.ToString());
        }

        void Add(string o, string g)
        {
            segs.Add(new Seg(o, g));
            if (segs.Count > 80) segs.RemoveRange(0, 40);
        }

        public Natija OnChar(char c)
        {
            Seg prev = segs.Count > 0 ? segs[segs.Count - 1] : null;
            string cs = c.ToString();

            if ((c == 'h' || c == 'H') && prev != null)
            {
                string rep = null;
                bool qaytar = false;
                switch (prev.Out)
                {
                    case "s": rep = "ş"; break;
                    case "S": rep = "Ş"; break;
                    case "c": rep = "ç"; break;
                    case "C": rep = "Ç"; break;
                    case "ş": rep = "s"; qaytar = true; break;
                    case "Ş": rep = "S"; qaytar = true; break;
                    case "ç": rep = "c"; qaytar = true; break;
                    case "Ç": rep = "C"; qaytar = true; break;
                }
                if (rep != null)
                {
                    segs.RemoveAt(segs.Count - 1);
                    if (qaytar) { Add(rep, rep); Add(cs, cs); return Javob(1, rep + cs); }
                    Add(rep, prev.Orig + cs);
                    return Javob(1, rep);
                }
            }
            else if (IsApos(c) && prev != null)
            {
                string rep = null;
                bool qaytar = false;
                switch (prev.Out)
                {
                    case "o": rep = "ö"; break;
                    case "O": rep = "Ö"; break;
                    case "g": rep = "ğ"; break;
                    case "G": rep = "Ğ"; break;
                    case "ö": rep = "o"; qaytar = true; break;
                    case "Ö": rep = "O"; qaytar = true; break;
                    case "ğ": rep = "g"; qaytar = true; break;
                    case "Ğ": rep = "G"; qaytar = true; break;
                }
                // so'z boshidagi "o" (masalan qo'shtirnoq ichida 'o') emas — oldida harf bo'lishi shart emas,
                // Word plagini ham shunday ishlaydi.
                if (rep != null)
                {
                    segs.RemoveAt(segs.Count - 1);
                    if (qaytar) { Add(rep, rep); Add(cs, cs); return Javob(1, rep + cs); }
                    Add(rep, prev.Orig + cs);
                    return Javob(1, rep);
                }
            }
            else if (IsBoundary(c))
            {
                var n = WordEndFix(c);
                Add(cs, cs);
                return n;
            }

            Add(cs, cs);
            return Natija.Otkaz;
        }

        static Natija Javob(int bs, string text)
        {
            return new Natija { Yutish = true, Ochirish = bs, Yozish = text };
        }

        // So'z tugadi: o'girilgan, lekin istisno yoki chet so'z bo'lsa — asl holiga qaytaramiz
        Natija WordEndFix(char terminator)
        {
            int start = segs.Count;
            while (start > 0 && !(segs[start - 1].Out.Length == 1 && IsBoundary(segs[start - 1].Out[0]))) start--;
            if (start == segs.Count) return Natija.Otkaz;

            var outSb = new StringBuilder();
            var origSb = new StringBuilder();
            for (int i = start; i < segs.Count; i++) { outSb.Append(segs[i].Out); origSb.Append(segs[i].Orig); }
            string o = outSb.ToString(), g = origSb.ToString();
            if (o == g) return Natija.Otkaz;

            // chetdagi apostroflarni (qo'shtirnoqlarni) olib tashlab tekshiramiz
            string core = g.Trim(APOS.ToCharArray());
            if (core.Length == 0) return Natija.Otkaz;
            if (!(IsException(core) || LooksForeign(core))) return Natija.Otkaz;

            segs.RemoveRange(start, segs.Count - start);
            foreach (char ch in g) Add(ch.ToString(), ch.ToString());
            return new Natija { Yutish = true, Ochirish = o.Length, Yozish = g + terminator };
        }
    }
}
