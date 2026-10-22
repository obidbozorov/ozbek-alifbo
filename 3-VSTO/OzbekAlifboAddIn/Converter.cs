// Yangi o'zbek alifbosi konvertori — converter.js bilan bir xil algoritm.
// Oʻ→Ö, Gʻ→Ğ, Sh→Ş, Ch→Ç (Ng o'zgarmaydi)
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OzbekAlifboAddIn
{
    public sealed class Edit
    {
        public int Pos;      // matndagi 0 dan boshlangan o'rin
        public int Len;
        public string From;
        public string To;
    }

    public sealed class ConverterOptions
    {
        public bool Tutuq = false;
        public bool SkipEnglish = true;
        public List<string> Exceptions = new List<string>(Converter.DefaultExceptions);
        public List<string> CyrExceptions = new List<string>(Converter.DefaultCyrExceptions);
        public List<string> LatCyr = null;     // null — standart lug'at

        /// <summary>Standart ro'yxat + istisnolar.txt (VBA bilan umumiy fayl).</summary>
        public static ConverterOptions Load()
        {
            return new ConverterOptions
            {
                Exceptions = Converter.MergeExceptions(ExceptionStore.ReadLines()),
                CyrExceptions = Converter.MergeCyrExceptions(ExceptionStore.ReadLines(Converter.CyrExceptionsPath))
            };
        }
    }

    /// <summary>%APPDATA%\OzbekAlifbo\istisnolar.txt bilan ishlash.</summary>
    public static class ExceptionStore
    {
        public static List<string> ReadLines(string path = null)
        {
            path = path ?? Converter.ExceptionsPath;
            try
            {
                if (File.Exists(path))
                    return File.ReadAllLines(path, Encoding.UTF8).ToList();
            }
            catch { }
            return new List<string>();
        }

        /// <summary>Foydalanuvchi so'zlari va o'chirilgan standartlarni faylga yozadi.</summary>
        public static void Save(IEnumerable<string> userWords, IEnumerable<string> disabledDefaults, string path = null)
        {
            path = path ?? Converter.ExceptionsPath;
            bool cyr = string.Equals(path, Converter.CyrExceptionsPath, StringComparison.OrdinalIgnoreCase);
            var sb = new StringBuilder();
            if (cyr)
            {
                sb.Append("; Yangi alifbo - kirilldan o'girish istisnolari. Har qatorga bitta.\r\n");
                sb.Append("; so'z = o'girilmaydi; so'z* = shu bilan boshlanganlar; o'zak=Lotin = maxsus shakl.\r\n");
                sb.Append("; !o'zak = standart yozuvni o'chirish.\r\n");
            }
            else
            {
                sb.Append("; Yangi alifbo - istisno so'zlar. Har qatorga bitta.\r\n");
                sb.Append("; Katta harf bor bo'lsa - registrga sezgir. * = shu bilan boshlanadi, # = keyin raqam.\r\n");
                sb.Append("; !so'z = standart istisnoni o'chirish.\r\n");
            }
            foreach (var d in disabledDefaults) sb.Append('!').Append(d).Append("\r\n");
            foreach (var w in userWords) sb.Append(w).Append("\r\n");
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            File.WriteAllText(path, sb.ToString(), new UTF8Encoding(false));
        }

        /// <summary>Fayldagi qatorlarni foydalanuvchi so'zlari va o'chirilgan standartlarga ajratadi.</summary>
        public static void Split(IEnumerable<string> lines, List<string> userWords, List<string> disabled)
        {
            foreach (var raw in lines)
            {
                var ln = raw.Trim();
                if (ln.Length == 0 || ln[0] == ';') continue;
                if (ln[0] == '!') { var d = ln.Substring(1).Trim(); if (d.Length > 0 && !disabled.Contains(d, StringComparer.OrdinalIgnoreCase)) disabled.Add(d); }
                else if (ln.Contains('='))
                {
                    // "o'zak=Lotin": bir o'zak uchun bitta yozuv (oxirgisi qoladi)
                    userWords.RemoveAll(w => Converter.ExceptionKey(w) == Converter.ExceptionKey(ln));
                    userWords.Add(ln);
                }
                else if (!userWords.Contains(ln)) userWords.Add(ln);
            }
        }

        /// <summary>Bitta so'zni qo'shadi (takrorlanmasa).</summary>
        public static bool Add(string word)
        {
            word = (word ?? "").Trim();
            if (word.Length == 0) return false;
            var users = new List<string>(); var off = new List<string>();
            Split(ReadLines(), users, off);
            if (users.Contains(word)) return false;
            users.Add(word);
            Save(users, off);
            return true;
        }
    }

    public static class Converter
    {
        public const string Apos = "'‘’ʻʼ`´′ʹʽʾʿˈˊՙ՚׳´῾‛ꞌ";

        // Lotin so'z ichiga adashib kirgan kirill harflari (ko'rinishi bir xil): Tоshkent -> Toshkent
        static readonly Dictionary<char, char> Homo = new Dictionary<char, char>
        {
            ['а'] = 'a', ['е'] = 'e', ['о'] = 'o', ['р'] = 'p', ['с'] = 'c', ['х'] = 'x', ['у'] = 'y', ['к'] = 'k',
            ['А'] = 'A', ['В'] = 'B', ['Е'] = 'E', ['К'] = 'K', ['М'] = 'M', ['Н'] = 'H', ['О'] = 'O', ['Р'] = 'P',
            ['С'] = 'C', ['Т'] = 'T', ['Х'] = 'X',
        };

        public static string FixHomoglyphs(string tok)
        {
            bool hasLat = false, hasHomo = false;
            foreach (char c in tok)
            {
                if ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z')) hasLat = true;
                else if (Homo.ContainsKey(c)) hasHomo = true;
                else if (c >= '\u0400' && c <= '\u04FF') return null;
            }
            if (!hasLat || !hasHomo) return null;
            var sb = new StringBuilder(tok.Length);
            foreach (char c in tok) sb.Append(Homo.TryGetValue(c, out var l) ? l : c);
            return sb.ToString();
        }
        public const char TutuqChar = 'ʼ';

        public static readonly string[] DefaultExceptions =
        {
            "ishoq*", "mushaf*", "ashob*", "ishol*", "SHA#", "SSH#", "SSHv2", "OpenSSH", "SHAKE#",
            "SHACAL#", "ChaCha#", "CHACHA#", "XChaCha#", "hash", "hashes", "hashing", "hashed",
            "hashcash", "hashdump", "HashiCorp", "shellcode", "Shellshock", "phishing", "smishing",
            "vishing", "spearphishing", "Schnorr", "Shamir", "Shannon", "Schneier", "Chaum", "Shoup",
            "Shacham", "Bleichenbacher", "Threefish", "Shodan", "Shamoon", "SharpHound", "Dashlane",
            "Chainlink", "Chainalysis", "bash", "zsh", "shell", "powershell", "push", "flash", "fetch",
            "match", "search", "research", "branch", "approach", "chain", "change", "channel",
            "machine", "dashboard", "shift", "mesh", "refresh", "publish", "flush", "splash",
            "hashtag", "threshold", "Scheme", "Apache", "PyTorch", "Torch", "chrome", "chromium",
            "chromebook", "chatgpt", "sharepoint", "photoshop", "shopify", "Shotcut", "Chevrolet",
            "Chery", "Changan", "Bosch", "Porsche", "english", "Shakespeare", "Michigan",
            "Massachusetts", "Munich", "Zurich", "Utrecht", "O'Brien", "O'Connor", "O'Connell",
            "O'Neill", "O'Neil", "O'Reilly", "O'Hara", "O'Donnell", "O'Sullivan", "O'Keefe", "O'Leary",
            "O'Malley", "O'Toole", "O'Higgins", "o'clock", "g'day"
        };

        /// <summary>Kirilldan o'girish istisnolari: "so'z", "so'z*" — o'girilmaydi; "o'zak=Lotin" — maxsus shakl.</summary>
        public static readonly string[] DefaultCyrExceptions =
        {
            "Майкрософт=Microsoft", "Гугл=Google", "Виндоус=Windows", "Фейсбук=Facebook",
            "Ютуб=YouTube", "Эппл=Apple", "Хуавей=Huawei", "Линукс=Linux", "Айфон=iPhone",
            "Пауэрпойнт=PowerPoint", "Эксель=Excel", "Ворд=Word", "Ватсап=WhatsApp", "Твиттер=Twitter",
            "Скайп=Skype", "абзац=abzas", "авианосец=avianoses", "броненосец=bronenoses",
            "миноносец=minonoses", "марганец=marganes", "ранец=ranes", "коэффициент=koeffitsient",
            "целлюлоид=selluloid", "целлюлоза=selluloza", "октябр=oktabr", "сентябр=sentabr",
            "партьер=parter", "тюлен=tulen"
        };

        static readonly HashSet<string> EnWords = new HashSet<string>
        {
            "the", "and", "of", "are", "for", "with", "this", "that", "from", "which", "was",
            "were", "have", "has", "been", "these", "their", "into", "its"
        };

        static readonly HashSet<string> MonoFonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Courier New", "Courier", "Consolas", "Lucida Console", "Cascadia Code", "Cascadia Mono",
            "Source Code Pro", "Menlo", "Monaco", "JetBrains Mono", "Fira Code"
        };

        static readonly string[] FileExts =
        {
            "com", "uz", "org", "net", "ru", "io", "edu", "gov", "doc", "docx", "pdf", "xls", "xlsx",
            "ppt", "pptx", "exe", "py", "js", "htm", "html", "txt", "zip"
        };

        public static string ExceptionsPath =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                         "OzbekAlifbo", "istisnolar.txt");

        public static string CyrExceptionsPath =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                         "OzbekAlifbo", "kirill-istisnolar.txt");

        /// <summary>Istisno yozuvining kaliti ("o'zak=Lotin" dan o'zak qismi, kichik harf).</summary>
        public static string ExceptionKey(string e) => (e ?? "").Split('=')[0].Trim().ToLowerInvariant();

        public static bool IsApos(char c) => Apos.IndexOf(c) >= 0;
        public static bool IsLetter(char c) => char.ToLowerInvariant(c) != char.ToUpperInvariant(c);
        static bool IsDigit(char c) => c >= '0' && c <= '9';
        static bool IsAlnum(char c) => IsLetter(c) || IsDigit(c);
        static bool IsSpace(char c) => c <= ' ' || char.IsWhiteSpace(c);

        public static bool IsMonoFont(string name) => !string.IsNullOrEmpty(name) && MonoFonts.Contains(name);

        public static string Convert(string text, ConverterOptions opts)
        {
            var edits = ComputeEdits(text, opts);
            var sb = new StringBuilder(text);
            for (int k = edits.Count - 1; k >= 0; k--)
            {
                sb.Remove(edits[k].Pos, edits[k].Len);
                sb.Insert(edits[k].Pos, edits[k].To);
            }
            return sb.ToString();
        }

        public static List<Edit> ComputeEdits(string text, ConverterOptions opts)
        {
            var edits = new List<Edit>();
            if (string.IsNullOrEmpty(text)) return edits;
            if (opts.SkipEnglish && IsEnglishText(text)) return edits;

            bool quoteOpen = false;
            int i = 0, n = text.Length;
            while (i < n)
            {
                if (IsSpace(text[i])) { i++; continue; }
                int start = i;
                while (i < n && !IsSpace(text[i])) i++;
                quoteOpen = ProcessToken(text.Substring(start, i - start), start, quoteOpen, edits, opts);
            }
            return edits;
        }

        static bool ProcessToken(string orig, int baseIdx, bool quoteOpen, List<Edit> allEdits, ConverterOptions opts)
        {
            string fixedTok = FixHomoglyphs(orig);
            string tok = fixedTok ?? orig;
            var edits = new List<Edit>();
            int len = tok.Length;
            int lead = 0;
            while (lead < len && !IsAlnum(tok[lead]))
            {
                if (IsApos(tok[lead])) quoteOpen = true;
                lead++;
            }
            if (lead >= len) { AddHomo(orig, fixedTok, baseIdx, edits, allEdits); return quoteOpen; }

            int limit = len;
            int t = len;
            while (t > lead && !IsLetter(tok[t - 1]) && !IsApos(tok[t - 1]) && !IsDigit(tok[t - 1])) t--;
            if (t > lead && IsApos(tok[t - 1]))
            {
                int aposRun = 0;
                while (t - 1 - aposRun > lead && IsApos(tok[t - 1 - aposRun])) aposRun++;
                if (aposRun >= 2 && !quoteOpen) limit = t - aposRun - 1;
                else if (quoteOpen) { limit = t - 1; quoteOpen = false; }
            }

            string core = CoreOf(tok);
            bool skip = LooksLikeLinkOrCode(tok) || IsExceptionWord(core, opts.Exceptions) || LooksForeign(core);

            int i = skip ? limit : lead;
            while (i < limit)
            {
                char c = tok[i];
                char nx = i + 1 < limit ? tok[i + 1] : '\0';

                if ((c == 'o' || c == 'O' || c == 'g' || c == 'G') && nx != '\0' && IsApos(nx))
                {
                    char nx2 = i + 2 < limit ? tok[i + 2] : '\0';
                    if (nx2 != '\0' && IsApos(nx2)) { i += 3; continue; }
                    string rep = c == 'o' ? "ö" : c == 'O' ? "Ö" : c == 'g' ? "ğ" : "Ğ";
                    edits.Add(new Edit { Pos = baseIdx + i, Len = 2, From = orig.Substring(i, 2), To = rep });
                    i += 2; continue;
                }
                if ((c == 's' || c == 'S' || c == 'c' || c == 'C') && (nx == 'h' || nx == 'H'))
                {
                    string rep = c == 's' ? "ş" : c == 'S' ? "Ş" : c == 'c' ? "ç" : "Ç";
                    edits.Add(new Edit { Pos = baseIdx + i, Len = 2, From = orig.Substring(i, 2), To = rep });
                    i += 2; continue;
                }
                if (opts.Tutuq && IsApos(c) && c != TutuqChar && i > lead && IsLetter(tok[i - 1])
                    && nx != '\0' && IsLetter(nx))
                {
                    edits.Add(new Edit { Pos = baseIdx + i, Len = 1, From = c.ToString(), To = TutuqChar.ToString() });
                    i += 1; continue;
                }
                i++;
            }
            AddHomo(orig, fixedTok, baseIdx, edits, allEdits);
            return quoteOpen;
        }

        // o'girilmagan joylardagi o'xshash kirill harflari — bittadan almashtiriladi
        static void AddHomo(string orig, string fixedTok, int baseIdx, List<Edit> edits, List<Edit> allEdits)
        {
            if (fixedTok != null)
            {
                for (int h = 0; h < orig.Length; h++)
                {
                    if (orig[h] == fixedTok[h]) continue;
                    int p = baseIdx + h;
                    if (edits.Any(e => p >= e.Pos && p < e.Pos + e.Len)) continue;
                    edits.Add(new Edit { Pos = p, Len = 1, From = orig[h].ToString(), To = fixedTok[h].ToString() });
                }
                edits.Sort((a, b) => a.Pos.CompareTo(b.Pos));
            }
            allEdits.AddRange(edits);
        }

        static string CoreOf(string tok)
        {
            int s = 0, e = tok.Length;
            while (s < e && !IsAlnum(tok[s])) s++;
            while (e > s && !IsAlnum(tok[e - 1])) e--;
            return tok.Substring(s, e - s);
        }

        static bool LooksForeign(string core)
        {
            for (int i = 0; i < core.Length; i++)
            {
                char c = core[i];
                if (c == 'w' || c == 'W') return true;
                if (c == 'c' || c == 'C')
                {
                    char nx = i + 1 < core.Length ? core[i + 1] : '\0';
                    if (nx != 'h' && nx != 'H') return true;
                }
            }
            return false;
        }

        static bool LooksLikeLinkOrCode(string tok)
        {
            string lt = tok.ToLowerInvariant();
            if (lt.Contains("://") || lt.StartsWith("www.") || lt.Contains("@") || lt.StartsWith("#") || lt.Contains("\\"))
                return true;
            int d = lt.LastIndexOf('.');
            if (d > 0 && d < lt.Length - 1 && FileExts.Contains(lt.Substring(d + 1)))
                return lt.All(ch => (ch >= 'a' && ch <= 'z') || IsDigit(ch) || ch == '.' || ch == '-' || ch == '_');
            return false;
        }

        public static bool IsEnglishText(string text)
        {
            var words = new List<string>();
            var sb = new StringBuilder();
            foreach (char ch in text.ToLowerInvariant() + " ")
            {
                if (IsLetter(ch) || ch == '\'') sb.Append(ch);
                else if (sb.Length > 0) { words.Add(sb.ToString()); sb.Clear(); }
            }
            if (words.Count < 3) return false;
            int hits = words.Count(w => EnWords.Contains(w));
            return hits >= 2 && (double)hits / words.Count >= 0.08;
        }

        static bool IsExceptionWord(string core, List<string> exceptions) => FindException(core, exceptions) != null;

        /// <summary>So'zni himoya qilayotgan istisno yozuvi (yoki null).</summary>
        public static string FindException(string core, List<string> exceptions)
        {
            if (string.IsNullOrEmpty(core)) return null;
            var parts = new List<string> { core };
            parts.AddRange(core.Split('-', '/'));
            foreach (var p in parts)
                foreach (var e in exceptions)
                    if (MatchException(p, e)) return e;
            return null;
        }

        /// <summary>Standart ro'yxat + foydalanuvchi qatorlari ("!so'z" standartni o'chiradi,
        /// foydalanuvchining shu kalitli yozuvi standartni almashtiradi).</summary>
        public static List<string> MergeExceptions(IEnumerable<string> userLines, string[] defaults = null)
        {
            defaults = defaults ?? DefaultExceptions;
            var users = new List<string>(); var off = new List<string>();
            ExceptionStore.Split(userLines, users, off);
            var offKeys = new HashSet<string>(off.Select(ExceptionKey));
            var userKeys = new HashSet<string>(users.Select(ExceptionKey));
            var list = defaults.Where(d => !offKeys.Contains(ExceptionKey(d)) && !userKeys.Contains(ExceptionKey(d))).ToList();
            list.AddRange(users);
            return list;
        }

        public static List<string> MergeCyrExceptions(IEnumerable<string> userLines) => MergeExceptions(userLines, DefaultCyrExceptions);

        /// <summary>So'z nima bo'lishi va nega: (natija, sabab yoki null).</summary>
        public static Tuple<string, string> Explain(string word, ConverterOptions opts)
        {
            var tok = (word ?? "").Trim();
            var core = CoreOf(tok);
            var o = new ConverterOptions { Tutuq = opts.Tutuq, SkipEnglish = false, Exceptions = opts.Exceptions };
            string reason = null;
            if (LooksLikeLinkOrCode(tok)) reason = "havola yoki fayl nomi";
            else if (FindException(core, opts.Exceptions) != null) reason = "istisno: " + FindException(core, opts.Exceptions);
            else if (LooksForeign(core)) reason = "chet so'z (w yoki yolg'iz c harfi bor)";
            return Tuple.Create(Convert(tok, o), reason);
        }

        public static string WordCore(string tok) => CoreOf(tok ?? "");

        /// <summary>O'girilgan so'zni eski yozuvga qaytaradi: Şamir → Shamir, ŞAHAR → SHAHAR.</summary>
        public static string RevertWord(string w)
        {
            w = w ?? "";
            var letters = new string(w.Where(IsLetter).ToArray());
            bool allCaps = letters.Length > 1 && letters == letters.ToUpperInvariant();
            var sb = new StringBuilder();
            foreach (char c in w)
            {
                switch (c)
                {
                    case '\u015F': sb.Append("sh"); break;
                    case '\u00E7': sb.Append("ch"); break;
                    case '\u00F6': sb.Append("o\u02BB"); break;
                    case '\u011F': sb.Append("g\u02BB"); break;
                    case '\u015E': sb.Append(allCaps ? "SH" : "Sh"); break;
                    case '\u00C7': sb.Append(allCaps ? "CH" : "Ch"); break;
                    case '\u00D6': sb.Append("O\u02BB"); break;
                    case '\u011E': sb.Append("G\u02BB"); break;
                    default: sb.Append(c); break;
                }
            }
            return sb.ToString();
        }

        static bool MatchException(string word, string entry)
        {
            if (string.IsNullOrEmpty(entry) || string.IsNullOrEmpty(word)) return false;
            bool cs = entry.ToLowerInvariant() != entry;
            string a = cs ? word : word.ToLowerInvariant();
            string e = cs ? entry : entry.ToLowerInvariant();
            char last = e[e.Length - 1];
            if (last == '*') return a.StartsWith(e.Substring(0, e.Length - 1), StringComparison.Ordinal);
            if (last == '#')
            {
                string p = e.Substring(0, e.Length - 1);
                if (!a.StartsWith(p, StringComparison.Ordinal)) return false;
                return a.Substring(p.Length).All(ch => IsDigit(ch) || ch == '-' || ch == '/');
            }
            return a == e;
        }

        // ======================================================== KIRILL → YANGI LOTIN
        static readonly Dictionary<char, string> CyrMap = new Dictionary<char, string>
        {
            {'а',"a"},{'б',"b"},{'в',"v"},{'г',"g"},{'д',"d"},{'ж',"j"},{'з',"z"},{'и',"i"},{'й',"y"},
            {'к',"k"},{'л',"l"},{'м',"m"},{'н',"n"},{'о',"o"},{'п',"p"},{'р',"r"},{'с',"s"},{'т',"t"},
            {'у',"u"},{'ф',"f"},{'х',"x"},{'ч',"ç"},{'ш',"ş"},{'щ',"ş"},{'ы',"i"},{'э',"e"},
            {'ё',"yo"},{'ю',"yu"},{'я',"ya"},{'ў',"ö"},{'қ',"q"},{'ғ',"ğ"},{'ҳ',"h"}
        };
        const string CyrVowels = "аеёиоуэюяўы";
        static readonly HashSet<string> RuWords = new HashSet<string>
        {
            "в","на","не","что","для","это","как","при","или","по","же","был","была","были",
            "также","который","которые","которая","его","с","и","к"
        };

        static bool IsCyr(char c) => c >= 'Ѐ' && c <= 'ӿ';

        public static bool IsRussianText(string text)
        {
            if (text.IndexOfAny("ўқғҳЎҚҒҲ".ToCharArray()) >= 0) return false;
            var words = new List<string>();
            var sb = new StringBuilder();
            foreach (char ch in text.ToLowerInvariant() + " ")
            {
                if (IsCyr(ch)) sb.Append(ch);
                else if (sb.Length > 0) { words.Add(sb.ToString()); sb.Clear(); }
            }
            if (words.Count < 3) return false;
            int hits = words.Count(w => RuWords.Contains(w));
            return hits >= 2 && (double)hits / words.Count >= 0.08;
        }

        static string CyrChar(string text, int i, int runStart)
        {
            char lc = char.ToLowerInvariant(text[i]);
            char prev = i > runStart ? char.ToLowerInvariant(text[i - 1]) : '\0';
            char next = i + 1 < text.Length ? char.ToLowerInvariant(text[i + 1]) : '\0';
            switch (lc)
            {
                case 'е': return (i == runStart || (prev != '\0' && CyrVowels.IndexOf(prev) >= 0) || prev == 'ъ' || prev == 'ь') ? "ye" : "e";
                case 'ц': return (i == runStart || CyrVowels.IndexOf(prev) < 0) ? "s" : "ts";
                case 'ъ': return next != '\0' && "еёюя".IndexOf(next) >= 0 ? "" : TutuqChar.ToString();
                case 'ь': return next == 'о' ? "y" : "";
                case 'с': return next == 'ҳ' ? "s" + TutuqChar : "s";
            }
            return CyrMap.TryGetValue(lc, out var t) ? t : text[i].ToString();
        }

        sealed class CyrRule { public string Key; public bool Prefix; public string Latin; }

        // kalit -> qoida (keyingi yozuv avvalgisini almashtiradi)
        static Dictionary<string, CyrRule> ParseCyrRules(IEnumerable<string> list)
        {
            var map = new Dictionary<string, CyrRule>();
            foreach (var raw in list ?? DefaultCyrExceptions)
            {
                var e = (raw ?? "").Trim(); int eq = e.IndexOf('=');
                CyrRule r;
                if (eq > 0) r = new CyrRule { Key = e.Substring(0, eq).Trim().ToLowerInvariant(), Prefix = true, Latin = e.Substring(eq + 1).Trim() };
                else if (e.EndsWith("*")) r = new CyrRule { Key = e.Substring(0, e.Length - 1).ToLowerInvariant(), Prefix = true };
                else r = new CyrRule { Key = e.ToLowerInvariant(), Prefix = false };
                if (r.Key.Length > 0) map[r.Key] = r;
            }
            return map;
        }

        // eng uzun mos kalit (so'z boshidan)
        static CyrRule FindCyrRule(string run, Dictionary<string, CyrRule> map)
        {
            string rl = run.ToLowerInvariant();
            for (int len = rl.Length; len > 0; len--)
                if (map.TryGetValue(rl.Substring(0, len), out var r) && (r.Prefix || len == rl.Length)) return r;
            return null;
        }

        /// <summary>Kirill matnidagi o'zgarishlar (har bir kirill so'zi — bitta o'zgarish).</summary>
        public static List<Edit> ComputeCyrillicEdits(string text, bool skipRussian = true, IEnumerable<string> cyrExceptions = null)
        {
            var edits = new List<Edit>();
            if (string.IsNullOrEmpty(text)) return edits;
            if (skipRussian && IsRussianText(text)) return edits;
            var rules = ParseCyrRules(cyrExceptions);
            int n = text.Length, i = 0;
            while (i < n)
            {
                if (!IsCyr(text[i])) { i++; continue; }
                int start = i;
                while (i < n && IsCyr(text[i])) i++;
                string run = text.Substring(start, i - start);
                bool allCaps = run.Length > 1 && run == run.ToUpperInvariant() && run != run.ToLowerInvariant();
                var rule = FindCyrRule(run, rules);
                if (rule != null && rule.Latin == null) continue;          // istisno: o'girilmaydi
                var sb = new StringBuilder();
                int from = start;
                if (rule != null)
                {
                    string lat = rule.Latin;
                    if (allCaps) lat = lat.ToUpperInvariant();
                    else if (char.IsUpper(run[0]) && lat == lat.ToLowerInvariant() && lat.Length > 0)
                        lat = char.ToUpperInvariant(lat[0]) + lat.Substring(1);      // Абзац -> Abzas
                    sb.Append(lat);
                    from = start + rule.Key.Length;
                }
                for (int k = from; k < i; k++)
                {
                    string t = CyrChar(text, k, start);
                    if (t.Length > 0 && char.IsUpper(text[k]))
                        t = allCaps ? t.ToUpperInvariant() : char.ToUpperInvariant(t[0]) + t.Substring(1);
                    sb.Append(t);
                }
                string outS = sb.ToString();
                if (outS != run) edits.Add(new Edit { Pos = start, Len = run.Length, From = run, To = outS });
            }
            return edits;
        }

        public static string CyrToLatin(string text, bool skipRussian = true, IEnumerable<string> cyrExceptions = null)
        {
            var edits = ComputeCyrillicEdits(text, skipRussian, cyrExceptions);
            var sb = new StringBuilder(text);
            for (int k = edits.Count - 1; k >= 0; k--) { sb.Remove(edits[k].Pos, edits[k].Len); sb.Insert(edits[k].Pos, edits[k].To); }
            return sb.ToString();
        }

        // ======================================================== LOTIN → KIRILL
        /// <summary>Lotin → kirill lug'ati: "lotin=кирилл" (eski lotin shaklida). 5+ harfli yoki * bilan tugagan kalit — o'zak.</summary>
        public static readonly string[] DefaultLatCyr =
        {
            "abzas=абзац", "aksent=акцент", "avianoses=авианосец", "batsilla=бацилла",
            "biomitsin=биомицин", "botsman=боцман", "bronenoses=броненосец", "brutselloz=бруцеллоз",
            "cherepitsa=черепица", "dotsent=доцент", "fransuz=француз", "gaubitsa=гаубица",
            "gers=герц", "glitserin=глицерин", "gorchitsa=горчица", "gusenitsa=гусеница",
            "inersiya=инерция", "inssenirovka=инсценировка", "kalsiy=кальций", "kansler=канцлер",
            "koeffitsient=коэффициент", "konferens-zal=конференц-зал", "konsepsiya=концепция",
            "konsern=концерн", "konsert=концерт", "konslager=концлагерь", "kultivatsiya=культивация",
            "kvars=кварц", "litsey=лицей", "lotsman=лоцман", "marganes=марганец",
            "minonoses=миноносец", "munitsipalitet=муниципалитет", "ofitsiant=официант",
            "penitsillin=пенициллин", "platsdarm=плацдарм", "platskarta=плацкарта",
            "politsmeyster=полицмейстер", "ranes=ранец", "ritsar=рицарь", "sellofan=целлофан",
            "selluloid=целлюлоид", "selluloza=целлюлоза", "selsiy=цельсий", "sement=цемент",
            "sentner=центнер", "sentrifuga=центрифуга", "senz=ценз", "senzor=цензор",
            "senzura=цензура", "sex=цех", "shnitsel=шницель", "shveysar=швейцар",
            "siferblat=циферблат", "silindr=цилиндр", "silindrik=цилиндрик", "sink=цинк", "sirk=цирк",
            "sirkulyar=циркуляр", "sisterna=цистерна", "sitrus=цитрус", "sotsiologik=социологик",
            "sotsiologiya=социология", "ssenariy=сценарий", "kultivator=культиватор", "kuryer=курьер",
            "lager=лагерь", "losos=лосось", "mebel=мебель", "medal=медаль", "medalyon=медальон",
            "menshevik=меньшевик", "menshevizm=меньшевизм", "migren=мигрень", "mikrofilm=микрофильм",
            "mil=миль", "model=модель", "neft=нефть", "nikel=никель", "nippel=ниппель", "nol=ноль",
            "noyabr=ноябрь", "oktabr=октябрь", "palto=пальто", "panel=панель", "parallel=параллель",
            "parol=пароль", "parter=партьер", "patrul=патруль", "pavilyon=павильон", "pedal=педаль",
            "plastir=пластирь", "pochtalyon=почтальон", "porshen=поршень", "portfel=портфель",
            "povest=повесть", "predoxranitel=предохранитель", "premyera=премьера", "pristan=пристань",
            "puls=пульс", "pyesa=пьеса", "rels=рельс", "relyef=рельеф", "rentabel=рентабель",
            "rezba=резьба", "rol=роль", "royal=рояль", "rul=руль", "seld=сельд", "sentabr=сентябрь",
            "shinel=шинель", "shpatel=шпатель", "shpilka=шпилька", "shpindel=шпиндель",
            "shtapel=штапель", "shtempel=штемпель", "shtepsel=штепсель", "spektakl=спектакль",
            "spiral=спираль", "sterjen=стержень", "sudya=судья", "sulfat=сульфат", "tabel=табель",
            "tekstil=текстиль", "tokar=токарь", "tulen=тюлень", "tunnel=туннель",
            "umivalnik=умивальник", "valeryanka=валерьянка", "vals=вальс", "veksel=вексель",
            "velvet=вельвет", "ventil=вентиль", "vermishel=вермишель", "vimpel=вимпель",
            "violonchel=виолончель", "volfram=вольфрам", "volt=вольт", "volta=вольта",
            "voltmetr=вольтметр", "vulgar=вульгар", "yakor=якорь", "yanvar=январь",
            "yuriskonsult=юрисконсульт", "ansambl=ансамбль", "artel=артель", "artikl=артикль",
            "aryergard=арьергард", "asfalt=асфальт", "atelye=ателье", "avtomobil=автомобиль",
            "balzam=бальзам", "banderol=бандероль", "batalyon=батальон", "bilyard=бильярд",
            "binokl=бинокль", "bolshevik=большевик", "budilnik=будильник", "bulvar=бульвар",
            "dalton=дальтон", "dekabr=декабрь", "delfin=дельфин", "devalvatsiya=девальвация",
            "dirijabl=дирижабль", "dizel=дизель", "dizel-motor=дизель-мотор", "duel=дуэль",
            "dvigatel=двигатель", "emulsiya=эмульсия", "eskadrilya=эскадрилья",
            "fakultativ=факультатив", "fakultet=факультет", "falsifikator=фальсификатор",
            "falsifikatsiya=фальсификация", "feldmarshal=фельдмаршал", "feldsher=фельдшер",
            "festival=фестиваль", "fevral=февраль", "filtr=фильтр", "folklor=фольклор",
            "fotoalbom=фотоальбом", "fotoatelye=фотоателье", "gantel=гантель", "gastrol=гастроль",
            "gilza=гильза", "gospital=госпиталь", "grifel=грифель", "impuls=импульс", "insult=инсульт",
            "intervyu=интервью", "inyeksiya=инъекция", "italyan=итальян", "iyul=июль", "iyun=июнь",
            "kabel=кабель", "kalendar=календарь", "kanifol=канифоль", "karamel=карамель",
            "kartel=картель", "karusel=карусель", "katapulta=катапульта", "kinofestival=кинофестиваль",
            "kinofilm=кинофильм", "kisel=кисель", "kitel=китель", "kobalt=кобальт",
            "kompanyon=компаньон", "konferansye=конферансье", "syomka=съёмка", "korrupsiya=коррупция",
            "konyunktura=конъюнктура", "adyunktura=адъюнктура", "ishoq=исҳоқ", "mushaf=мусҳаф",
            "ashob=асҳоб", "ishol=исҳол", "mayor=майор", "rayon=район", "mikrorayon=микрорайон",
            "mayonez=майонез", "yogurt=йогурт", "yoga*=йога", "poyabzal=пойабзал", "poyandoz=пойандоз",
            "poyafzal=пойафзал", "poyustun=пойустун", "moyupa=мойупа", "konveyer=конвейер",
            "pleyer=плейер", "stayer=стайер", "foye*=фойе", "kompyuter=компьютер", "bulyon=бульон",
            "aprel=апрель", "film*=фильм", "rubl*=рубль", "prinsip=принцип", "sitata=цитата",
            "sivilizatsiya=цивилизация", "sikl*=цикл", "siklon=циклон", "sentr*=центр",
            "konsentr=концентр", "spetsif=специф", "spetsial=специал", "patsient=пациент",
            "ofitsial=официал", "stansiya=станция", "instansiya=инстанция", "distansiya=дистанция",
            "kvitansiya=квитанция", "tendensiya=тенденция", "potensial=потенциал",
            "litsenziya=лицензия", "absorbsiya=абсорбция", "abssess=абсцесс", "adsorbsiya=адсорбция",
            "adyunkt=адъюнкт", "aerofotosyomka=аэрофотосъёмка", "aerosyomka=аэросъёмка",
            "agnostitsizm=агностицизм", "aksioma=аксиома", "aksiz=акциз", "akvarel=акварель",
            "albatros=альбатрос", "albom=альбом", "alfa-zarralar=альфа-зарралар", "alfa=альфа",
            "alkogol=алкоголь", "alma-terapiya=альма-терапия", "altchi=альтчи", "altimetr=альтиметр",
            "alyans=альянс", "amalgama=амальгама", "anneksiya=аннексия", "apelsin=апельсин",
            "apopleksiya=апоплексия", "asfiksiya=асфиксия", "aviamodel=авиамодель",
            "avtomagistral=автомагистраль", "balneolog=бальнеолог", "balneologik=бальнеологик",
            "balneologiya=бальнеология", "balneoterapiya=бальнеотерапия",
            "balneotexnika=бальнеотехника", "barelyef=барельеф", "barrel=баррель", "baryer=барьер",
            "bazalt=базальт", "beletaj=бельэтаж", "belgiyalik=бельгиялик", "belting=бельтинг",
            "belveder=бельведер", "bemeyor=бемеъёр", "biofiltr=биофильтр",
            "biolyuminessensiya=биолюминесценция", "bolonya=болонья", "bolshevizm=большевизм",
            "brakonyerlik=браконьерлик", "broneavtomobil=бронеавтомобиль", "buldenej=бульденеж",
            "buldog=бульдог", "buldozer=бульдозер", "buldozerchi=бульдозерчи",
            "burgutsifat=бургутсифат", "burgutsimon=бургутсимон", "byulleten=бюллетень",
            "chizel=чизель", "daraxtsimon=дарахтсимон", "delta=дельта", "detal=деталь",
            "diafilm=диафильм", "differensiatsiya=дифференциация", "distansion=дистанцион",
            "drel=дрель", "emal=эмаль", "empiriokrititsizm=эмпириокритицизм", "endshpil=эндшпиль",
            "ensefalit=энцефалит", "ensefalogramma=энцефалограмма", "ensiklopedik=энциклопедик",
            "ensiklopedist=энциклопедист", "ensiklopediya=энциклопедия",
            "ensiklopediyachi=энциклопедиячи", "esminets=эсминец", "feldyeger=фельдъегерь",
            "felyeton=фельетон", "felyetonchi=фельетончи", "filmoy=филмой", "fitonsid=фитонцид",
            "fizkultura=физкультура", "fizkulturachi=физкультурачи", "fleksiya=флексия",
            "folga=фольга", "fonar=фонарь", "fortepyano=фортепьяно", "fotosyomka=фотосъёмка",
            "fransiya=франция", "frits=фриц", "gallitsizm=галлицизм", "gallyutsinatsiya=галлюцинация",
            "galvanik=гальваник", "galvanizatsiya=гальванизация", "galvanometr=гальванометр",
            "gavan=гавань", "gelmint=гельминт", "gelmintologiya=гельминтология",
            "gelmintoz=гельминтоз", "geraldika=геральдика", "gersog=герцог",
            "gidropodyomnik=гидроподъёмник", "giposulfit=гипосульфит", "golf=гольф",
            "gorelyef=горельеф", "guash=гуашь", "instarsiya=инстарция", "interyer=интерьер",
            "inventar=инвентарь", "itsigek=итсигек", "jannatsimon=жаннатсимон", "jenshen=женьшень",
            "kafel=кафель", "kalka=калька", "kalkalamoq=калькаламоқ", "kalkulyatsiya=калькуляция",
            "kanserogen=канцероген", "kapelmeyster=капельмейстер", "kapsyul=капсюль",
            "kartech=картечь", "karyer=карьер", "kastryul=кастрюль", "kastryulka=кастрюлька",
            "kaxeksiya=кахексия", "kinokonsert=киноконцерт", "kinossenariy=киносценарий",
            "klassitsizm=классицизм", "knyaz=князь", "kokteyl=коктейль", "kolonsifra=колонцифра",
            "kommersiya=коммерция", "konki=коньки", "konkichi=конькичи", "konsessiya=концессия",
            "konsorsium=консорциум", "konsultant=консультант", "konsultativ=консультатив",
            "konsultatsiya=консультация", "kontrol=контроль", "konyak=коньяк", "korol=король",
            "krovat=кровать", "kulminatsion=кульминацион", "kulminatsiya=кульминация",
            "kulturizm=культуризм", "kyat=кьят", "lanset=ланцет", "latun=латунь", "losyon=лосьон",
            "lourensiy=лоуренсий", "marselyeza=марсельеза", "matritsa=матрица", "mayub=маъюб",
            "mayus=маъюс", "mayuslanmoq=маъюсланмоқ", "mayuslik=маъюслик", "mayusona=маъюсона",
            "meditsina=медицина", "mehnatsevar=меҳнатсевар", "mehnatsevarlik=меҳнатсеварлик",
            "melxior=мельхиор", "mexanitsizm=механицизм", "meyor=меъёр", "meyoriy=меъёрий",
            "mikroinsult=микроинсульт", "mistitsizm=мистицизм", "mizanssena=мизансцена",
            "modelyer=модельер", "molbert=мольберт", "monastir=монастирь", "motel=мотель",
            "multimediya=мультимедия", "multimillioner=мультимиллионер",
            "multiplikator=мультипликатор", "multiplikatsion=мультипликацион",
            "multiplikatsiya=мультипликация", "nenets=ненец", "nenetslar=ненецлар",
            "nimpalto=нимпальто", "oblast=область", "ochered=очередь", "oftalmologiya=офтальмология",
            "okkultizm=оккультизм", "opsion=опцион", "ordinarets=ординарец", "otel=отель",
            "palma=пальма", "palmazor=пальмазор", "palpatsiya=пальпация", "pechat=печать",
            "pechenye=печенье", "penalti=пенальти", "petlitsa=петлица", "pinset=пинцет",
            "pitssa=пицца", "platskart=плацкарт", "plebissit=плебисцит", "porsillamoq=порсилламоқ",
            "porsilloq=порсиллоқ", "porsin=порсин", "portsigar=портсигар",
            "premyer-ministr=премьер-министр", "press-papye=пресс-папье",
            "press-sekretar=пресс-секретарь", "profil=профиль", "provinsializm=провинциализм",
            "provinsiya=провинция", "pulmonologiya=пульмонология", "pulpit=пульпит", "pult=пульт",
            "pulverizator=пульверизатор", "qaynegachi=қайнэгачи", "qimmatsiramoq=қимматсирамоқ",
            "radiospektakl=радиоспектакль", "rantye=рантье", "raya-andisha=раъя-андиша", "raya=раъя",
            "reket=рэкет", "rekognossirovka=рекогносцировка", "relsli=рельсли", "relssiz=рельссиз",
            "retush=ретушь", "revalvatsiya=ревальвация", "revolver=револьвер", "riyel=риель",
            "rubilnik=рубильник", "salto=сальто", "sanga=цанга", "sapfa=цапфа", "sedra=цедра",
            "sekretar=секретарь", "selderey=сельдерей", "sent=цент", "senyor=сеньор",
            "senyora=сеньора", "seytnot=цейтнот", "seziy=цезий", "shagren=шагрень", "shampun=шампунь",
            "sherst=шерсть", "shifonyer=шифоньер", "shmutstitul=шмуцтитул", "shprits=шприц",
            "shtangensirkul=штангенциркуль", "sian=циан", "sianli=цианли", "sinkalamoq=синькаламоқ",
            "sinkograf=цинкограф", "sinkografiya=цинкография", "siren=сирень", "sirkchi=циркчи",
            "sirkoniy=цирконий", "sirkul=циркуль", "sirroz=цирроз", "sistit=цистит",
            "sitodiagnostika=цитодиагностика", "sitokimyo=цитокимё", "sitoliz=цитолиз",
            "sitologiya=цитология", "skalpel=скальпель", "skeptitsizm=скептицизм", "slanets=сланец",
            "slesar=слесарь", "sobol=соболь", "sokol=цоколь", "ssenarist=сценарист", "statya=статья",
            "stelka=стелька", "stil=стиль", "stoitsizm=стоицизм", "stronsiy=стронций", "sunami=цунами",
            "sutemizuvchilar=сутэмизувчилар", "tablitsa=таблица", "talk=тальк", "telefilm=телефильм",
            "teleinssenirovka=телеинсценировка", "tendensioz=тенденциоз",
            "tendensiozlik=тенденциозлик", "teplitsa=теплица", "terset=терцет", "tersiya=терция",
            "tigel=тигель", "tonnel=тоннель", "transkripsion=транскрипцион",
            "transkripsiya=транскрипция", "transkripsiyalamoq=транскрипцияламоқ", "tyul=тюль",
            "ultimatum=ультиматум", "ultrabinafsha=ультрабинафша", "ultramikroskop=ультрамикроскоп",
            "ultraqisqa=ультрақисқа", "ultratovush=ультратовуш", "util=утиль", "vaksina=вакцина",
            "valvatsiya=вальвация", "vanil=ваниль", "varyete=варьете", "vedomost=ведомость",
            "verf=верфь", "vernyer=верньер", "vestibyul=вестибюль", "videofilm=видеофильм",
            "viklyuchatel=виключатель", "vintsimon=винтсимон", "vinyetka=виньетка",
            "vklyuchatel=включатель", "vodevil=водевиль", "volost=волость", "voltli=вольтли",
            "xemosorbsiya=хемосорбция", "xolodilnik=холодильник", "xrustal=хрусталь", "yodli=йодли",
            "yotsiramoq=ётсирамоқ", "yuan=юань"
        };

        static readonly Dictionary<char, string> LcMap = new Dictionary<char, string>
        {
            {'a',"а"},{'b',"б"},{'d',"д"},{'e',"е"},{'f',"ф"},{'g',"г"},{'h',"ҳ"},{'i',"и"},{'j',"ж"},
            {'k',"к"},{'l',"л"},{'m',"м"},{'n',"н"},{'o',"о"},{'p',"п"},{'q',"қ"},{'r',"р"},{'s',"с"},
            {'t',"т"},{'u',"у"},{'v',"в"},{'x',"х"},{'y',"й"},{'z',"з"},
            {'ö',"ў"},{'ğ',"ғ"},{'ş',"ш"},{'ç',"ч"}
        };
        const string CyrVowOut = "аеёиоуэюяў";
        const string UzLat = "abdefghijklmnopqrstuvxyzöğşç";     // c va w yo'q
        static bool IsNewLetter(char c) => c == 'ö' || c == 'ğ' || c == 'ş' || c == 'ç';
        static bool IsUzLat(char c) => c != '\0' && UzLat.IndexOf(c) >= 0;

        static string LatKey(string w)
        {
            var sb = new StringBuilder();
            foreach (char c0 in w)
            {
                char c = char.ToLowerInvariant(c0);
                if (c == 'ş') sb.Append("sh"); else if (c == 'ç') sb.Append("ch");
                else if (c == 'ö') sb.Append("o'"); else if (c == 'ğ') sb.Append("g'");
                else if (IsApos(c)) sb.Append('\''); else sb.Append(c);
            }
            return sb.ToString();
        }

        sealed class LatRule { public string Cyr; public bool Prefix; }

        static Dictionary<string, LatRule> ParseLatCyr(IEnumerable<string> list)
        {
            var map = new Dictionary<string, LatRule>();
            foreach (var raw in list)
            {
                var e = (raw ?? "").Trim(); int eq = e.IndexOf('=');
                if (eq <= 0) continue;
                string k = e.Substring(0, eq).Trim(), v = e.Substring(eq + 1).Trim().ToLowerInvariant();
                bool star = k.EndsWith("*");
                if (star) k = k.Substring(0, k.Length - 1);
                k = LatKey(k);
                if (k.Length > 0 && v.Length > 0) map[k] = new LatRule { Cyr = v, Prefix = star || k.Length >= 5 };
            }
            return map;
        }

        /// <summary>Kirill istisnolaridan teskari yozuvlar (faqat kichik harfli lotin shakllari; brendlar emas).</summary>
        static IEnumerable<string> ReverseCyrExceptions(IEnumerable<string> list)
        {
            foreach (var raw in list ?? DefaultCyrExceptions)
            {
                var e = (raw ?? "").Trim(); int eq = e.IndexOf('=');
                if (eq <= 0) continue;
                string cyr = e.Substring(0, eq).Trim(), lat = e.Substring(eq + 1).Trim();
                if (lat.Length > 0 && lat == lat.ToLowerInvariant() && lat.All(ch => IsUzLat(ch) || ch == 'c' || IsApos(ch) || ch == '-'))
                    yield return lat + "=" + cyr;
            }
        }

        static Tuple<int, string> FindLatRule(string key, Dictionary<string, LatRule> map)
        {
            for (int len = key.Length; len > 0; len--)
                if (map.TryGetValue(key.Substring(0, len), out var r) && (len == key.Length || r.Prefix))
                    return Tuple.Create(len, r.Cyr);
            return null;
        }

        static char At(string s, int i) => i >= 0 && i < s.Length ? s[i] : '\0';
        static string Sub(string s, int i, int n) => i >= 0 && i < s.Length ? s.Substring(i, Math.Min(n, s.Length - i)) : "";

        static string LatWordToCyr(string run, bool oldMode, Dictionary<string, LatRule> map)
        {
            string lower = run.ToLowerInvariant();
            var letters = new string(run.Where(char.IsLetter).ToArray());
            bool allCaps = letters.Length > 1 && letters == letters.ToUpperInvariant();
            bool capFirst = run[0] != lower[0];
            if (!allCaps)
                for (int k = 1; k < run.Length; k++)
                    if (char.IsLower(run[k - 1]) && char.IsUpper(run[k])) return null;   // iPhone, McAfee

            var outB = new StringBuilder();
            int i = 0, n = lower.Length;
            var rule = FindLatRule(LatKey(run), map);
            if (rule != null)
            {
                int consumed = 0, kk = 0;
                while (kk < rule.Item1 && consumed < n) { kk += IsNewLetter(lower[consumed]) ? 2 : 1; consumed++; }
                string cyr = rule.Item2;
                if (consumed < n && cyr.EndsWith("ь")) cyr = cyr.Substring(0, cyr.Length - 1);   // мебель + лар
                outB.Append(cyr);
                i = consumed;
            }
            Func<bool> prevVowel = () => outB.Length > 0 && CyrVowOut.IndexOf(outB[outB.Length - 1]) >= 0;

            while (i < n)
            {
                char c = lower[i], nx = At(lower, i + 1), nx2 = At(lower, i + 2);
                bool atStart = outB.Length == 0;
                if ((c == 'o' || c == 'g') && IsApos(nx)) { outB.Append(c == 'o' ? "ў" : "ғ"); i += 2; continue; }
                if (oldMode && c == 's' && nx == 'h') { outB.Append("ш"); i += 2; continue; }
                if (oldMode && c == 'c' && nx == 'h') { outB.Append("ч"); i += 2; continue; }
                if (IsApos(c))
                {
                    if (At(lower, i - 1) == 's' && nx == 'h') { i++; continue; }        // sʼh -> сҳ
                    if (i > 0 && IsUzLat(nx)) { outB.Append("ъ"); i++; continue; }      // tutuq -> ъ
                    outB.Append(run[i]); i++; continue;
                }
                if ((c == 't' || c == 'k') && nx == 's' && nx2 == 'i' && (At(lower, i + 3) == 'y' || At(lower, i + 3) == 'o'))
                {
                    outB.Append(c == 't' ? "ц" : "кц"); i += 2; continue;              // -tsiya, -ksiya
                }
                if (c == 't' && nx == 's' && i > 0)
                {
                    char a3 = At(lower, i + 3), a4 = At(lower, i + 4);
                    if (nx2 == 'e' || (nx2 == 'i' && a3 != '\0' && a3 != 'z' && (a3 != 'n' || a4 == 't') && IsUzLat(a3) &&
                        ("aeiou".IndexOf(a3) < 0 || a3 == 'a' || a3 == 'e')))
                    { outB.Append("ц"); i += 2; continue; }                              // atseton, defitsit
                }
                if (c == 'e' && Sub(lower, i + 1, 3) == "nsi" && (At(lower, i + 4) == 'y' || Sub(lower, i + 4, 2) == "al") &&
                    At(lower, i - 1) != 'p' && outB.Length > 0)
                { outB.Append(prevVowel() ? "э" : "е").Append("нц"); i += 3; continue; } // konferensiya, konfidensial
                if (Sub(lower, i, 5) == "porsi") { outB.Append("порц"); i += 4; continue; }
                if (Sub(lower, i, 8) == "stansiya") { outB.Append("станц"); i += 5; continue; }
                if (c == 'y')
                {
                    if (nx == 'o' && !IsApos(nx2)) { outB.Append("ё"); i += 2; continue; }
                    if (nx == 'a') { outB.Append("я"); i += 2; continue; }
                    if (nx == 'u') { outB.Append("ю"); i += 2; continue; }
                    if (nx == 'e') { outB.Append(atStart || prevVowel() ? "е" : "ъе"); i += 2; continue; }
                    outB.Append("й"); i++; continue;
                }
                if (c == 'e') { outB.Append(atStart || prevVowel() ? "э" : "е"); i++; continue; }
                outB.Append(LcMap.TryGetValue(c, out var m) ? m : run[i].ToString());
                i++;
            }
            string o = outB.ToString();
            if (allCaps) return o.ToUpperInvariant();
            if (capFirst && o.Length > 0) return char.ToUpperInvariant(o[0]) + o.Substring(1);
            return o;
        }

        /// <summary>Lotin matnidagi (yangi yoki amaldagi alifbo) o'zgarishlar — kirillga.</summary>
        public static List<Edit> ComputeLatCyrEdits(string text, ConverterOptions opts)
        {
            var edits = new List<Edit>();
            if (string.IsNullOrEmpty(text)) return edits;
            if (opts.SkipEnglish && IsEnglishText(text)) return edits;
            var map = ParseLatCyr(ReverseCyrExceptions(opts.CyrExceptions).Concat(opts.LatCyr ?? (IEnumerable<string>)DefaultLatCyr));
            var exceptions = opts.Exceptions;
            bool quoteOpen = false;
            int i = 0, n = text.Length;
            while (i < n)
            {
                if (IsSpace(text[i])) { i++; continue; }
                int start = i;
                while (i < n && !IsSpace(text[i])) i++;
                string tok = text.Substring(start, i - start);

                int lead = 0;
                while (lead < tok.Length && !IsAlnum(tok[lead])) { if (IsApos(tok[lead])) quoteOpen = true; lead++; }
                int limit = tok.Length, t = tok.Length;
                while (t > lead && !IsLetter(tok[t - 1]) && !IsApos(tok[t - 1])) t--;
                if (t > lead && IsApos(tok[t - 1]) && quoteOpen) { limit = t - 1; quoteOpen = false; }

                string core = CoreOf(tok);
                if (core.Length == 0 || LooksLikeLinkOrCode(tok)) continue;
                if (FindException(core, exceptions) != null && FindLatRule(LatKey(core), map) == null) continue;
                if (core.Any(IsCyr)) continue;

                int j = lead;
                while (j < limit)
                {
                    if (!IsLetter(tok[j])) { j++; continue; }
                    int ws = j;
                    while (j < limit && (IsLetter(tok[j]) ||
                           (IsApos(tok[j]) && (IsLetter(At(tok, j + 1)) || "ogOG".IndexOf(At(tok, j - 1)) >= 0)))) j++;
                    string run = tok.Substring(ws, j - ws);
                    string rl = run.ToLowerInvariant();
                    bool hasNew = rl.Any(IsNewLetter), foreign = false;
                    for (int q = 0; q < rl.Length; q++)
                    {
                        char ch = rl[q];
                        if (IsApos(ch)) continue;
                        if (ch == 'c') { if (At(rl, q + 1) == 'h' && !hasNew) continue; foreign = true; break; }
                        if (!IsUzLat(ch)) { foreign = true; break; }
                    }
                    if (foreign) continue;
                    string cyr = LatWordToCyr(run, !hasNew, map);
                    if (cyr != null && cyr != run) edits.Add(new Edit { Pos = start + ws, Len = run.Length, From = run, To = cyr });
                }
            }
            return edits;
        }

        public static string LatToCyr(string text, ConverterOptions opts = null)
        {
            var edits = ComputeLatCyrEdits(text, opts ?? new ConverterOptions());
            var sb = new StringBuilder(text);
            for (int k = edits.Count - 1; k >= 0; k--) { sb.Remove(edits[k].Pos, edits[k].Len); sb.Insert(edits[k].Pos, edits[k].To); }
            return sb.ToString();
        }

        // ======================================================== YANGI LOTIN → ESKI LOTIN
        /// <summary>Özbekiston → Oʻzbekiston, ŞAHAR → SHAHAR (har bir so'z — bitta o'zgarish).</summary>
        public static List<Edit> ComputeNewToOldEdits(string text)
        {
            var edits = new List<Edit>();
            if (string.IsNullOrEmpty(text)) return edits;
            int i = 0, n = text.Length;
            while (i < n)
            {
                if (IsSpace(text[i])) { i++; continue; }
                int start = i;
                while (i < n && !IsSpace(text[i])) i++;
                string tok = text.Substring(start, i - start);
                if (!tok.Any(c => IsNewLetter(char.ToLowerInvariant(c)))) continue;
                string o = RevertWord(tok);
                if (o != tok) edits.Add(new Edit { Pos = start, Len = tok.Length, From = tok, To = o });
            }
            return edits;
        }

        public static string NewToOld(string text)
        {
            var edits = ComputeNewToOldEdits(text);
            var sb = new StringBuilder(text);
            for (int k = edits.Count - 1; k >= 0; k--) { sb.Remove(edits[k].Pos, edits[k].Len); sb.Insert(edits[k].Pos, edits[k].To); }
            return sb.ToString();
        }
    }
}
