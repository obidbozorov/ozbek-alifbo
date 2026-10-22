// Özbekça — Word uchun o'rnatuvchi (Özbekça.exe)
//
// Nima qiladi:
//   1. Wordni yashirin ishga tushiradi, yangi hujjatga OzbekAlifbo.bas makros modulini import qiladi
//      va uni Ozbekca.dotm (makrosli shablon) sifatida saqlaydi.
//   2. Shablonga "Yangi alifbo" lentasini (customUI14.xml) qo'shadi.
//   3. Shablonni Wordning STARTUP papkasiga nusxalaydi — Word har safar uni avtomatik yuklaydi.
// Administrator huquqi kerak emas (hammasi joriy foydalanuvchi papkasida).
//
// Buyruq qatori:  Özbekça.exe /S  — oynasiz o'rnatish;  /U — oynasiz o'chirish.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Win32;

[assembly: AssemblyTitle("Özbekça 1.0")]
[assembly: AssemblyDescription("Word, Excel va PowerPoint uçun yangi özbek alifbosi plagini")]
[assembly: AssemblyCompany("Mirzo Uluğbek nomidagi Özbekiston Milliy universiteti")]
[assembly: AssemblyProduct("Özbekça")]
[assembly: AssemblyCopyright("© 2026 Bozorov Obidjon Norqobilovich")]
[assembly: AssemblyTrademark("Özbekça")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]
[assembly: AssemblyInformationalVersion("1.0")]

namespace YangiAlifbo
{
    internal sealed class InstallerException : Exception
    {
        public InstallerException(string msg) : base(msg) { }
    }

    internal static class Installer
    {
        public const string Version = "1.0";
        public const string DotmName = "Ozbekca.dotm";
        const string OldDotmName = "YangiAlifbo.dotm";   // 1.3 va undan oldingi versiyalar
        const string RegKey = @"Software\OzbekAlifbo";
        const string OfficeLibGuid = "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}";
        const string RelsNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        const string UiRelType = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility";
        const string CtNs = "http://schemas.openxmlformats.org/package/2006/content-types";

        // ------------------------------------------------------------ holat
        public static string DefaultStartupDir =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Microsoft\Word\STARTUP");

        public static string InstalledPath
        {
            get
            {
                using (var k = Registry.CurrentUser.OpenSubKey(RegKey))
                {
                    var p = k?.GetValue("InstallPath") as string;
                    if (!string.IsNullOrEmpty(p) && File.Exists(p)) return p;
                }
                var d = Path.Combine(DefaultStartupDir, DotmName);
                if (File.Exists(d)) return d;
                d = Path.Combine(DefaultStartupDir, OldDotmName);
                return File.Exists(d) ? d : null;
            }
        }

        public static string InstalledVersion
        {
            get { using (var k = Registry.CurrentUser.OpenSubKey(RegKey)) return k?.GetValue("Version") as string; }
        }

        public static bool WordIsRunning => Process.GetProcessesByName("WINWORD").Length > 0;

        /// <summary>O'rnatilgan Word versiyasi ("16.0"), topilmasa null.</summary>
        public static string WordVersion()
        {
            using (var k = Registry.ClassesRoot.OpenSubKey(@"Word.Application\CurVer"))
            {
                var cur = k?.GetValue("") as string;            // "Word.Application.16"
                if (cur != null)
                {
                    var n = cur.Substring(cur.LastIndexOf('.') + 1);
                    if (int.TryParse(n, out _)) return n + ".0";
                }
            }
            try { return Type.GetTypeFromProgID("Word.Application") != null ? "16.0" : null; }
            catch { return null; }
        }

        // ------------------------------------------------------------ o'rnatish (hammasi)
        public const string XlamName = "Ozbekca.xlam", PpamName = "Ozbekca.ppam";
        // Excel/PowerPoint qo'shimchalarining odatiy papkasi
        public static string AddInsDir =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Microsoft\AddIns");

        static readonly string[][] OfficeProcs = { new[] { "WINWORD", "Word" }, new[] { "EXCEL", "Excel" }, new[] { "POWERPNT", "PowerPoint" } };

        /// <summary>Hozir ochiq turgan Office dasturlari (Word, Excel, PowerPoint).</summary>
        public static List<string> RunningApps() =>
            OfficeProcs.Where(p => Process.GetProcessesByName(p[0]).Length > 0).Select(p => p[1]).ToList();

        /// <summary>"Excel.Application" -> "16.0"; topilmasa null.</summary>
        public static string AppVersion(string progId)
        {
            using (var k = Registry.ClassesRoot.OpenSubKey(progId + @"\CurVer"))
            {
                var cur = k?.GetValue("") as string;
                if (cur != null)
                {
                    var n = cur.Substring(cur.LastIndexOf('.') + 1);
                    if (int.TryParse(n, out _)) return n + ".0";
                }
            }
            try { return Type.GetTypeFromProgID(progId) != null ? "16.0" : null; }
            catch { return null; }
        }

        public static List<string> AvailableApps()
        {
            var l = new List<string>();
            if (WordVersion() != null) l.Add("Word");
            if (AppVersion("Excel.Application") != null) l.Add("Excel");
            if (AppVersion("PowerPoint.Application") != null) l.Add("PowerPoint");
            return l;
        }

        public static List<string> InstalledApps()
        {
            var l = new List<string>();
            if (InstalledPath != null) l.Add("Word");
            if (File.Exists(Path.Combine(AddInsDir, XlamName))) l.Add("Excel");
            if (File.Exists(Path.Combine(AddInsDir, PpamName))) l.Add("PowerPoint");
            return l;
        }

        /// <summary>Kompyuterdagi Word, Excel va PowerPoint ga o'rnatadi; o'rnatilganlar ro'yxatini qaytaradi.</summary>
        public static List<string> Install(Action<string> log)
        {
            var running = RunningApps();
            if (running.Count > 0)
                throw new InstallerException(string.Join(", ", running) + " oçiq. Ularni yopib, qayta urinib köring.");
            var done = new List<string>();
            var errors = new List<string>();
            var apps = AvailableApps();
            if (apps.Count == 0) throw new InstallerException("Microsoft Word, Excel yoki PowerPoint topilmadi. Avval Microsoft Office ni örnating.");

            foreach (var app in new[] { "Word", "Excel", "PowerPoint" })
            {
                if (!apps.Contains(app)) { log(app + " topilmadi — ötkazib yuborildi."); log(""); continue; }
                log("━━━ " + app + " ━━━");
                try
                {
                    if (app == "Word") InstallWord(log);
                    else InstallAddin(app, log);
                    done.Add(app);
                }
                catch (InstallerException ex) { errors.Add(app); log("XATO (" + app + "): " + ex.Message); }
                catch (Exception ex) { errors.Add(app); log("XATO (" + app + "): " + ex.Message); }
                log("");
            }
            if (done.Count == 0) throw new InstallerException("Hech bir dasturga örnatib bölmadi. Yuqoridagi xabarlarni köring.");
            using (var k = Registry.CurrentUser.CreateSubKey(RegKey)) k.SetValue("Version", Version);
            log("Tayyor! " + string.Join(", ", done) + " ni oçing — yuqorida «Özbekça» lentasi paydo böladi.");
            if (errors.Count > 0) log("Örnatilmadi: " + string.Join(", ", errors) + " (sababi yuqorida).");
            return done;
        }

        // ------------------------------------------------------------ Word
        static string InstallWord(Action<string> log)
        {
            string ver = WordVersion();

            string work = Path.Combine(Path.GetTempPath(), "YangiAlifbo_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(work);
            string basPath = Path.Combine(work, "OzbekAlifbo.bas");
            File.WriteAllBytes(basPath, Resource("OzbekAlifbo.bas"));
            string uiXml = Encoding.UTF8.GetString(Resource("customUI14.xml"));
            string tmpDotm = Path.Combine(work, DotmName);
            string startup = null;

            log("1/5. Microsoft Word topildi (versiya " + ver + ").");
            string secKey = @"Software\Microsoft\Office\" + ver + @"\Word\Security";
            object oldVbom = ReadValue(secKey, "AccessVBOM");
            WriteDword(secKey, "AccessVBOM", 1);   // makros modulini import qilish uchun vaqtincha ruxsat
            object app = null, doc = null;
            try
            {
                log("2/5. Plagin tayyorlanmoqda, biroz kuting...");
                app = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));
                Set(app, "Visible", false);
                Set(app, "DisplayAlerts", 0);
                startup = Get(app, "StartupPath") as string;

                doc = Call(Get(app, "Documents"), "Add");
                object vbp;
                try { vbp = Get(doc, "VBProject"); Get(vbp, "Name"); }
                catch (Exception)
                {
                    throw new InstallerException(
                        "Word makros loyihasiga kirişga ruxsat bermadi (taşkilot siyosati bölişi mumkin).\n" +
                        "Qölda örnatiş: Word → Alt+F11 → File → Import File → OzbekAlifbo.bas (README ga qarang).");
                }
                try { Call(Get(vbp, "References"), "AddFromGuid", OfficeLibGuid, 2, 0); } catch { /* allaqachon bor */ }
                try { Set(vbp, "Name", "Ozbekca"); } catch { }

                log("3/5. Ögiriş funksiyalari qöşilmoqda...");
                Call(Get(vbp, "VBComponents"), "Import", basPath);

                log("4/5. Tezkor tugmalar tekşirilib, böşlari qöyilmoqda...");
                int bound = BindHotkeys(app, doc, startup, log);
                log(bound == Hotkeys.Length + 4
                    ? "        " + string.Join(", ", ChosenKeys.Values) + " — tayyor."
                    : "        Diqqat: " + bound + " / " + (Hotkeys.Length + 4) + " ta tugma qöyildi.");

                Call(doc, "SaveAs2", tmpDotm, 15);           // 15 = wdFormatXMLTemplateMacroEnabled
                Call(doc, "Close", false);
                Release(ref doc);
                Call(app, "Quit", false);
                Release(ref app);
            }
            catch (TargetInvocationException ex)
            {
                throw new InstallerException("Word xatosi: " + (ex.InnerException?.Message ?? ex.Message));
            }
            finally
            {
                if (doc != null) { try { Call(doc, "Close", false); } catch { } Release(ref doc); }
                if (app != null) { try { Call(app, "Quit", false); } catch { } Release(ref app); }
                RestoreValue(secKey, "AccessVBOM", oldVbom);
                GC.Collect(); GC.WaitForPendingFinalizers();
            }

            log("5/5. Wordga «Özbekça» lentasi va ikonkalar qöşilmoqda...");
            InjectRibbon(tmpDotm, uiXml);

            if (string.IsNullOrEmpty(startup)) startup = DefaultStartupDir;
            Directory.CreateDirectory(startup);
            string target = Path.Combine(startup, DotmName);
            try
            {
                // eski nusxa (boshqa joyda bo'lsa) — olib tashlaymiz
                var old = InstalledPath;
                if (old != null && !string.Equals(old, target, StringComparison.OrdinalIgnoreCase)) File.Delete(old);
                var legacy = Path.Combine(startup, OldDotmName);
                if (File.Exists(legacy)) File.Delete(legacy);
                File.Copy(tmpDotm, target, true);
            }
            catch (IOException)
            {
                throw new InstallerException("Şablon fayli band. Word töliq yopilganiga işonç hosil qilib, qayta urinib köring.");
            }

            using (var k = Registry.CurrentUser.CreateSubKey(RegKey))
            {
                k.SetValue("InstallPath", target);
                k.SetValue("Version", Version);
            }
            try { Directory.Delete(work, true); } catch { }
            log("        Word: tayyor.");
            return target;
        }

        // ------------------------------------------------------------ Excel, PowerPoint
        /// <summary>Ozbekca.xlam / Ozbekca.ppam yasaydi, lentani qo'shadi va dasturga ro'yxatdan o'tkazadi.</summary>
        static void InstallAddin(string app, Action<string> log)
        {
            bool xl = app == "Excel";
            string progId = app + ".Application";
            string ver = AppVersion(progId);
            string work = Path.Combine(Path.GetTempPath(), "Ozbekca_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(work);
            string basPath = Path.Combine(work, "OzbekAlifbo.bas");
            File.WriteAllBytes(basPath, Resource("OzbekAlifbo_" + app + ".bas"));
            string uiXml = Encoding.UTF8.GetString(Resource("customUI14_" + app.ToLowerInvariant() + ".xml"));
            string fileName = xl ? XlamName : PpamName;
            string tmp = Path.Combine(work, fileName);

            log("1/3. Microsoft " + app + " topildi (versiya " + ver + ").");
            // eski qo'shimcha yasash paytida yuklanmasin (buzilgan bo'lsa, dastur xato oynasi chiqarib to'xtab qoladi)
            try
            {
                if (xl) UnregisterExcel(ver);
                else Registry.CurrentUser.DeleteSubKeyTree(@"Software\Microsoft\Office\" + ver + @"\PowerPoint\AddIns\Ozbekca", false);
                var oldFile = Path.Combine(AddInsDir, fileName);
                if (File.Exists(oldFile)) File.Delete(oldFile);
            }
            catch { }
            string secKey = @"Software\Microsoft\Office\" + ver + @"\" + app + @"\Security";
            object oldVbom = ReadValue(secKey, "AccessVBOM");
            WriteDword(secKey, "AccessVBOM", 1);
            object o = null, doc = null;
            try
            {
                log("2/3. Ögiriş funksiyalari qöşilmoqda...");
                o = Activator.CreateInstance(Type.GetTypeFromProgID(progId));
                if (xl)
                {
                    Set(o, "Visible", false);
                    Set(o, "DisplayAlerts", false);
                    doc = Call(Get(o, "Workbooks"), "Add");
                }
                else
                {
                    try { Set(o, "DisplayAlerts", 1); } catch { }          // ppAlertsNone
                    // oynasiz (WithWindow=0) yaratilgan taqdimotni PowerPoint qo'shimcha sifatida bo'sh saqlaydi
                    doc = Call(Get(o, "Presentations"), "Add", -1);
                    try { Set(Get(o, "ActiveWindow"), "WindowState", 2); } catch { }   // ppWindowMinimized
                }
                object vbp;
                try { vbp = Get(doc, "VBProject"); Get(vbp, "Name"); }
                catch (Exception)
                {
                    throw new InstallerException(app + " makros loyihasiga kirişga ruxsat bermadi (taşkilot siyosati bölişi mumkin).");
                }
                try { Call(Get(vbp, "References"), "AddFromGuid", OfficeLibGuid, 2, 0); } catch { /* allaqachon bor */ }
                Call(Get(vbp, "VBComponents"), "Import", basPath);
                Call(doc, "SaveAs", tmp, xl ? 55 : 30);                  // xlOpenXMLAddIn / ppSaveAsOpenXMLAddin
                if (xl) Call(doc, "Close", false); else Call(doc, "Close");
                Release(ref doc);
                Call(o, "Quit");
                Release(ref o);
            }
            catch (TargetInvocationException ex)
            {
                throw new InstallerException(app + " xatosi: " + (ex.InnerException?.Message ?? ex.Message));
            }
            finally
            {
                if (doc != null) { try { if (xl) Call(doc, "Close", false); else Call(doc, "Close"); } catch { } Release(ref doc); }
                if (o != null) { try { Call(o, "Quit"); } catch { } Release(ref o); }
                RestoreValue(secKey, "AccessVBOM", oldVbom);
                GC.Collect(); GC.WaitForPendingFinalizers();
            }

            log("3/3. «Özbekça» lentasi va ikonkalar qöşilmoqda...");
            InjectRibbon(tmp, uiXml);
            if (LastFix != null) log("        Eslatma: " + LastFix + ".");
            Directory.CreateDirectory(AddInsDir);
            string target = Path.Combine(AddInsDir, fileName);
            try { File.Copy(tmp, target, true); }
            catch (IOException) { throw new InstallerException(fileName + " band. " + app + " töliq yopilganiga işonç hosil qiling."); }
            if (xl) RegisterExcel(ver, target);
            else RegisterPowerPoint(ver, target);
            try { Directory.Delete(work, true); } catch { }
            log("        " + app + ": tayyor." + (xl ? " Tezkor tugmalar: Alt+A, Alt+E, Alt+C, Alt+I, Alt+O." : ""));
        }

        // Excel ishga tushganda ochadigan qo'shimchalar: ...\Excel\Options  OPEN, OPEN1, OPEN2 ... (ketma-ket bo'lishi shart)
        static List<string> ExcelOpenList(RegistryKey k) =>
            Enumerable.Range(0, 100).Select(i => i == 0 ? "OPEN" : "OPEN" + i)
                      .Select(n => k.GetValue(n) as string).TakeWhile(v => v != null).ToList();

        static void WriteExcelOpenList(RegistryKey k, List<string> vals)
        {
            foreach (var n in k.GetValueNames().Where(n => n.StartsWith("OPEN", StringComparison.OrdinalIgnoreCase)
                                                            && n.Substring(4).All(char.IsDigit)))
                k.DeleteValue(n, false);
            for (int i = 0; i < vals.Count; i++) k.SetValue(i == 0 ? "OPEN" : "OPEN" + i, vals[i]);
        }

        static void RegisterExcel(string ver, string path)
        {
            using (var k = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Office\" + ver + @"\Excel\Options"))
            {
                var vals = ExcelOpenList(k).Where(v => v.IndexOf(XlamName, StringComparison.OrdinalIgnoreCase) < 0).ToList();
                vals.Add("\"" + path + "\"");
                WriteExcelOpenList(k, vals);
            }
        }

        static void UnregisterExcel(string ver)
        {
            using (var k = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\" + ver + @"\Excel\Options", true))
            {
                if (k == null) return;
                var all = ExcelOpenList(k);
                var vals = all.Where(v => v.IndexOf(XlamName, StringComparison.OrdinalIgnoreCase) < 0).ToList();
                if (vals.Count != all.Count) WriteExcelOpenList(k, vals);
            }
        }

        static void RegisterPowerPoint(string ver, string path)
        {
            using (var k = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Office\" + ver + @"\PowerPoint\AddIns\Ozbekca"))
            {
                k.SetValue("Path", path);
                k.SetValue("AutoLoad", 1, RegistryValueKind.DWord);
            }
        }

        // Tezkor tugmalar. Hujjat hali oddiy hujjat bo'lganda (shablonga aylanmasdan oldin) qo'yiladi,
        // shuning uchun ular Ozbekca.dotm ning o'ziga yoziladi. Har bir amal uchun nomzodlar tartibda
        // tekshiriladi: Wordning o'zi, Normal.dotm yoki boshqa plagin (MathType, Zotero ...) band qilgan
        // tugma tashlab o'tiladi va birinchi bo'sh tugma olinadi.
        //   { amal id, makros, 1-harf, 2-harf }
        static readonly string[][] Hotkeys =
        {
            new[] { "Lotin", "YangiAlifbo_Lotin", "A", "U" },
            new[] { "Eski", "YangiAlifbo_EskiLotinAvto", "E", "D" },
            new[] { "LotinKirill", "YangiAlifbo_LotinKirillAvto", "C", "T" },
            new[] { "Yozishda", "YangiAlifbo_Yozishda", "Z", "Q" },
            new[] { "Istisno", "YangiAlifbo_IstisnoAvto", "I", "J" },
            new[] { "Istisnolar", "YangiAlifbo_Istisnolar", "O", "B" },
        };
        const int WdKeyCategoryMacro = 2, WdKeyCategoryNil = -1, WdKeyCategoryDisable = 0,
                  WdKeyShift = 256, WdKeyControl = 512, WdKeyAlt = 1024;
        const string VbaKey = @"Software\VB and VBA Program Settings\OzbekAlifbo\Tugmalar";

        /// <summary>O'rnatishda tanlangan tugmalar: amal id -> "Alt+A".</summary>
        public static readonly Dictionary<string, string> ChosenKeys = new Dictionary<string, string>();

        static string KeyName(int code)
        {
            string n = "";
            if ((code & WdKeyControl) != 0) n += "Ctrl+";
            if ((code & WdKeyAlt) != 0) n += "Alt+";
            if ((code & WdKeyShift) != 0) n += "Shift+";
            return n + (char)(code & 255);
        }

        static bool IsOurs(string cmd) =>
            cmd.IndexOf("YangiAlifbo_", StringComparison.OrdinalIgnoreCase) >= 0 ||
            cmd.IndexOf("Kalit_", StringComparison.OrdinalIgnoreCase) >= 0;

        /// <summary>Tugma kim tomonidan band: null — bo'sh (yoki bizniki).</summary>
        static string BusyBy(object app, int code)
        {
            try
            {
                object kb = Call(app, "FindKey", code);
                int cat = Convert.ToInt32(Get(kb, "KeyCategory"));
                string cmd = Get(kb, "Command") as string ?? "";
                if (cat == WdKeyCategoryDisable) return "öçirib qöyilgan";
                if (cat == WdKeyCategoryNil || cmd.Length == 0 || IsOurs(cmd)) return null;
                return cmd;
            }
            catch { return null; }
        }

        /// <summary>Boshqa global shablonlarni (STARTUP dagi) yuklaydi — ularning tugmalari ham hisobga olinsin.</summary>
        static void LoadOtherAddIns(object app, string startup)
        {
            if (string.IsNullOrEmpty(startup) || !Directory.Exists(startup)) return;
            object addIns;
            try { addIns = Get(app, "AddIns"); } catch { return; }
            foreach (var f in Directory.GetFiles(startup))
            {
                string ext = Path.GetExtension(f).ToLowerInvariant(), name = Path.GetFileName(f);
                if (ext != ".dot" && ext != ".dotm" && ext != ".dotx") continue;
                if (name.Equals(DotmName, StringComparison.OrdinalIgnoreCase) || name.Equals(OldDotmName, StringComparison.OrdinalIgnoreCase)) continue;
                if (name.StartsWith("~$")) continue;
                try { Call(addIns, "Add", f, true); } catch { }
            }
        }

        static int BindHotkeys(object app, object doc, string startup, Action<string> log)
        {
            int ok = 0;
            ChosenKeys.Clear();
            LoadOtherAddIns(app, startup);
            try { Set(app, "CustomizationContext", doc); } catch { return 0; }
            object kbs = Get(app, "KeyBindings");
            var taken = new HashSet<int>();
            using (var reg = Registry.CurrentUser.CreateSubKey(VbaKey))
            {
            // foydalanuvchi "Tezkor tugmalar" oynasida o'zi tanlagan tugmalar (Normal.dotm da) — ularga tegmaymiz
            var manual = new HashSet<string>();
            foreach (var hk in Hotkeys)
                if ((reg.GetValue("Qolda_" + hk[0]) as string) == "1" && int.TryParse(reg.GetValue("Kod_" + hk[0]) as string, out var uc))
                { manual.Add(hk[0]); taken.Add(uc); }
            foreach (var hk in Hotkeys)
            {
                var cands = new List<int>
                {
                    WdKeyAlt + hk[2][0], WdKeyAlt + hk[3][0],
                    WdKeyAlt + WdKeyShift + hk[2][0], WdKeyAlt + WdKeyShift + hk[3][0],
                    WdKeyControl + WdKeyAlt + WdKeyShift + hk[2][0],
                };
                int code = -1; string firstBusy = null;
                foreach (var c in cands)
                {
                    if (taken.Contains(c)) continue;
                    var by = BusyBy(app, c);
                    if (by == null) { code = c; break; }
                    if (firstBusy == null) firstBusy = KeyName(c) + " band (" + by + ")";
                }
                if (code < 0) code = cands[0];          // hammasi band — asosiysini baribir qo'yamiz
                taken.Add(code);
                if (manual.Contains(hk[0])) ok++;          // foydalanuvchining o'z tugmasi Normal.dotm da turibdi
                else
                foreach (var name in new[] { "OzbekAlifbo." + hk[1], hk[1] })
                {
                    try { Call(kbs, "Add", WdKeyCategoryMacro, name, code); ok++; break; }
                    catch { /* keyingi nom shaklini sinaymiz */ }
                }
                reg.SetValue("Avto_" + hk[0], KeyName(code));
                reg.SetValue("AvtoKod_" + hk[0], code.ToString());
                if (manual.Contains(hk[0]))
                {
                    ChosenKeys[hk[0]] = reg.GetValue(hk[0]) as string ?? KeyName(code);   // o'zi tanlagani qoladi
                    continue;
                }
                ChosenKeys[hk[0]] = KeyName(code);
                reg.SetValue(hk[0], KeyName(code));
                reg.SetValue("Kod_" + hk[0], code.ToString());
                if (code != cands[0] && firstBusy != null)
                    log("        " + firstBusy + " → " + KeyName(code) + " qöyildi");
            }
            }
            // Eski versiyalar Normal.dotm ga bog'lagan h, ', ` tugmalarini tozalaymiz — aks holda Word ularni yozmaydi.
            // (faqat Normal.dotm da SAQLANGAN bog'lanishlar)
            try
            {
                object normal = Get(app, "NormalTemplate");
                Set(app, "CustomizationContext", normal);
                bool changed = false;
                foreach (int code in new[] { 72, WdKeyShift + 72, 222, 192 })
                {
                    try
                    {
                        object kb = Call(app, "FindKey", code);
                        string ctxName = null;
                        try { ctxName = Get(Get(kb, "Context"), "Name") as string; } catch { }
                        if ((Get(kb, "Command") as string ?? "").IndexOf("Kalit_", StringComparison.OrdinalIgnoreCase) >= 0
                            && string.Equals(ctxName, Get(normal, "Name") as string, StringComparison.OrdinalIgnoreCase))
                        { Call(kb, "Clear"); changed = true; }
                    }
                    catch { }
                }
                if (changed) { Call(normal, "Save"); log("        Eski «h», «'» tugma bog'lanişlari tozalandi."); }
                Set(app, "CustomizationContext", doc);
            }
            catch { try { Set(app, "CustomizationContext", doc); } catch { } }

            // Yozish paytida o'girish: h, Shift+H, ', ` — PLAGIN SHABLONINING O'ZIGA bog'lanadi
            // (holatni lentadagi tugma yoqadi/o'chiradi; o'chiq paytda makros harfni odatdagidek yozadi)
            var live = new[] { new object[] { 72, "Kalit_h" }, new object[] { WdKeyShift + 72, "Kalit_ShiftH" },
                               new object[] { 222, "Kalit_Apostrof" }, new object[] { 192, "Kalit_Backtick" } };
            foreach (var lk in live)
            {
                foreach (var name in new[] { "OzbekAlifbo." + lk[1], (string)lk[1] })
                {
                    try { Call(kbs, "Add", WdKeyCategoryMacro, name, (int)lk[0]); ok++; break; }
                    catch { }
                }
            }
            return ok;
        }

        /// <summary>Foydalanuvchi o'zi qo'ygan tugmalarni (Normal.dotm) tozalaydi.</summary>
        static void CleanNormal(Action<string> log)
        {
            string disabled;
            bool any = false;
            using (var reg = Registry.CurrentUser.OpenSubKey(VbaKey))
            {
                if (reg == null) return;
                disabled = reg.GetValue("Ochirilgan") as string ?? "";
                foreach (var hk in Hotkeys) if ((reg.GetValue("Qolda_" + hk[0]) as string) == "1") any = true;
            }
            if (!any && disabled.Length == 0) return;
            log("Normal.dotm dagi tezkor tugmalar tozalanmoqda...");
            object app = null;
            try
            {
                app = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));
                Set(app, "Visible", false);
                Set(app, "DisplayAlerts", 0);
                object normal = Get(app, "NormalTemplate");
                Set(app, "CustomizationContext", normal);
                object kbs = Get(app, "KeyBindings");
                int n = Convert.ToInt32(Get(kbs, "Count"));
                for (int i = n; i >= 1; i--)
                {
                    try
                    {
                        object kb = Call(kbs, "Item", i);
                        if (IsOurs(Get(kb, "Command") as string ?? "")) Call(kb, "Clear");
                    }
                    catch { }
                }
                foreach (var c in disabled.Split(','))
                    if (int.TryParse(c, out var code)) { try { Call(Call(app, "FindKey", code), "Clear"); } catch { } }
                Call(normal, "Save");
            }
            catch { log("  (Normal.dotm ni tozalab bölmadi — zarari yöq)"); }
            finally
            {
                if (app != null) { try { Call(app, "Quit", false); } catch { } Release(ref app); }
                GC.Collect(); GC.WaitForPendingFinalizers();
            }
        }

        // ------------------------------------------------------------ o'chirish
        public static bool Uninstall(Action<string> log)
        {
            var installed = InstalledApps();
            if (installed.Count == 0) { log("Plagin örnatilmagan."); return false; }
            var running = RunningApps();
            if (running.Count > 0)
                throw new InstallerException(string.Join(", ", running) + " oçiq. Ularni yopib, qayta urinib köring.");
            var path = InstalledPath;
            if (path != null)
            {
                try
                {
                    File.Delete(path);
                    var legacy = Path.Combine(DefaultStartupDir, OldDotmName);
                    if (File.Exists(legacy)) File.Delete(legacy);
                }
                catch (IOException) { throw new InstallerException("Fayl band — Wordni yoping va qayta urinib köring."); }
                CleanNormal(log);
                try { Registry.CurrentUser.DeleteSubKeyTree(VbaKey, false); } catch { }
                log("Word: plagin olib taşlandi.");
            }
            foreach (var app in new[] { "Excel", "PowerPoint" })
            {
                string fn = app == "Excel" ? XlamName : PpamName;
                string f = Path.Combine(AddInsDir, fn);
                string ver = AppVersion(app + ".Application") ?? "16.0";

                try
                {
                    if (app == "Excel") UnregisterExcel(ver);
                    else Registry.CurrentUser.DeleteSubKeyTree(@"Software\Microsoft\Office\" + ver + @"\PowerPoint\AddIns\Ozbekca", false);
                }
                catch { }
                if (File.Exists(f))
                {
                    try { File.Delete(f); }
                    catch (IOException) { throw new InstallerException(app + " fayli band — " + app + " ni yoping va qayta urinib köring."); }
                    log(app + ": plagin olib taşlandi.");
                }
            }
            using (var k = Registry.CurrentUser.OpenSubKey(RegKey, true))
            {
                k?.DeleteValue("InstallPath", false);
                k?.DeleteValue("Version", false);
            }
            log("Istisno sözlaringiz saqlanib qoldi: %APPDATA%\\OzbekAlifbo\\istisnolar.txt");
            return true;
        }

        // ------------------------------------------------------------ lenta (customUI)
        /// <summary>.dotm (zip) ichiga customUI/customUI14.xml va tegishli bog'lanishni qo'shadi.</summary>
        public static string LastFix;

        public static void InjectRibbon(string dotmPath, string uiXml)
        {
            // Faylni "yangilash" rejimida ochmaymiz (PowerPoint fayllarida qismlarni buzib qo'yardi):
            // hamma qismlar o'qiladi, xotirada o'zgartiriladi va yangi zip qaytadan yoziladi.
            var parts = new List<KeyValuePair<string, byte[]>>();
            using (var z = ZipFile.OpenRead(dotmPath))
                foreach (var e in z.Entries)
                {
                    if (e.FullName.EndsWith("/")) continue;
                    using (var src = e.Open())
                    using (var ms = new MemoryStream())
                    {
                        src.CopyTo(ms);
                        parts.Add(new KeyValuePair<string, byte[]>(e.FullName, ms.ToArray()));
                    }
                }
            Func<string, byte[]> Get = n => parts.FirstOrDefault(x => x.Key == n).Value;
            Action<string, byte[]> Put = (n, data) =>
            {
                parts.RemoveAll(x => x.Key == n);
                parts.Add(new KeyValuePair<string, byte[]>(n, data));
            };
            Func<XDocument, byte[]> XmlBytes = d =>
            {
                using (var ms = new MemoryStream()) { SaveXml(d, ms); return ms.ToArray(); }
            };

            // PowerPoint ba'zan qo'shimchaning asosiy qismini bo'sh saqlaydi — minimal to'g'ri tarkib yoziladi
            LastFix = null;
            var pres = Get("ppt/presentation.xml");
            if (pres != null && (pres.Length == 0 || Array.IndexOf(pres, (byte)'<') < 0))
            {
                Put("ppt/presentation.xml", new UTF8Encoding(false).GetBytes(
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                    "<p:presentation xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                    "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">" +
                    "<p:sldSz cx=\"12192000\" cy=\"6858000\"/><p:notesSz cx=\"6858000\" cy=\"9144000\"/></p:presentation>"));
                LastFix = "ppt/presentation.xml bo'sh edi — tiklandi";
            }
            var empty = parts.Where(x => x.Key.EndsWith(".xml") && x.Value.Length == 0).Select(x => x.Key).ToList();
            if (empty.Count > 0)
                throw new InstallerException("Office faylni noto'g'ri saqladi (bo'sh qism: " + string.Join(", ", empty) + ").");

            // lenta
            parts.RemoveAll(x => x.Key.StartsWith("customUI/"));
            Put("customUI/customUI14.xml", new UTF8Encoding(false).GetBytes(uiXml));

            var relsData = Get("_rels/.rels") ?? throw new InstallerException("Şablon tuzilişi notöğri (_rels/.rels yöq).");
            XDocument rels;
            using (var ms = new MemoryStream(relsData)) rels = XDocument.Load(ms);
            XNamespace rn = RelsNs;
            rels.Root.Elements(rn + "Relationship").Where(e => (string)e.Attribute("Type") == UiRelType).Remove();
            var ids = rels.Root.Elements(rn + "Relationship").Select(e => (string)e.Attribute("Id")).ToList();
            string id = "rIdYangiAlifbo";
            while (ids.Contains(id)) id += "1";
            rels.Root.Add(new XElement(rn + "Relationship",
                new XAttribute("Id", id), new XAttribute("Type", UiRelType), new XAttribute("Target", "customUI/customUI14.xml")));
            Put("_rels/.rels", XmlBytes(rels));

            // ikonkalar: customUI/images/*.png + customUI/_rels/customUI14.xml.rels
            var asm = Assembly.GetExecutingAssembly();
            var icons = asm.GetManifestResourceNames().Where(n => n.StartsWith("icon_", StringComparison.Ordinal)).ToList();
            var uiRels = new XElement(rn + "Relationships");
            foreach (var res in icons)
            {
                string iconId = res.Substring(5, res.Length - 9);           // icon_iYangi.png -> iYangi
                using (var src = asm.GetManifestResourceStream(res))
                using (var ms = new MemoryStream()) { src.CopyTo(ms); Put("customUI/images/" + iconId + ".png", ms.ToArray()); }
                uiRels.Add(new XElement(rn + "Relationship", new XAttribute("Id", iconId),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"),
                    new XAttribute("Target", "images/" + iconId + ".png")));
            }
            if (icons.Count > 0) Put("customUI/_rels/customUI14.xml.rels", XmlBytes(new XDocument(uiRels)));

            var ctData = Get("[Content_Types].xml");
            if (ctData != null)
            {
                XDocument ct;
                using (var ms = new MemoryStream(ctData)) ct = XDocument.Load(ms);
                XNamespace cn = CtNs;
                foreach (var ext in new[] { new[] { "xml", "application/xml" }, new[] { "png", "image/png" } })
                {
                    bool has = ct.Root.Elements(cn + "Default")
                        .Any(e => string.Equals((string)e.Attribute("Extension"), ext[0], StringComparison.OrdinalIgnoreCase));
                    if (!has)
                        ct.Root.AddFirst(new XElement(cn + "Default", new XAttribute("Extension", ext[0]), new XAttribute("ContentType", ext[1])));
                }
                Put("[Content_Types].xml", XmlBytes(ct));
            }

            // yangi zip: [Content_Types].xml birinchi
            string tmp = dotmPath + ".new";
            using (var fs = File.Create(tmp))
            using (var z = new ZipArchive(fs, ZipArchiveMode.Create))
                foreach (var part in parts.OrderBy(x => x.Key == "[Content_Types].xml" ? 0 : 1))
                    using (var dst = z.CreateEntry(part.Key, CompressionLevel.Optimal).Open())
                        dst.Write(part.Value, 0, part.Value.Length);
            File.Delete(dotmPath);
            File.Move(tmp, dotmPath);
        }

        // ------------------------------------------------------------ yordamchilar
        static void SaveXml(XDocument doc, Stream s)          // BOM siz UTF-8 (Office fayllaridagi kabi)
        {
            var settings = new System.Xml.XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = false };
            using (var w = System.Xml.XmlWriter.Create(s, settings)) doc.Save(w);
        }

        static byte[] Resource(string name)
        {
            var asm = Assembly.GetExecutingAssembly();
            var full = asm.GetManifestResourceNames().First(n => n.EndsWith(name, StringComparison.OrdinalIgnoreCase));
            using (var s = asm.GetManifestResourceStream(full))
            using (var m = new MemoryStream()) { s.CopyTo(m); return m.ToArray(); }
        }

        static object Get(object o, string p, params object[] a) => o.GetType().InvokeMember(p, BindingFlags.GetProperty, null, o, a);
        static object Call(object o, string m, params object[] a) => o.GetType().InvokeMember(m, BindingFlags.InvokeMethod, null, o, a);
        static void Set(object o, string p, object v) => o.GetType().InvokeMember(p, BindingFlags.SetProperty, null, o, new[] { v });
        static void Release(ref object o) { if (o != null && Marshal.IsComObject(o)) Marshal.FinalReleaseComObject(o); o = null; }

        static object ReadValue(string key, string name)
        {
            using (var k = Registry.CurrentUser.OpenSubKey(key)) return k?.GetValue(name);
        }
        static void WriteDword(string key, string name, int v)
        {
            using (var k = Registry.CurrentUser.CreateSubKey(key)) k.SetValue(name, v, RegistryValueKind.DWord);
        }
        static void RestoreValue(string key, string name, object old)
        {
            try
            {
                using (var k = Registry.CurrentUser.CreateSubKey(key))
                {
                    if (old == null) k.DeleteValue(name, false);
                    else k.SetValue(name, old, RegistryValueKind.DWord);
                }
            }
            catch { }
        }
    }

    // ================================================================ oyna
    internal sealed class SetupForm : Form
    {
        static readonly Color Accent = Color.FromArgb(30, 110, 90);
        readonly TextBox _log = new TextBox();
        readonly Label _state = new Label();
        readonly Button _btnInstall = new Button();
        readonly Button _btnRemove = new Button();
        readonly Button _btnClose = new Button();
        readonly ProgressBar _bar = new ProgressBar();

        public SetupForm()
        {
            Text = "Özbekça 1.0 — örnatiş";
            SuspendLayout();
            // Masshtabga bog'liq emas: oyna har doim 100% dagidek chiziladi, kattalashtirishni
            // Windowsning o'zi butun oynaga bir xil qo'llaydi — joylashuv hech qachon buzilmaydi
            AutoScaleMode = AutoScaleMode.None;
            Font = new Font("Times New Roman", 11f);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(520, 470);
            BackColor = Color.White;
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

            // Hamma narsa jadval ichida: matn qancha joy olsa, pastdagilar shuncha suriladi —
            // hech narsa bir-birining ustiga tushmaydi (har qanday masshtabda)
            var header = new TableLayoutPanel
            {
                Dock = DockStyle.Top, AutoSize = true, BackColor = Accent, ColumnCount = 1,
                Padding = new Padding(14, 8, 14, 12)
            };
            header.Controls.Add(new Label
            {
                Text = "Özbekça 1.0", ForeColor = Color.White, Font = new Font("Times New Roman", 22f, FontStyle.Bold),
                AutoSize = true, Margin = new Padding(0, 0, 0, 10)
            });
            header.Controls.Add(new Label
            {
                Text = "Oʻ ↔ Ö    Gʻ ↔ Ğ    Sh ↔ Ş    Ch ↔ Ç    ·    Kirill ↔ lotin",
                ForeColor = Color.FromArgb(220, 240, 235), AutoSize = true, Font = new Font("Times New Roman", 12f),
                Margin = new Padding(3, 0, 0, 0)
            });

            var body = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, Padding = new Padding(16, 10, 16, 12) };
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            body.RowStyles.Add(new RowStyle(SizeType.AutoSize));          // tavsif
            body.RowStyles.Add(new RowStyle(SizeType.AutoSize));          // holat
            body.RowStyles.Add(new RowStyle(SizeType.Percent, 100));      // natijalar oynasi
            body.RowStyles.Add(new RowStyle(SizeType.AutoSize));          // jarayon chizig'i
            body.RowStyles.Add(new RowStyle(SizeType.AutoSize));          // tugmalar

            var info = new Label
            {
                Text = "Word, Excel va PowerPoint uçun plagin. Örnatilgaç, ularda «Özbekça» lentasi paydo böladi: " +
                       "yangi ↔ eski lotin, kirill ↔ lotin, istisno sözlar (Wordda yoziş paytida ögiriş ham). " +
                       "Örnatişdan oldin bu dasturlarni yoping.",
                AutoSize = true, Margin = new Padding(0, 0, 0, 6)
            };
            body.Controls.Add(info, 0, 0);
            // matn oyna kengligida qatorlarga bo'linsin
            body.Layout += (s, e) =>
            {
                int w = body.ClientSize.Width - body.Padding.Horizontal - 4;
                if (w > 50 && info.MaximumSize.Width != w) info.MaximumSize = new Size(w, 0);
            };

            _state.AutoSize = true; _state.Margin = new Padding(0, 0, 0, 4);
            _state.Font = new Font(Font, FontStyle.Bold);
            body.Controls.Add(_state, 0, 1);

            _log.Multiline = true; _log.ReadOnly = true; _log.ScrollBars = ScrollBars.Vertical;
            _log.Dock = DockStyle.Fill; _log.Margin = new Padding(0);
            _log.BackColor = Color.FromArgb(246, 248, 248); _log.BorderStyle = BorderStyle.FixedSingle;
            _log.Font = new Font("Times New Roman", 11f);
            body.Controls.Add(_log, 0, 2);

            _bar.Dock = DockStyle.Fill; _bar.Height = 6; _bar.Margin = new Padding(0, 4, 0, 0);
            _bar.Style = ProgressBarStyle.Marquee; _bar.Visible = false;
            body.Controls.Add(_bar, 0, 3);

            var buttons = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 4, Margin = new Padding(0, 12, 0, 0) };
            buttons.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            buttons.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            buttons.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _btnInstall.Size = new Size(150, 34); _btnInstall.Margin = new Padding(0, 0, 8, 0);
            _btnInstall.BackColor = Accent; _btnInstall.ForeColor = Color.White; _btnInstall.FlatStyle = FlatStyle.Flat;
            _btnInstall.Font = new Font(Font, FontStyle.Bold);
            _btnInstall.Click += (s, e) => Run(true);
            _btnRemove.Text = "Öçiriş"; _btnRemove.Size = new Size(110, 34); _btnRemove.Margin = new Padding(0);
            _btnRemove.Click += (s, e) => Run(false);
            _btnClose.Text = "Yopiş"; _btnClose.Size = new Size(100, 34); _btnClose.Margin = new Padding(0);
            _btnClose.Click += (s, e) => Close();
            buttons.Controls.Add(_btnInstall, 0, 0);
            buttons.Controls.Add(_btnRemove, 1, 0);
            buttons.Controls.Add(_btnClose, 3, 0);
            body.Controls.Add(buttons, 0, 4);

            var author = new Label
            {
                Text = "Muallif: Bozorov Obidjon Norqobilovich  ·  o.bozorov@nuu.uz\n" +
                       "Mirzo Uluğbek nomidagi ÖzMU, Axborot xavfsizligi kafedrasi",
                Dock = DockStyle.Bottom, Height = 40, TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.DimGray, BackColor = Color.FromArgb(246, 248, 248), Font = new Font("Times New Roman", 9.5f)
            };
            Controls.Add(body);      // Fill — oxirida qo'shiladi, sarlavha (Top) ustida qoladi
            Controls.Add(header);
            Controls.Add(author);
            AcceptButton = _btnInstall; CancelButton = _btnClose;

            ResumeLayout(false);
            PerformLayout();
            RefreshState();
        }

        void RefreshState()
        {
            var inst = Installer.InstalledApps();
            if (inst.Count > 0)
            {
                _state.Text = "Holat: örnatilgan" + (Installer.InstalledVersion != null ? " (v" + Installer.InstalledVersion + ")" : "") +
                              " — " + string.Join(", ", inst);
                _state.ForeColor = Accent;
                _btnInstall.Text = "Yangilaş";
                _btnRemove.Enabled = true;
            }
            else
            {
                var av = Installer.AvailableApps();
                _state.Text = av.Count == 0 ? "Holat: Microsoft Office topilmadi" : "Holat: örnatilmagan (" + string.Join(", ", av) + " topildi)";
                _state.ForeColor = Color.DimGray;
                _btnInstall.Text = "Örnatiş";
                _btnRemove.Enabled = false;
            }
        }

        void Log(string s)
        {
            if (InvokeRequired) { BeginInvoke((Action<string>)Log, s); return; }
            _log.AppendText(s + Environment.NewLine);
        }

        void Run(bool install)
        {
            List<string> running;
            while ((running = Installer.RunningApps()).Count > 0)
            {
                var r = MessageBox.Show(this, string.Join(", ", running) + " oçiq. Işingizni saqlab, ularning barça oynalarini yoping va «Qayta uriniş» ni bosing.",
                    Text, MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning);
                if (r != DialogResult.Retry) return;
            }
            if (!install && MessageBox.Show(this, "Plagin Word, Excel va PowerPoint dan olib taşlansinmi?", Text,
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;

            SetBusy(true);
            _log.Clear();
            // Word bilan ishlash STA oqimida bo'lishi kerak
            var t = new Thread(() =>
            {
                string error = null;
                List<string> done = null;
                try
                {
                    if (install) done = Installer.Install(Log);
                    else Installer.Uninstall(Log);
                }
                catch (InstallerException ex) { error = ex.Message; }
                catch (Exception ex) { error = "Kutilmagan xato: " + ex.Message; }
                BeginInvoke((Action)(() =>
                {
                    SetBusy(false);
                    RefreshState();
                    if (error != null)
                    {
                        Log("XATO: " + error);
                        MessageBox.Show(this, error, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else if (install)
                    {
                        string msg = "Plagin örnatildi: " + string.Join(", ", done) + ".\n\n" +
                                     "Dasturni oçing — yuqorida «Özbekça» lentasi paydo böladi.\n";
                        if (done.Contains("Word"))
                            msg += "\nWord tezkor tugmalari: " + K("Lotin") + " (yangi alifboga), " + K("Eski") + " (eskichaga), " +
                                   K("LotinKirill") + " (kirillga),\n" +
                                   K("Yozishda") + " (yozişda), " + K("Istisno") + " (sözni qöşiş), " + K("Istisnolar") + " (istisno sözlar).\n";
                        if (done.Contains("Excel"))
                            msg += "\nExcel: Alt+A, Alt+E, Alt+C, Alt+I, Alt+O.\n";
                        if (done.Contains("PowerPoint"))
                            msg += "\nPowerPoint: lentadagi tugmalar (PowerPoint tezkor tugmaga ruxsat bermaydi).\n";
                        MessageBox.Show(this, msg, Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }));
            });
            t.SetApartmentState(ApartmentState.STA);
            t.IsBackground = true;
            t.Start();
        }

        static string K(string id) => Installer.ChosenKeys.TryGetValue(id, out var k) ? k : "—";

        void SetBusy(bool on)
        {
            _bar.Visible = on;
            _btnInstall.Enabled = !on;
            _btnRemove.Enabled = !on && Installer.InstalledApps().Count > 0;
            _btnClose.Enabled = !on;
            UseWaitCursor = on;
        }
    }

    internal static class Program
    {
        [DllImport("kernel32.dll")] static extern bool AttachConsole(int pid);

        [STAThread]
        static int Main(string[] args)
        {
            var a = args.Select(x => x.ToLowerInvariant()).ToList();
            if (a.Contains("/s") || a.Contains("/u"))
            {
                AttachConsole(-1);   // buyruq qatoridan ishga tushirilganda xabarlar ko'rinsin
                try
                {
                    if (a.Contains("/u")) Installer.Uninstall(Console.WriteLine);
                    else Installer.Install(Console.WriteLine);
                    return 0;
                }
                catch (Exception ex) { Console.WriteLine("XATO: " + ex.Message); return 1; }
            }
            // SetProcessDPIAware chaqirilmaydi — oyna masshtabdan mustaqil bo'lsin (yuqoriga qarang)
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new SetupForm());
            return 0;
        }
    }
}
