/*
 * Yangi o'zbek alifbosi konvertori (umumiy yadro)
 * Oʻ→Ö, Gʻ→Ğ, Sh→Ş, Ch→Ç  (Ng o'zgarmaydi — yangi alifboda n+g deb yoziladi)
 *
 * Bu fayl Office Add-in ichida ishlatiladi; VBA va VSTO versiyalari
 * aynan shu algoritmning ko'chirmasi. Testlar: tests/converter.test.js
 */
(function (root) {
  "use strict";

  // Apostrof o'rnida ishlatiladigan barcha belgilar
  // (kengaytirilgan ro'yxat: klaviatura, PDF va OCR dan keladigan o'xshash belgilar)
  var APOS = "'‘’ʻʼ`´′ʹʽʾʿˈˊՙ՚׳´῾‛ꞌ";

  // Lotin so'z ichiga adashib kirib qolgan kirill harflari (ko'rinishi bir xil): Tоshkent -> Toshkent
  var HOMO = { "а": "a", "е": "e", "о": "o", "р": "p", "с": "c", "х": "x", "у": "y", "к": "k",
    "А": "A", "В": "B", "Е": "E", "К": "K", "М": "M", "Н": "H", "О": "O", "Р": "P", "С": "C", "Т": "T", "Х": "X" };

  /** So'zda lotin harfi va faqat "o'xshash" kirill harflari bo'lsa — tuzatilgan shakl, aks holda null. */
  function fixHomoglyphs(tok) {
    var hasLat = false, hasHomo = false, i, c;
    for (i = 0; i < tok.length; i++) {
      c = tok.charAt(i);
      if (/[A-Za-z]/.test(c)) hasLat = true;
      else if (HOMO.hasOwnProperty(c)) hasHomo = true;
      else if (/[\u0400-\u04FF]/.test(c)) return null;   // haqiqiy kirill harfi bor — kirill so'z
    }
    if (!hasLat || !hasHomo) return null;
    var out = "";
    for (i = 0; i < tok.length; i++) { c = tok.charAt(i); out += HOMO.hasOwnProperty(c) ? HOMO[c] : c; }
    return out;
  }
  var TUTUQ = "ʼ"; // ʼ

  // Inglizcha abzasni aniqlash uchun so'zlar (o'zbekchada uchramaydi)
  var EN_WORDS = ["the", "and", "of", "are", "for", "with", "this", "that", "from",
    "which", "was", "were", "have", "has", "been", "these", "their", "into", "its"];

  // Standart istisnolar. Katta harf bor bo'lsa — registrga sezgir.
  // "*" — shu bilan boshlanadigan har qanday so'z; "#" — keyin raqam/-/ kelishi mumkin.
  var DEFAULT_EXCEPTIONS = [
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
    "O'Toole", "O'Higgins", "o'clock", "g'day"
  ];

  // Kirilldan o'girish istisnolari: "so'z" — o'girilmaydi, "so'z*" — shu bilan boshlanganlar,
  // "o'zak=Lotin" — maxsus shakl (qo'shimchalari odatdagidek o'giriladi: Майкрософтнинг → Microsoftning)
  var DEFAULT_CYR_EXCEPTIONS = [
    "Майкрософт=Microsoft", "Гугл=Google", "Виндоус=Windows", "Фейсбук=Facebook", "Ютуб=YouTube",
    "Эппл=Apple", "Хуавей=Huawei", "Линукс=Linux", "Айфон=iPhone", "Пауэрпойнт=PowerPoint",
    "Эксель=Excel", "Ворд=Word", "Ватсап=WhatsApp", "Твиттер=Twitter", "Скайп=Skype", "абзац=abzas",
    "авианосец=avianoses", "броненосец=bronenoses", "миноносец=minonoses", "марганец=marganes",
    "ранец=ranes", "коэффициент=koeffitsient", "целлюлоид=selluloid", "целлюлоза=selluloza",
    "октябр=oktabr", "сентябр=sentabr", "партьер=parter", "тюлен=tulen"
  ];

  var MONO_FONTS = ["courier new", "courier", "consolas", "lucida console", "cascadia code",
    "cascadia mono", "source code pro", "menlo", "monaco", "jetbrains mono", "fira code"];

  function isApos(c) { return c !== undefined && c !== "" && APOS.indexOf(c) >= 0; }
  function isLetter(c) { return c !== undefined && c !== "" && c.toLowerCase() !== c.toUpperCase(); }
  function isSpace(c) { return /\s/.test(c); }
  function isLower(c) { return c === c.toLowerCase(); }

  function matchException(word, entry) {
    if (!entry) return false;
    var cs = entry.toLowerCase() !== entry; // katta harf bor -> registrga sezgir
    var a = cs ? word : word.toLowerCase();
    var e = cs ? entry : entry.toLowerCase();
    var last = e.charAt(e.length - 1);
    if (last === "*") return a.indexOf(e.slice(0, -1)) === 0;
    if (last === "#") {
      var p = e.slice(0, -1);
      if (a.indexOf(p) !== 0) return false;
      return /^[0-9\-\/]*$/.test(a.slice(p.length));
    }
    return a === e;
  }

  // So'zni himoya qilayotgan istisno yozuvini qaytaradi (yoki null)
  function findException(core, exceptions) {
    if (!core) return null;
    var parts = [core].concat(core.split(/[-\/]/));
    for (var i = 0; i < parts.length; i++) {
      for (var j = 0; j < exceptions.length; j++) {
        if (matchException(parts[i], exceptions[j])) return exceptions[j];
      }
    }
    return null;
  }
  function isExceptionWord(core, exceptions) { return findException(core, exceptions) !== null; }

  /**
   * Foydalanuvchi qatorlarini standart ro'yxat bilan birlashtiradi.
   *   so'z    -> istisno qo'shiladi
   *   !so'z   -> shu nomdagi standart istisno o'chiriladi
   *   ; izoh  -> e'tiborsiz
   */
  function mergeExceptions(userLines, defaults) {
    defaults = defaults || DEFAULT_EXCEPTIONS;
    var keyOf = function (e) { return String(e).split("=")[0].trim().toLowerCase(); };
    var off = {}, extra = [];
    (userLines || []).forEach(function (raw) {
      var ln = String(raw).trim();
      if (!ln || ln.charAt(0) === ";") return;
      if (ln.charAt(0) === "!") { off[keyOf(ln.slice(1))] = true; return; }
      if (extra.indexOf(ln) < 0) extra.push(ln);
    });
    // foydalanuvchi yozuvi standartdagi shu kalitni almashtiradi
    var userKeys = {};
    extra.forEach(function (e) { userKeys[keyOf(e)] = true; });
    return defaults.filter(function (e) { return !off[keyOf(e)] && !userKeys[keyOf(e)]; })
      .concat(extra);
  }
  function mergeCyrExceptions(userLines) { return mergeExceptions(userLines, DEFAULT_CYR_EXCEPTIONS); }

  // "so'z", "so'z*", "o'zak=Lotin" -> kalit bo'yicha jadval (keyingi yozuv avvalgisini almashtiradi)
  function parseCyrExceptions(list) {
    var map = {};
    (list || []).forEach(function (raw) {
      var e = String(raw).trim(), eq = e.indexOf("="), r;
      if (eq > 0) r = { key: e.slice(0, eq).trim().toLowerCase(), prefix: true, latin: e.slice(eq + 1).trim() };
      else if (e.charAt(e.length - 1) === "*") r = { key: e.slice(0, -1).toLowerCase(), prefix: true, latin: null };
      else r = { key: e.toLowerCase(), prefix: false, latin: null };
      if (r.key) map[r.key] = r;
    });
    return map;
  }
  // Eng uzun mos kalit (so'z boshidan)
  function findCyrRule(run, map) {
    var rl = run.toLowerCase();
    for (var L = rl.length; L > 0; L--) {
      var r = map[rl.slice(0, L)];
      if (r && (r.prefix || L === rl.length)) return r;
    }
    return null;
  }

  /**
   * O'girilgan so'zni eski yozuvga qaytaradi (istisnoga qo'shishda kerak):
   * Şamir -> Shamir, ŞAHAR -> SHAHAR, Özbek -> Oʻzbek
   */
  function revertWord(w) {
    w = String(w);
    var letters = w.replace(/[^\p{L}]/gu, "");
    var allCaps = letters.length > 1 && letters === letters.toUpperCase();
    var map = { "\u015F": "sh", "\u00E7": "ch", "\u00F6": "o\u02BB", "\u011F": "g\u02BB",
                "\u015E": allCaps ? "SH" : "Sh", "\u00C7": allCaps ? "CH" : "Ch",
                "\u00D6": "O\u02BB", "\u011E": "G\u02BB" };
    var out = "";
    for (var i = 0; i < w.length; i++) out += map[w.charAt(i)] || w.charAt(i);
    return out;
  }

  /** Bitta so'z nima bo'lishini va nega shundayligini tushuntiradi. */
  function explainWord(word, opts) {
    opts = opts || {};
    var exceptions = opts.exceptions || DEFAULT_EXCEPTIONS;
    var tok = String(word).trim();
    var core = coreOf(tok);
    var res = convertText(tok, { tutuq: opts.tutuq, skipEnglish: false, exceptions: exceptions });
    var reason = null;
    if (looksLikeLinkOrCode(tok)) reason = "havola yoki fayl nomi";
    else if (findException(core, exceptions)) reason = "istisno: " + findException(core, exceptions);
    else if (looksForeign(core)) reason = "chet so'z (w yoki yolg'iz c harfi bor)";
    return { input: tok, output: res.text, changed: res.text !== tok, reason: reason };
  }

  // O'zbek lotin alifbosida yo'q harflar: w, yolg'iz c (ch emas) -> chet so'z
  function looksForeign(core) {
    for (var i = 0; i < core.length; i++) {
      var c = core.charAt(i);
      if (c === "w" || c === "W") return true;
      if ((c === "c" || c === "C")) {
        var n = core.charAt(i + 1);
        if (n !== "h" && n !== "H") return true;
      }
    }
    return false;
  }

  function looksLikeLinkOrCode(tok) {
    if (tok.indexOf("://") >= 0) return true;
    if (/^www\./i.test(tok)) return true;
    if (tok.indexOf("@") >= 0) return true;
    if (tok.charAt(0) === "#") return true;
    if (tok.indexOf("\\") >= 0) return true;
    if (/^[\w.\-]+\.(com|uz|org|net|ru|io|edu|gov|docx?|pdf|xlsx?|pptx?|exe|py|js|html?|txt|zip)$/i.test(tok)) return true;
    return false;
  }

  function isEnglishText(text) {
    var words = text.toLowerCase().split(/[^a-zÀ-ɏЀ-ӿ']+/).filter(Boolean);
    if (words.length < 3) return false;
    var hits = 0;
    for (var i = 0; i < words.length; i++) if (EN_WORDS.indexOf(words[i]) >= 0) hits++;
    return hits >= 2 && hits / words.length >= 0.08;
  }

  function isMonoFont(name) {
    return !!name && MONO_FONTS.indexOf(String(name).toLowerCase()) >= 0;
  }

  // Tokendan old/oxiridagi tinish belgilarini olib tashlash (apostroflar ham)
  function coreOf(tok) {
    var s = 0, e = tok.length;
    while (s < e && !isLetter(tok.charAt(s)) && !/[0-9]/.test(tok.charAt(s))) s++;
    while (e > s && !isLetter(tok.charAt(e - 1)) && !/[0-9]/.test(tok.charAt(e - 1))) e--;
    return tok.slice(s, e);
  }

  /**
   * Abzas matnini tahlil qilib, o'zgarishlar ro'yxatini qaytaradi.
   * @returns {{edits: Array<{pos:number,len:number,from:string,to:string}>, text: string, skipped: string|null}}
   */
  function convertText(text, opts) {
    opts = opts || {};
    var tutuq = !!opts.tutuq;
    var skipEnglish = opts.skipEnglish !== false;
    var exceptions = opts.exceptions || DEFAULT_EXCEPTIONS;
    var edits = [];

    if (skipEnglish && isEnglishText(text)) return { edits: [], text: text, skipped: "english" };

    var quoteOpen = false;
    var i = 0, n = text.length;
    while (i < n) {
      if (isSpace(text.charAt(i))) { i++; continue; }
      var start = i;
      while (i < n && !isSpace(text.charAt(i))) i++;
      var tok = text.slice(start, i);
      quoteOpen = processToken(tok, start, quoteOpen, edits, tutuq, exceptions);
    }

    // Natijaviy matn
    var out = text;
    for (var k = edits.length - 1; k >= 0; k--) {
      var ed = edits[k];
      out = out.slice(0, ed.pos) + ed.to + out.slice(ed.pos + ed.len);
    }
    return { edits: edits, text: out, skipped: null };
  }

  function processToken(tok, base, quoteOpen, edits, tutuq, exceptions) {
    var orig = tok, fixed = fixHomoglyphs(tok), local = [];
    if (fixed) tok = fixed;
    var len = tok.length;

    // Tokenning boshidagi tinish belgilari; apostrof bo'lsa — qo'shtirnoq ochilishi
    var lead = 0;
    while (lead < len && !isLetter(tok.charAt(lead)) && !/[0-9]/.test(tok.charAt(lead))) {
      if (isApos(tok.charAt(lead))) quoteOpen = true;
      lead++;
    }
    if (lead >= len) return quoteOpen;

    // Oxiridagi tinish belgilari (apostrofdan tashqari)
    var limit = len;
    var t = len;
    while (t > lead && !isLetter(tok.charAt(t - 1)) && !isApos(tok.charAt(t - 1)) && !/[0-9]/.test(tok.charAt(t - 1))) t--;
    // t — oxirgi apostrof/harf/raqamdan keyingi pozitsiya
    if (t > lead && isApos(tok.charAt(t - 1))) {
      var aposRun = 0;
      while (t - 1 - aposRun > lead && isApos(tok.charAt(t - 1 - aposRun))) aposRun++;
      if (aposRun >= 2 && !quoteOpen) {
        // masalan  boo''  -> ataylab yozilgan, oxiriga tegmaymiz
        limit = t - aposRun - 1;
      } else if (quoteOpen) {
        // ‘... tog’’  yoki  ‘kino’  -> oxirgisi yopuvchi qo'shtirnoq
        limit = t - 1; // yopuvchi qo'shtirnoq
        quoteOpen = false;
      }
    }

    var core = coreOf(tok);
    var skip = looksLikeLinkOrCode(tok) || isExceptionWord(core, exceptions) || looksForeign(core);

    var i = skip ? limit : lead;
    while (i < limit) {
      var c = tok.charAt(i);
      var nx = i + 1 < limit ? tok.charAt(i + 1) : "";

      if ((c === "o" || c === "O" || c === "g" || c === "G") && isApos(nx)) {
        var nx2 = i + 2 < limit ? tok.charAt(i + 2) : "";
        if (isApos(nx2)) { i += 3; continue; } // o''  -> ataylab yozilgan, tegmaymiz
        var rep = c === "o" ? "ö" : c === "O" ? "Ö" : c === "g" ? "ğ" : "Ğ";
        local.push({ pos: base + i, len: 2, from: orig.substr(i, 2), to: rep });
        i += 2; continue;
      }
      if ((c === "s" || c === "S" || c === "c" || c === "C") && (nx === "h" || nx === "H")) {
        var low = isLower(c);
        var r = (c === "s" || c === "S") ? (low ? "ş" : "Ş") : (low ? "ç" : "Ç");
        local.push({ pos: base + i, len: 2, from: orig.substr(i, 2), to: r });
        i += 2; continue;
      }
      if (tutuq && isApos(c) && c !== TUTUQ && i > lead && isLetter(tok.charAt(i - 1)) && isLetter(nx)) {
        local.push({ pos: base + i, len: 1, from: c, to: TUTUQ });
        i += 1; continue;
      }
      i++;
    }
    if (fixed) {
      // o'girilmagan joylardagi o'xshash kirill harflari — bittadan almashtiriladi
      for (var h = 0; h < len; h++) {
        if (orig.charAt(h) === tok.charAt(h)) continue;
        var covered = false;
        for (var q = 0; q < local.length; q++) {
          if (base + h >= local[q].pos && base + h < local[q].pos + local[q].len) { covered = true; break; }
        }
        if (!covered) local.push({ pos: base + h, len: 1, from: orig.charAt(h), to: tok.charAt(h) });
      }
      local.sort(function (a, b) { return a.pos - b.pos; });
    }
    for (var e = 0; e < local.length; e++) edits.push(local[e]);
    return quoteOpen;
  }

  // ======================================================== KIRILL → YANGI LOTIN
  var CYR_MAP = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "ж": "j",
    "з": "z", "и": "i", "й": "y", "к": "k", "л": "l", "м": "m",
    "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t",
    "у": "u", "ф": "f", "х": "x", "ч": "ç", "ш": "ş",
    "щ": "ş", "ы": "i", "э": "e", "ё": "yo", "ю": "yu", "я": "ya",
    "ў": "ö", "қ": "q", "ғ": "ğ", "ҳ": "h"
  };
  var CYR_VOWELS = "аеёиоуэюяўы"; // а е ё и о у э ю я ў ы
  var RU_WORDS = ["в", "на", "не", "что", "для",
    "это", "как", "при", "или", "по",
    "же", "был", "была", "были",
    "также", "который",
    "которые", "которая",
    "его", "с", "и", "к"]; // в на не что для это как при или по же был была были также который которые которая его с и к

  function isCyr(c) { return c >= "Ѐ" && c <= "ӿ"; }

  // O'zbekcha harflar (ў қ ғ ҳ) bo'lmasa va ruscha yordamchi so'zlar ko'p bo'lsa — ruscha abzas
  function isRussianText(text) {
    if (/[ўқғҳЎҚҒҲ]/.test(text)) return false;
    var words = text.toLowerCase().split(/[^Ѐ-ӿ]+/).filter(Boolean);
    if (words.length < 3) return false;
    var hits = 0;
    for (var i = 0; i < words.length; i++) if (RU_WORDS.indexOf(words[i]) >= 0) hits++;
    return hits >= 2 && hits / words.length >= 0.08;
  }

  function cyrChar(text, i, runStart) {
    var c = text.charAt(i), lc = c.toLowerCase();
    var prev = i > runStart ? text.charAt(i - 1).toLowerCase() : "";
    var next = text.charAt(i + 1).toLowerCase();
    switch (lc) {
      case "е": // е
        return (i === runStart || CYR_VOWELS.indexOf(prev) >= 0 || prev === "ъ" || prev === "ь") ? "ye" : "e";
      case "ц": // ц
        return (i === runStart || CYR_VOWELS.indexOf(prev) < 0) ? "s" : "ts";
      case "ъ": // ъ
        return "еёюя".indexOf(next) >= 0 ? "" : TUTUQ;
      case "ь": // ь
        return next === "о" ? "y" : "";
      case "с": // с + ҳ -> sʼh (aks holda "sh" ş bilan adashadi)
        return next === "ҳ" ? "s" + TUTUQ : "s";
    }
    return CYR_MAP[lc] !== undefined ? CYR_MAP[lc] : c;
  }

  /** Kirill matnini yangi lotin alifbosiga o'giradi. */
  function cyrToLatin(text, opts) {
    opts = opts || {};
    var edits = [];
    if (opts.skipRussian !== false && isRussianText(text)) return { edits: [], text: text, skipped: "russian" };
    var rules = parseCyrExceptions(opts.cyrExceptions || DEFAULT_CYR_EXCEPTIONS);
    var n = text.length, i = 0;
    while (i < n) {
      if (!isCyr(text.charAt(i))) { i++; continue; }
      var start = i;
      while (i < n && isCyr(text.charAt(i))) i++;
      var run = text.slice(start, i);
      var allCaps = run.length > 1 && run === run.toUpperCase() && run !== run.toLowerCase();
      var rule = findCyrRule(run, rules);
      if (rule && rule.latin === null) continue;            // istisno: o'girilmaydi
      var out = "", from = start;
      if (rule) {
        out = rule.latin;
        if (allCaps) out = out.toUpperCase();
        else if (run.charAt(0) !== run.charAt(0).toLowerCase() && out === out.toLowerCase())
          out = out.charAt(0).toUpperCase() + out.slice(1);       // Абзац -> Abzas
        from = start + rule.key.length;
      }
      for (var k = from; k < i; k++) {
        var c = text.charAt(k);
        var t = cyrChar(text, k, start);
        if (t && c !== c.toLowerCase()) {
          t = allCaps ? t.toUpperCase() : t.charAt(0).toUpperCase() + t.slice(1);
        }
        out += t;
      }
      if (out !== run) edits.push({ pos: start, len: run.length, from: run, to: out });
    }
    var res = text;
    for (var e = edits.length - 1; e >= 0; e--) {
      res = res.slice(0, edits[e].pos) + edits[e].to + res.slice(edits[e].pos + edits[e].len);
    }
    return { edits: edits, text: res, skipped: null };
  }

  // ======================================================== LOTIN → KIRILL
  // Yangi alifbo (ö ğ ş ç) ham, amaldagi lotin (oʻ gʻ sh ch) ham qabul qilinadi.
  // Lug'at: "lotin=кирилл" (eski lotin shaklida, kichik harf). 5+ harfli kalit — o'zak sifatida
  // (qo'shimchalari bilan: oktabrda -> октябрда), qisqasi — faqat aniq so'z.
  var DEFAULT_LAT_CYR = [
    "abzas=абзац", "aksent=акцент", "avianoses=авианосец", "batsilla=бацилла", "biomitsin=биомицин",
    "botsman=боцман", "bronenoses=броненосец", "brutselloz=бруцеллоз", "cherepitsa=черепица",
    "dotsent=доцент", "fransuz=француз", "gaubitsa=гаубица", "gers=герц", "glitserin=глицерин",
    "gorchitsa=горчица", "gusenitsa=гусеница", "inersiya=инерция", "inssenirovka=инсценировка",
    "kalsiy=кальций", "kansler=канцлер", "koeffitsient=коэффициент", "konferens-zal=конференц-зал",
    "konsepsiya=концепция", "konsern=концерн", "konsert=концерт", "konslager=концлагерь",
    "kultivatsiya=культивация", "kvars=кварц", "litsey=лицей", "lotsman=лоцман", "marganes=марганец",
    "minonoses=миноносец", "munitsipalitet=муниципалитет", "ofitsiant=официант",
    "penitsillin=пенициллин", "platsdarm=плацдарм", "platskarta=плацкарта",
    "politsmeyster=полицмейстер", "ranes=ранец", "ritsar=рицарь", "sellofan=целлофан",
    "selluloid=целлюлоид", "selluloza=целлюлоза", "selsiy=цельсий", "sement=цемент", "sentner=центнер",
    "sentrifuga=центрифуга", "senz=ценз", "senzor=цензор", "senzura=цензура", "sex=цех",
    "shnitsel=шницель", "shveysar=швейцар", "siferblat=циферблат", "silindr=цилиндр",
    "silindrik=цилиндрик", "sink=цинк", "sirk=цирк", "sirkulyar=циркуляр", "sisterna=цистерна",
    "sitrus=цитрус", "sotsiologik=социологик", "sotsiologiya=социология", "ssenariy=сценарий",
    "kultivator=культиватор", "kuryer=курьер", "lager=лагерь", "losos=лосось", "mebel=мебель",
    "medal=медаль", "medalyon=медальон", "menshevik=меньшевик", "menshevizm=меньшевизм",
    "migren=мигрень", "mikrofilm=микрофильм", "mil=миль", "model=модель", "neft=нефть", "nikel=никель",
    "nippel=ниппель", "nol=ноль", "noyabr=ноябрь", "oktabr=октябрь", "palto=пальто", "panel=панель",
    "parallel=параллель", "parol=пароль", "parter=партьер", "patrul=патруль", "pavilyon=павильон",
    "pedal=педаль", "plastir=пластирь", "pochtalyon=почтальон", "porshen=поршень", "portfel=портфель",
    "povest=повесть", "predoxranitel=предохранитель", "premyera=премьера", "pristan=пристань",
    "puls=пульс", "pyesa=пьеса", "rels=рельс", "relyef=рельеф", "rentabel=рентабель", "rezba=резьба",
    "rol=роль", "royal=рояль", "rul=руль", "seld=сельд", "sentabr=сентябрь", "shinel=шинель",
    "shpatel=шпатель", "shpilka=шпилька", "shpindel=шпиндель", "shtapel=штапель", "shtempel=штемпель",
    "shtepsel=штепсель", "spektakl=спектакль", "spiral=спираль", "sterjen=стержень", "sudya=судья",
    "sulfat=сульфат", "tabel=табель", "tekstil=текстиль", "tokar=токарь", "tulen=тюлень",
    "tunnel=туннель", "umivalnik=умивальник", "valeryanka=валерьянка", "vals=вальс", "veksel=вексель",
    "velvet=вельвет", "ventil=вентиль", "vermishel=вермишель", "vimpel=вимпель",
    "violonchel=виолончель", "volfram=вольфрам", "volt=вольт", "volta=вольта", "voltmetr=вольтметр",
    "vulgar=вульгар", "yakor=якорь", "yanvar=январь", "yuriskonsult=юрисконсульт", "ansambl=ансамбль",
    "artel=артель", "artikl=артикль", "aryergard=арьергард", "asfalt=асфальт", "atelye=ателье",
    "avtomobil=автомобиль", "balzam=бальзам", "banderol=бандероль", "batalyon=батальон",
    "bilyard=бильярд", "binokl=бинокль", "bolshevik=большевик", "budilnik=будильник", "bulvar=бульвар",
    "dalton=дальтон", "dekabr=декабрь", "delfin=дельфин", "devalvatsiya=девальвация",
    "dirijabl=дирижабль", "dizel=дизель", "dizel-motor=дизель-мотор", "duel=дуэль",
    "dvigatel=двигатель", "emulsiya=эмульсия", "eskadrilya=эскадрилья", "fakultativ=факультатив",
    "fakultet=факультет", "falsifikator=фальсификатор", "falsifikatsiya=фальсификация",
    "feldmarshal=фельдмаршал", "feldsher=фельдшер", "festival=фестиваль", "fevral=февраль",
    "filtr=фильтр", "folklor=фольклор", "fotoalbom=фотоальбом", "fotoatelye=фотоателье",
    "gantel=гантель", "gastrol=гастроль", "gilza=гильза", "gospital=госпиталь", "grifel=грифель",
    "impuls=импульс", "insult=инсульт", "intervyu=интервью", "inyeksiya=инъекция", "italyan=итальян",
    "iyul=июль", "iyun=июнь", "kabel=кабель", "kalendar=календарь", "kanifol=канифоль",
    "karamel=карамель", "kartel=картель", "karusel=карусель", "katapulta=катапульта",
    "kinofestival=кинофестиваль", "kinofilm=кинофильм", "kisel=кисель", "kitel=китель",
    "kobalt=кобальт", "kompanyon=компаньон", "konferansye=конферансье", "syomka=съёмка",
    "korrupsiya=коррупция", "konyunktura=конъюнктура", "adyunktura=адъюнктура", "ishoq=исҳоқ",
    "mushaf=мусҳаф", "ashob=асҳоб", "ishol=исҳол", "mayor=майор", "rayon=район",
    "mikrorayon=микрорайон", "mayonez=майонез", "yogurt=йогурт", "yoga*=йога", "poyabzal=пойабзал",
    "poyandoz=пойандоз", "poyafzal=пойафзал", "poyustun=пойустун", "moyupa=мойупа",
    "konveyer=конвейер", "pleyer=плейер", "stayer=стайер", "foye*=фойе", "kompyuter=компьютер",
    "bulyon=бульон", "aprel=апрель", "film*=фильм", "rubl*=рубль", "prinsip=принцип", "sitata=цитата",
    "sivilizatsiya=цивилизация", "sikl*=цикл", "siklon=циклон", "sentr*=центр", "konsentr=концентр",
    "spetsif=специф", "spetsial=специал", "patsient=пациент", "ofitsial=официал", "stansiya=станция",
    "instansiya=инстанция", "distansiya=дистанция", "kvitansiya=квитанция", "tendensiya=тенденция",
    "potensial=потенциал", "litsenziya=лицензия", "absorbsiya=абсорбция", "abssess=абсцесс",
    "adsorbsiya=адсорбция", "adyunkt=адъюнкт", "aerofotosyomka=аэрофотосъёмка",
    "aerosyomka=аэросъёмка", "agnostitsizm=агностицизм", "aksioma=аксиома", "aksiz=акциз",
    "akvarel=акварель", "albatros=альбатрос", "albom=альбом", "alfa-zarralar=альфа-зарралар",
    "alfa=альфа", "alkogol=алкоголь", "alma-terapiya=альма-терапия", "altchi=альтчи",
    "altimetr=альтиметр", "alyans=альянс", "amalgama=амальгама", "anneksiya=аннексия",
    "apelsin=апельсин", "apopleksiya=апоплексия", "asfiksiya=асфиксия", "aviamodel=авиамодель",
    "avtomagistral=автомагистраль", "balneolog=бальнеолог", "balneologik=бальнеологик",
    "balneologiya=бальнеология", "balneoterapiya=бальнеотерапия", "balneotexnika=бальнеотехника",
    "barelyef=барельеф", "barrel=баррель", "baryer=барьер", "bazalt=базальт", "beletaj=бельэтаж",
    "belgiyalik=бельгиялик", "belting=бельтинг", "belveder=бельведер", "bemeyor=бемеъёр",
    "biofiltr=биофильтр", "biolyuminessensiya=биолюминесценция", "bolonya=болонья",
    "bolshevizm=большевизм", "brakonyerlik=браконьерлик", "broneavtomobil=бронеавтомобиль",
    "buldenej=бульденеж", "buldog=бульдог", "buldozer=бульдозер", "buldozerchi=бульдозерчи",
    "burgutsifat=бургутсифат", "burgutsimon=бургутсимон", "byulleten=бюллетень", "chizel=чизель",
    "daraxtsimon=дарахтсимон", "delta=дельта", "detal=деталь", "diafilm=диафильм",
    "differensiatsiya=дифференциация", "distansion=дистанцион", "drel=дрель", "emal=эмаль",
    "empiriokrititsizm=эмпириокритицизм", "endshpil=эндшпиль", "ensefalit=энцефалит",
    "ensefalogramma=энцефалограмма", "ensiklopedik=энциклопедик", "ensiklopedist=энциклопедист",
    "ensiklopediya=энциклопедия", "ensiklopediyachi=энциклопедиячи", "esminets=эсминец",
    "feldyeger=фельдъегерь", "felyeton=фельетон", "felyetonchi=фельетончи", "filmoy=филмой",
    "fitonsid=фитонцид", "fizkultura=физкультура", "fizkulturachi=физкультурачи", "fleksiya=флексия",
    "folga=фольга", "fonar=фонарь", "fortepyano=фортепьяно", "fotosyomka=фотосъёмка",
    "fransiya=франция", "frits=фриц", "gallitsizm=галлицизм", "gallyutsinatsiya=галлюцинация",
    "galvanik=гальваник", "galvanizatsiya=гальванизация", "galvanometr=гальванометр", "gavan=гавань",
    "gelmint=гельминт", "gelmintologiya=гельминтология", "gelmintoz=гельминтоз",
    "geraldika=геральдика", "gersog=герцог", "gidropodyomnik=гидроподъёмник", "giposulfit=гипосульфит",
    "golf=гольф", "gorelyef=горельеф", "guash=гуашь", "instarsiya=инстарция", "interyer=интерьер",
    "inventar=инвентарь", "itsigek=итсигек", "jannatsimon=жаннатсимон", "jenshen=женьшень",
    "kafel=кафель", "kalka=калька", "kalkalamoq=калькаламоқ", "kalkulyatsiya=калькуляция",
    "kanserogen=канцероген", "kapelmeyster=капельмейстер", "kapsyul=капсюль", "kartech=картечь",
    "karyer=карьер", "kastryul=кастрюль", "kastryulka=кастрюлька", "kaxeksiya=кахексия",
    "kinokonsert=киноконцерт", "kinossenariy=киносценарий", "klassitsizm=классицизм", "knyaz=князь",
    "kokteyl=коктейль", "kolonsifra=колонцифра", "kommersiya=коммерция", "konki=коньки",
    "konkichi=конькичи", "konsessiya=концессия", "konsorsium=консорциум", "konsultant=консультант",
    "konsultativ=консультатив", "konsultatsiya=консультация", "kontrol=контроль", "konyak=коньяк",
    "korol=король", "krovat=кровать", "kulminatsion=кульминацион", "kulminatsiya=кульминация",
    "kulturizm=культуризм", "kyat=кьят", "lanset=ланцет", "latun=латунь", "losyon=лосьон",
    "lourensiy=лоуренсий", "marselyeza=марсельеза", "matritsa=матрица", "mayub=маъюб", "mayus=маъюс",
    "mayuslanmoq=маъюсланмоқ", "mayuslik=маъюслик", "mayusona=маъюсона", "meditsina=медицина",
    "mehnatsevar=меҳнатсевар", "mehnatsevarlik=меҳнатсеварлик", "melxior=мельхиор",
    "mexanitsizm=механицизм", "meyor=меъёр", "meyoriy=меъёрий", "mikroinsult=микроинсульт",
    "mistitsizm=мистицизм", "mizanssena=мизансцена", "modelyer=модельер", "molbert=мольберт",
    "monastir=монастирь", "motel=мотель", "multimediya=мультимедия", "multimillioner=мультимиллионер",
    "multiplikator=мультипликатор", "multiplikatsion=мультипликацион",
    "multiplikatsiya=мультипликация", "nenets=ненец", "nenetslar=ненецлар", "nimpalto=нимпальто",
    "oblast=область", "ochered=очередь", "oftalmologiya=офтальмология", "okkultizm=оккультизм",
    "opsion=опцион", "ordinarets=ординарец", "otel=отель", "palma=пальма", "palmazor=пальмазор",
    "palpatsiya=пальпация", "pechat=печать", "pechenye=печенье", "penalti=пенальти",
    "petlitsa=петлица", "pinset=пинцет", "pitssa=пицца", "platskart=плацкарт", "plebissit=плебисцит",
    "porsillamoq=порсилламоқ", "porsilloq=порсиллоқ", "porsin=порсин", "portsigar=портсигар",
    "premyer-ministr=премьер-министр", "press-papye=пресс-папье", "press-sekretar=пресс-секретарь",
    "profil=профиль", "provinsializm=провинциализм", "provinsiya=провинция",
    "pulmonologiya=пульмонология", "pulpit=пульпит", "pult=пульт", "pulverizator=пульверизатор",
    "qaynegachi=қайнэгачи", "qimmatsiramoq=қимматсирамоқ", "radiospektakl=радиоспектакль",
    "rantye=рантье", "raya-andisha=раъя-андиша", "raya=раъя", "reket=рэкет",
    "rekognossirovka=рекогносцировка", "relsli=рельсли", "relssiz=рельссиз", "retush=ретушь",
    "revalvatsiya=ревальвация", "revolver=револьвер", "riyel=риель", "rubilnik=рубильник",
    "salto=сальто", "sanga=цанга", "sapfa=цапфа", "sedra=цедра", "sekretar=секретарь",
    "selderey=сельдерей", "sent=цент", "senyor=сеньор", "senyora=сеньора", "seytnot=цейтнот",
    "seziy=цезий", "shagren=шагрень", "shampun=шампунь", "sherst=шерсть", "shifonyer=шифоньер",
    "shmutstitul=шмуцтитул", "shprits=шприц", "shtangensirkul=штангенциркуль", "sian=циан",
    "sianli=цианли", "sinkalamoq=синькаламоқ", "sinkograf=цинкограф", "sinkografiya=цинкография",
    "siren=сирень", "sirkchi=циркчи", "sirkoniy=цирконий", "sirkul=циркуль", "sirroz=цирроз",
    "sistit=цистит", "sitodiagnostika=цитодиагностика", "sitokimyo=цитокимё", "sitoliz=цитолиз",
    "sitologiya=цитология", "skalpel=скальпель", "skeptitsizm=скептицизм", "slanets=сланец",
    "slesar=слесарь", "sobol=соболь", "sokol=цоколь", "ssenarist=сценарист", "statya=статья",
    "stelka=стелька", "stil=стиль", "stoitsizm=стоицизм", "stronsiy=стронций", "sunami=цунами",
    "sutemizuvchilar=сутэмизувчилар", "tablitsa=таблица", "talk=тальк", "telefilm=телефильм",
    "teleinssenirovka=телеинсценировка", "tendensioz=тенденциоз", "tendensiozlik=тенденциозлик",
    "teplitsa=теплица", "terset=терцет", "tersiya=терция", "tigel=тигель", "tonnel=тоннель",
    "transkripsion=транскрипцион", "transkripsiya=транскрипция",
    "transkripsiyalamoq=транскрипцияламоқ", "tyul=тюль", "ultimatum=ультиматум",
    "ultrabinafsha=ультрабинафша", "ultramikroskop=ультрамикроскоп", "ultraqisqa=ультрақисқа",
    "ultratovush=ультратовуш", "util=утиль", "vaksina=вакцина", "valvatsiya=вальвация", "vanil=ваниль",
    "varyete=варьете", "vedomost=ведомость", "verf=верфь", "vernyer=верньер", "vestibyul=вестибюль",
    "videofilm=видеофильм", "viklyuchatel=виключатель", "vintsimon=винтсимон", "vinyetka=виньетка",
    "vklyuchatel=включатель", "vodevil=водевиль", "volost=волость", "voltli=вольтли",
    "xemosorbsiya=хемосорбция", "xolodilnik=холодильник", "xrustal=хрусталь", "yodli=йодли",
    "yotsiramoq=ётсирамоқ", "yuan=юань"
  ];

  var LC_MAP = {
    a: "а", b: "б", d: "д", e: "е", f: "ф", g: "г", h: "ҳ", i: "и",
    j: "ж", k: "к", l: "л", m: "м", n: "н", o: "о", p: "п", q: "қ",
    r: "р", s: "с", t: "т", u: "у", v: "в", x: "х", y: "й", z: "з",
    "ö": "ў", "ğ": "ғ", "ş": "ш", "ç": "ч"
  };
  var CYR_VOW_OUT = "аеёиоуэюяў";  // а е ё и о у э ю я ў
  var UZ_LAT = /^[abdefghijklmnopqrstuvxyz\u00F6\u011F\u015F\u00E7]$/;   // c va w yo'q

  // eski lotin kalit shakli: kichik harf, yangi harflar -> sh/ch/o'/g', apostroflar -> '
  function latKey(w) {
    var out = "";
    for (var i = 0; i < w.length; i++) {
      var c = w.charAt(i).toLowerCase();
      if (c === "ş") out += "sh"; else if (c === "ç") out += "ch";
      else if (c === "ö") out += "o'"; else if (c === "ğ") out += "g'";
      else if (isApos(c)) out += "'"; else out += c;
    }
    return out;
  }

  function parseLatCyr(list) {
    var map = {};
    (list || []).forEach(function (raw) {
      var e = String(raw).trim(), eq = e.indexOf("=");
      if (eq <= 0) return;
      var k = e.slice(0, eq).trim(), v = e.slice(eq + 1).trim().toLowerCase(), star = false;
      if (k.charAt(k.length - 1) === "*") { star = true; k = k.slice(0, -1); }
      k = latKey(k);
      if (k && v) map[k] = { cyr: v, prefix: star || k.length >= 5 };
    });
    return map;
  }

  // Kirill istisnolaridan teskari yozuvlar: faqat oddiy (kichik harfli) lotin shakllari
  // (октябр=oktabr -> oktabr=октябр), brendlar (Microsoft) olinmaydi.
  function reverseCyrExceptions(list) {
    var out = [];
    (list || []).forEach(function (raw) {
      var e = String(raw).trim(), eq = e.indexOf("=");
      if (eq <= 0) return;
      var cyr = e.slice(0, eq).trim(), lat = e.slice(eq + 1).trim();
      if (lat && lat === lat.toLowerCase() && /^[a-zöğşç'ʻʼ-]+$/.test(lat)) out.push(lat + "=" + cyr);
    });
    return out;
  }

  function findLatRule(key, map) {
    for (var L = key.length; L > 0; L--) {
      var v = map[key.slice(0, L)];
      if (v !== undefined && (L === key.length || v.prefix)) return { len: L, cyr: v.cyr };
    }
    return null;
  }

  /** Bitta so'z (harflar va ichki apostroflar) — kirillga. */
  function latWordToCyr(run, oldMode, map) {
    var lower = run.toLowerCase();
    var letters = run.replace(/[^A-Za-zÀ-ɏ]/g, "");
    var allCaps = letters.length > 1 && letters === letters.toUpperCase();
    var capFirst = run.charAt(0) !== lower.charAt(0);

    // so'z ichidagi aralash registr (iPhone, McAfee) — chet so'z
    if (!allCaps && /[a-z][A-Z]/.test(run)) return null;

    var out = "", i = 0, n = lower.length;
    // lug'at
    var key = latKey(run), rule = findLatRule(key, map), skipTo = 0;
    if (rule) {
      // kalit uzunligini asl so'zdagi belgilar soniga moslash
      var consumed = 0, kk = 0;
      while (kk < rule.len && consumed < n) {
        var ch0 = lower.charAt(consumed);
        if (ch0 === "ş" || ch0 === "ç" || ch0 === "ö" || ch0 === "ğ") kk += 2; else kk += 1;
        consumed++;
      }
      out = rule.cyr;
      if (consumed < n && out.charAt(out.length - 1) === "ь") out = out.slice(0, -1);   // мебель + лар -> мебеллар
      skipTo = consumed;
    }
    i = skipTo;
    var prevVowel = function () { return out.length === 0 ? null : CYR_VOW_OUT.indexOf(out.charAt(out.length - 1)) >= 0; };

    while (i < n) {
      var c = lower.charAt(i), nx = lower.charAt(i + 1), nx2 = lower.charAt(i + 2);
      var atStart = out.length === 0;
      if ((c === "o" || c === "g") && isApos(nx)) { out += c === "o" ? "ў" : "ғ"; i += 2; continue; }
      if (oldMode && c === "s" && nx === "h") { out += "ш"; i += 2; continue; }
      if (oldMode && c === "c" && nx === "h") { out += "ч"; i += 2; continue; }
      if (isApos(c)) {
        if (lower.charAt(i - 1) === "s" && nx === "h") { i++; continue; }      // sʼh -> сҳ
        if (i > 0 && nx && UZ_LAT.test(nx)) { out += "ъ"; i++; continue; }  // tutuq -> ъ
        out += run.charAt(i); i++; continue;
      }
      // -tsiya/-tsion -> -ция/-цион, -ksiya/-ksion -> -кция/-кцион (informatsiya, funksiya, operatsion)
      if ((c === "t" || c === "k") && nx === "s" && nx2 === "i" && (lower.charAt(i + 3) === "y" || lower.charAt(i + 3) === "o")) {
        out += c === "t" ? "\u0446" : "\u043A\u0446"; i += 2; continue;
      }
      // ts + e / ts + i(undosh) -> ц (atseton, vitse, defitsit, genotsid); o'zbekcha -sa/-siz/-sin -> тс
      if (c === "t" && nx === "s" && i > 0) {
        var a3 = lower.charAt(i + 3), a4 = lower.charAt(i + 4);
        if (nx2 === "e" || (nx2 === "i" && a3 && a3 !== "z" && (a3 !== "n" || a4 === "t") && UZ_LAT.test(a3) &&
            ("aeiou".indexOf(a3) < 0 || a3 === "a" || a3 === "e"))) {
          out += "\u0446"; i += 2; continue;
        }
      }
      // -ensiya/-ensial -> -енция/-енциал (konferensiya, konfidensial), lekin pensiya; porsiya -> порция
      if (c === "e" && lower.substr(i + 1, 3) === "nsi" && (lower.charAt(i + 4) === "y" || lower.substr(i + 4, 2) === "al") &&
          lower.charAt(i - 1) !== "p" && out.length > 0) {
        out += (prevVowel() ? "\u044D" : "\u0435") + "\u043D\u0446"; i += 3; continue;
      }
      if (lower.substr(i, 5) === "porsi") { out += "\u043F\u043E\u0440\u0446"; i += 4; continue; }
      if (lower.substr(i, 8) === "stansiya") { out += "\u0441\u0442\u0430\u043D\u0446"; i += 5; continue; }  // ...станция
      if (c === "y") {
        var oNext = nx === "o" && !isApos(nx2);
        if (oNext) { out += "ё"; i += 2; continue; }                       // ё
        if (nx === "a") { out += "я"; i += 2; continue; }                   // я
        if (nx === "u") { out += "ю"; i += 2; continue; }                   // ю
        if (nx === "e") {
          out += (atStart || prevVowel()) ? "е" : "ъе";          // е / ъе (obyekt)
          i += 2; continue;
        }
        out += "й"; i++; continue;
      }
      if (c === "e") { out += (atStart || prevVowel()) ? "э" : "е"; i++; continue; }  // э / е
      out += LC_MAP[c] !== undefined ? LC_MAP[c] : run.charAt(i);
      i++;
    }
    if (allCaps) return out.toUpperCase();
    if (capFirst) return out.charAt(0).toUpperCase() + out.slice(1);
    return out;
  }

  /** Lotin matnini (yangi yoki amaldagi alifbo) o'zbek kirill yozuviga o'giradi. */
  function latToCyr(text, opts) {
    opts = opts || {};
    var edits = [];
    if (opts.skipEnglish !== false && isEnglishText(text)) return { edits: [], text: text, skipped: "english" };
    var oldMode = opts.source === "old" ? true : opts.source === "new" ? false : !/[öğşçÖĞŞÇ]/.test(text);
    var exceptions = opts.exceptions || DEFAULT_EXCEPTIONS;
    // kirill istisnolarining teskarisi avval, lug'at keyin (lug'at ustun turadi)
    var map = parseLatCyr(reverseCyrExceptions(opts.cyrExceptions || DEFAULT_CYR_EXCEPTIONS).concat(opts.latCyr || DEFAULT_LAT_CYR));

    var quoteOpen = false, i = 0, n = text.length;
    while (i < n) {
      if (isSpace(text.charAt(i))) { i++; continue; }
      var start = i;
      while (i < n && !isSpace(text.charAt(i))) i++;
      var tok = text.slice(start, i);

      // qo'shtirnoqlar (convertText dagi kabi)
      var lead = 0;
      while (lead < tok.length && !isLetter(tok.charAt(lead)) && !/[0-9]/.test(tok.charAt(lead))) {
        if (isApos(tok.charAt(lead))) quoteOpen = true;
        lead++;
      }
      var limit = tok.length, t = tok.length;
      while (t > lead && !isLetter(tok.charAt(t - 1)) && !isApos(tok.charAt(t - 1))) t--;
      if (t > lead && isApos(tok.charAt(t - 1)) && quoteOpen) { limit = t - 1; quoteOpen = false; }

      var core = coreOf(tok);
      if (!core || looksLikeLinkOrCode(tok)) continue;
      // istisno (SHA-256, Microsoft ...) — lekin lug'atda bor so'z (Ishoq -> Исҳоқ) baribir o'giriladi
      if (findException(core, exceptions) && !findLatRule(latKey(core), map)) continue;
      if (/[Ѐ-ӿ]/.test(core)) continue;                               // allaqachon kirill

      // so'zlarni (harflar + ichki apostrof) ajratib o'giramiz
      var j = lead;
      while (j < limit) {
        if (!isLetter(tok.charAt(j))) { j++; continue; }
        var ws = j;
        while (j < limit && (isLetter(tok.charAt(j)) ||
               (isApos(tok.charAt(j)) && (isLetter(tok.charAt(j + 1)) || /[og]/i.test(tok.charAt(j - 1)))))) j++;
        var run = tok.slice(ws, j);
        var rl = run.toLowerCase(), foreign = false;
        for (var q = 0; q < rl.length; q++) {
          var ch = rl.charAt(q);
          if (isApos(ch)) continue;
          if (ch === "c") { if (rl.charAt(q + 1) === "h" && !/[\u00F6\u011F\u015F\u00E7]/.test(rl)) continue; foreign = true; break; }
          if (!UZ_LAT.test(ch)) { foreign = true; break; }
        }
        if (foreign) continue;
        // so'zda yangi harf bo'lsa — yangi alifbo (sh = s+h), aks holda abzas rejimi
        var wordOld = !/[\u00F6\u011F\u015F\u00E7]/.test(rl);
        var cyr = latWordToCyr(run, wordOld, map);
        if (cyr !== null && cyr !== run) edits.push({ pos: start + ws, len: run.length, from: run, to: cyr });
      }
    }
    var res = text;
    for (var e = edits.length - 1; e >= 0; e--) res = res.slice(0, edits[e].pos) + edits[e].to + res.slice(edits[e].pos + edits[e].len);
    return { edits: edits, text: res, skipped: null, oldMode: oldMode };
  }

  // ======================================================== YANGI LOTIN → ESKI (AMALDAGI) LOTIN
  /** Özbekiston → Oʻzbekiston, ŞAHAR → SHAHAR (ö→oʻ, ğ→gʻ, ş→sh, ç→ch). */
  function newToOld(text) {
    var edits = [], i = 0, n = text.length;
    while (i < n) {
      if (isSpace(text.charAt(i))) { i++; continue; }
      var start = i;
      while (i < n && !isSpace(text.charAt(i))) i++;
      var tok = text.slice(start, i);
      if (!/[öğşçÖĞŞÇ]/.test(tok)) continue;
      var out = revertWord(tok);
      if (out !== tok) edits.push({ pos: start, len: tok.length, from: tok, to: out });
    }
    var res = text;
    for (var e = edits.length - 1; e >= 0; e--) res = res.slice(0, edits[e].pos) + edits[e].to + res.slice(edits[e].pos + edits[e].len);
    return { edits: edits, text: res, skipped: null };
  }

  var api = {
    convertText: convertText,
    mergeExceptions: mergeExceptions,
    mergeCyrExceptions: mergeCyrExceptions,
    DEFAULT_CYR_EXCEPTIONS: DEFAULT_CYR_EXCEPTIONS,
    findException: findException,
    explainWord: explainWord,
    revertWord: revertWord,
    cyrToLatin: cyrToLatin,
    latToCyr: latToCyr,
    newToOld: newToOld,
    DEFAULT_LAT_CYR: DEFAULT_LAT_CYR,
    isRussianText: isRussianText,
    coreOf: coreOf,
    isEnglishText: isEnglishText,
    isMonoFont: isMonoFont,
    DEFAULT_EXCEPTIONS: DEFAULT_EXCEPTIONS,
    APOS: APOS,
    fixHomoglyphs: fixHomoglyphs
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.OzbekAlifbo = api;
})(this);
