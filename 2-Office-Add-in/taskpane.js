/* Yangi alifbo — Word Add-in paneli */
/* global Office, Word, OzbekAlifbo */
(function () {
  "use strict";

  var BATCH = 250;               // bir sync'da nechta abzas
  var CTRL = /[\u0000-\u0008\u000B-\u001F]/; // maydon, izoh belgisi va h.k.
  var live = { handler: null, timer: null, ids: {}, busy: false };

  // ---------------------------------------------------------------- sozlamalar
  function store(key, val) {
    try { if (val === undefined) return localStorage.getItem("ozal." + key); localStorage.setItem("ozal." + key, val); }
    catch (e) { return null; }
  }
  // Istisnolar holati: foydalanuvchi so'zlari va o'chirilgan standartlar
  var exc = { user: [], off: [] };
  var cyrExc = { user: [], off: [] };   // kirill istisnolari

  function loadExc() {
    try { exc.user = JSON.parse(store("excUser") || "[]"); } catch (e) { exc.user = []; }
    try { exc.off = JSON.parse(store("excOff") || "[]"); } catch (e) { exc.off = []; }
    var old = store("exc");                      // eski versiyadagi matn maydoni
    if (old && !exc.user.length) { importLines(old.split(/\r?\n/), false); store("exc", ""); }
  }
  function loadCyr() {
    try { cyrExc.user = JSON.parse(store("cyrUser") || "[]"); } catch (e) { cyrExc.user = []; }
    try { cyrExc.off = JSON.parse(store("cyrOff") || "[]"); } catch (e) { cyrExc.off = []; }
  }
  function saveCyr() {
    store("cyrUser", JSON.stringify(cyrExc.user));
    store("cyrOff", JSON.stringify(cyrExc.off));
  }
  function cyrLines() { return cyrExc.off.map(function (d) { return "!" + d; }).concat(cyrExc.user); }
  function cyrKey(e) { return String(e).split("=")[0].trim().toLowerCase(); }
  function cyrOpts() {
    return { skipRussian: document.getElementById("chkRussian").checked,
             cyrExceptions: OzbekAlifbo.mergeCyrExceptions(cyrLines()) };
  }

  function saveExc() {
    store("excUser", JSON.stringify(exc.user));
    store("excOff", JSON.stringify(exc.off));
  }
  function excLines() {
    return exc.off.map(function (d) { return "!" + d; }).concat(exc.user);
  }
  function importLines(lines, replace) {
    if (replace) { exc.user = []; exc.off = []; }
    lines.forEach(function (raw) {
      var ln = String(raw).trim();
      if (!ln || ln.charAt(0) === ";") return;
      if (ln.charAt(0) === "!") {
        var d = ln.slice(1).trim();
        if (d && exc.off.indexOf(d) < 0) exc.off.push(d);
      } else if (exc.user.indexOf(ln) < 0) exc.user.push(ln);
    });
    saveExc();
  }

  function getOpts() {
    return {
      tutuq: document.getElementById("chkTutuq").checked,
      skipEnglish: document.getElementById("chkEnglish").checked,
      exceptions: OzbekAlifbo.mergeExceptions(excLines())
    };
  }
  function setStatus(msg, cls) {
    var el = document.getElementById("status");
    el.textContent = msg; el.className = cls || "";
  }
  function busy(on) {
    document.getElementById("btnDoc").disabled = on;
    document.getElementById("btnSel").disabled = on;
    document.getElementById("btnCyrDoc").disabled = on;
    document.getElementById("btnCyrSel").disabled = on;
    document.getElementById("btnL2CDoc").disabled = on;
    document.getElementById("btnL2CSel").disabled = on;
  }

  // ---------------------------------------------------------------- yadro
  /**
   * Abzaslar to'plamini o'giradi. selRange berilsa — faqat uning ichidagi so'zlar.
   * Har bir so'z (bo'shliq bilan ajratilgan bo'lak) alohida almashtiriladi,
   * shuning uchun abzasning qolgan formatlashi saqlanadi.
   */
  async function convertParagraphItems(ctx, items, selRange, opts, stat) {
    for (var s = 0; s < items.length; s += BATCH) {
      var part = items.slice(s, s + BATCH);
      var jobs = [];
      part.forEach(function (p) {
        if (OzbekAlifbo.isMonoFont(p.font.name)) return;
        var res = opts.convert ? opts.convert(p.text) : OzbekAlifbo.convertText(p.text, opts);
        if (!res.edits.length) return;
        var chunks = p.getTextRanges([" ", "\t"], true);
        chunks.load("items/text,items/font/name");
        jobs.push({ p: p, text: p.text, edits: res.edits, chunks: chunks });
      });
      if (!jobs.length) continue;
      await ctx.sync();

      var plan = [];
      jobs.forEach(function (job) {
        var cursor = 0, ei = 0;
        for (var c = 0; c < job.chunks.items.length; c++) {
          var ch = job.chunks.items[c];
          var t = ch.text;
          if (!t) continue;
          var at = job.text.indexOf(t, cursor);
          if (at < 0) { stat.skipped++; break; }
          var end = at + t.length;
          cursor = end;
          var mine = [];
          while (ei < job.edits.length && job.edits[ei].pos < end) {
            var ed = job.edits[ei++];
            if (ed.pos >= at && ed.pos + ed.len <= end) mine.push(ed);
          }
          if (!mine.length) continue;
          if (CTRL.test(t) || OzbekAlifbo.isMonoFont(ch.font.name)) { stat.skipped += mine.length; continue; }
          var nt = t;
          for (var k = mine.length - 1; k >= 0; k--) {
            var off = mine[k].pos - at;
            nt = nt.slice(0, off) + mine[k].to + nt.slice(off + mine[k].len);
          }
          var item = { range: ch, text: nt, n: mine.length, rel: null };
          if (selRange) item.rel = ch.compareLocationWith(selRange);
          plan.push(item);
        }
      });
      if (selRange && plan.length) await ctx.sync();

      plan.forEach(function (it) {
        if (it.rel && ["Inside", "InsideStart", "InsideEnd", "Equal"].indexOf(it.rel.value) < 0) return;
        it.range.insertText(it.text, "Replace");
        stat.changed += it.n;
      });
      await ctx.sync();
    }
  }

  async function collectBodies(ctx) {
    var bodies = [ctx.document.body];
    var sections = ctx.document.sections;
    sections.load("items");
    await ctx.sync();
    sections.items.forEach(function (sec) {
      ["Primary", "FirstPage", "EvenPages"].forEach(function (type) {
        bodies.push(sec.getHeader(type));
        bodies.push(sec.getFooter(type));
      });
    });
    if (Office.context.requirements.isSetSupported("WordApi", "1.5")) {
      try {
        var fn = ctx.document.body.footnotes, en = ctx.document.body.endnotes;
        fn.load("items/body"); en.load("items/body");
        await ctx.sync();
        fn.items.forEach(function (n) { bodies.push(n.body); });
        en.items.forEach(function (n) { bodies.push(n.body); });
      } catch (e) { /* izohlar qo'llab-quvvatlanmasa — o'tkazamiz */ }
    }
    return bodies;
  }

  async function runConvert(onlySelection, mode) {
    var opts = getOpts();
    var cyr = mode === "cyr";
    if (cyr) {
      var co = cyrOpts();
      opts.convert = function (t) { return OzbekAlifbo.cyrToLatin(t, co); };
    }
    var l2c = mode === "l2c";
    if (l2c) {
      var lo = { skipEnglish: opts.skipEnglish, exceptions: opts.exceptions, cyrExceptions: cyrOpts().cyrExceptions };
      opts.convert = function (t) { return OzbekAlifbo.latToCyr(t, lo); };
    }
    var n2o = mode === "n2o";
    if (n2o) opts.convert = function (t) { return OzbekAlifbo.newToOld(t); };
    var stat = { changed: 0, skipped: 0 };
    busy(true);
    setStatus("Oʻgirilmoqda…");
    try {
      await Word.run(async function (ctx) {
        if (onlySelection) {
          var sel = ctx.document.getSelection();
          sel.load("isEmpty");
          var paras = sel.paragraphs;
          paras.load("items/text,items/font/name");
          await ctx.sync();
          if (sel.isEmpty) { setStatus("Avval matnni belgilang.", "warn"); return; }
          await convertParagraphItems(ctx, paras.items, sel, opts, stat);
        } else {
          var bodies = await collectBodies(ctx);
          var lists = bodies.map(function (b) {
            var ps = b.paragraphs; ps.load("items/text,items/font/name"); return ps;
          });
          await ctx.sync();
          for (var i = 0; i < lists.length; i++) {
            await convertParagraphItems(ctx, lists[i].items, null, opts, stat);
          }
        }
        setStatus(stat.changed + (cyr ? " ta soʻz kirilldan oʻgirildi." : l2c ? " ta soʻz kirillga oʻgirildi." : n2o ? " ta soʻz eski lotinga oʻgirildi." : " ta harf oʻgirildi.") +
          (stat.skipped ? " " + stat.skipped + " ta joy oʻtkazildi (maydon/izoh belgisi)." : "") +
          " Bekor qilish: Ctrl+Z", stat.skipped ? "warn" : "");
      });
    } catch (e) {
      setStatus("Xatolik: " + (e && e.message ? e.message : e), "err");
    } finally {
      busy(false);
    }
  }

  // ---------------------------------------------------------------- yozish paytida
  async function liveOn() {
    if (!Office.context.requirements.isSetSupported("WordApi", "1.6")) {
      setStatus("Bu Word versiyasi yozish paytida oʻgirishni qoʻllamaydi (WordApi 1.6 kerak).", "warn");
      return false;
    }
    await Word.run(async function (ctx) {
      live.handler = ctx.document.onParagraphChanged.add(onParagraphChanged);
      await ctx.sync();
    });
    setStatus("Yozish paytida oʻgirish yoqildi.");
    return true;
  }

  async function liveOff() {
    if (!live.handler) return;
    await Word.run(live.handler.context, async function (ctx) {
      live.handler.remove();
      await ctx.sync();
    });
    live.handler = null;
    setStatus("Yozish paytida oʻgirish oʻchirildi.");
  }

  function onParagraphChanged(ev) {
    (ev.uniqueLocalIds || []).forEach(function (id) { live.ids[id] = true; });
    clearTimeout(live.timer);
    live.timer = setTimeout(flushLive, 700); // yozish to'xtaganda
    return Promise.resolve();
  }

  async function flushLive() {
    if (live.busy) { live.timer = setTimeout(flushLive, 400); return; }
    var ids = Object.keys(live.ids);
    live.ids = {};
    if (!ids.length) return;
    live.busy = true;
    try {
      await Word.run(async function (ctx) {
        var items = ids.map(function (id) {
          var p = ctx.document.getParagraphByUniqueLocalId(id);
          p.load("text,font/name");
          return p;
        });
        await ctx.sync();
        await convertParagraphItems(ctx, items, null, getOpts(), { changed: 0, skipped: 0 });
      });
    } catch (e) { /* abzas o'chirilgan bo'lishi mumkin */ }
    finally { live.busy = false; }
  }

  // ---------------------------------------------------------------- istisnolar paneli
  function $(id) { return document.getElementById(id); }

  function kind(w) {
    var t = [];
    if (/\*$/.test(w)) t.push("boshlanishi");
    else if (/#$/.test(w)) t.push("+ raqam");
    if (w.toLowerCase() !== w) t.push("registr");
    return t.join(", ");
  }

  function li(word, opts) {
    var el = document.createElement("li");
    var w = document.createElement("span"); w.className = "w"; w.textContent = word;
    var tg = document.createElement("span"); tg.className = "tag"; tg.textContent = kind(word);
    if (opts.checkbox) {
      var cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = !opts.off;
      cb.title = "Yoqilgan / oʻchirilgan";
      cb.addEventListener("change", function () { opts.onToggle(cb.checked); });
      el.appendChild(cb);
      if (opts.off) el.className = "off";
    }
    el.appendChild(w); el.appendChild(tg);
    if (opts.onDelete) {
      var del = document.createElement("button"); del.textContent = "✕"; del.title = "Oʻchirish";
      del.setAttribute("aria-label", "Oʻchirish: " + word);
      del.addEventListener("click", opts.onDelete);
      el.appendChild(del);
    }
    return el;
  }

  function renderExc() {
    var q = $("excSearch").value.trim().toLowerCase();
    var hit = function (w) { return !q || w.toLowerCase().indexOf(q) >= 0; };

    var ul = $("excUser"); ul.innerHTML = "";
    var users = exc.user.filter(hit);
    users.forEach(function (w) {
      ul.appendChild(li(w, { onDelete: function () {
        exc.user.splice(exc.user.indexOf(w), 1); saveExc(); excChanged("«" + w + "» oʻchirildi.");
      } }));
    });
    if (!users.length) {
      var e = document.createElement("li"); e.className = "empty";
      e.textContent = exc.user.length ? "Hech narsa topilmadi" : "Hali soʻz qoʻshilmagan";
      ul.appendChild(e);
    }

    var ud = $("excDef"); ud.innerHTML = "";
    OzbekAlifbo.DEFAULT_EXCEPTIONS.filter(hit).forEach(function (d) {
      var isOff = exc.off.some(function (o) { return o.toLowerCase() === d.toLowerCase(); });
      ud.appendChild(li(d, { checkbox: true, off: isOff, onToggle: function (on) {
        exc.off = exc.off.filter(function (o) { return o.toLowerCase() !== d.toLowerCase(); });
        if (!on) exc.off.push(d);
        saveExc(); excChanged();
      } }));
    });

    var active = OzbekAlifbo.DEFAULT_EXCEPTIONS.length - exc.off.length;
    $("defCount").textContent = active + " / " + OzbekAlifbo.DEFAULT_EXCEPTIONS.length + " yoqilgan";
    $("excCount").textContent = (active + exc.user.length) + " ta";
    if (document.activeElement !== $("excBulk")) $("excBulk").value = excLines().join("\n");
  }

  function excMsg(text, cls) { var m = $("excMsg"); m.textContent = text || ""; m.style.color = cls === "err" ? "#c62828" : ""; }

  function excChanged(msg) {
    renderExc(); testWord(); preview();
    if (msg !== undefined) excMsg(msg);
  }

  function addWord(raw) {
    var w = OzbekAlifbo.coreOf(String(raw).trim().split(/\s+/)[0] || "");
    var keepPattern = /[*#]$/.test(String(raw).trim());          // shablon belgisini saqlash
    if (keepPattern) w = String(raw).trim();
    if (!w) { excMsg("Soʻz kiriting.", "err"); return false; }
    if (/\s/.test(w)) { excMsg("Faqat bitta soʻz (boʻshliqsiz).", "err"); return false; }
    if (exc.user.indexOf(w) >= 0) { excMsg("«" + w + "» roʻyxatda bor.", "err"); return false; }
    var isDef = OzbekAlifbo.DEFAULT_EXCEPTIONS.filter(function (d) { return d.toLowerCase() === w.toLowerCase(); })[0];
    if (isDef) {
      exc.off = exc.off.filter(function (o) { return o.toLowerCase() !== isDef.toLowerCase(); });
      saveExc(); excChanged("«" + isDef + "» standart roʻyxatda bor — yoqib qoʻyildi.");
      return true;
    }
    exc.user.push(w); saveExc();
    excChanged("«" + w + "» qoʻshildi.");
    return true;
  }

  async function addFromSelection() {
    try {
      await Word.run(async function (ctx) {
        var sel = ctx.document.getSelection();
        sel.load("text");
        await ctx.sync();
        var t = (sel.text || "").trim();
        if (!t) { excMsg("Hujjatda soʻzni belgilang (ustiga ikki marta bosing).", "err"); return; }
        if (/\s/.test(t)) { excMsg("Bitta soʻzni belgilang.", "err"); return; }
        // Allaqachon o'girilgan bo'lsa (Şamir) — hujjatda asl holiga qaytaramiz
        var orig = OzbekAlifbo.revertWord(t);
        if (orig !== t) { sel.insertText(orig, "Replace"); await ctx.sync(); }
        if (addWord(orig) && orig !== t) excMsg("«" + OzbekAlifbo.coreOf(orig) + "» qoʻshildi va hujjatda asl holiga qaytarildi.");
      });
    } catch (e) { excMsg("Xatolik: " + (e && e.message ? e.message : e), "err"); }
  }

  function testWord() {
    var v = $("excTest").value.trim(), out = $("excTestOut");
    if (!v) { out.textContent = ""; out.className = "out"; return; }
    var r = OzbekAlifbo.explainWord(v, getOpts());
    if (r.reason) { out.textContent = "Oʻgirilmaydi — " + r.reason; out.className = "out skip"; }
    else if (r.changed) { out.textContent = r.input + " → " + r.output; out.className = "out ok"; }
    else { out.textContent = "Oʻzgarmaydi (oʻgiriladigan harf yoʻq)"; out.className = "out"; }
  }

  function preview() {
    var tryIn = $("txtTry");
    if (!tryIn) return;
    var t = OzbekAlifbo.cyrToLatin(tryIn.value, cyrOpts()).text;
    var o = getOpts();
    $("outTry").textContent = OzbekAlifbo.convertText(t, o).text;
    $("outTryCyr").textContent = OzbekAlifbo.latToCyr(tryIn.value,
      { skipEnglish: o.skipEnglish, exceptions: o.exceptions, cyrExceptions: cyrOpts().cyrExceptions }).text;
  }

  // ---------------------------------------------------------------- kirill istisnolari
  function cyrLabel(e) {
    var i = e.indexOf("=");
    return i > 0 ? e.slice(0, i) + " → " + e.slice(i + 1) : e + "  (oʻgirilmaydi)";
  }
  function cyrMsg(t, cls) { var m = $("cyrMsg"); m.textContent = t || ""; m.style.color = cls === "err" ? "#c62828" : ""; }

  function renderCyr() {
    var ul = $("cyrUser"); ul.innerHTML = "";
    cyrExc.user.forEach(function (e) {
      ul.appendChild(li(cyrLabel(e), { onDelete: function () {
        cyrExc.user.splice(cyrExc.user.indexOf(e), 1); saveCyr(); renderCyr(); preview(); cyrMsg("Oʻchirildi.");
      } }));
    });
    if (!cyrExc.user.length) { var x = document.createElement("li"); x.className = "empty"; x.textContent = "Hali yozuv yoʻq"; ul.appendChild(x); }

    var ud = $("cyrDef"); ud.innerHTML = "";
    OzbekAlifbo.DEFAULT_CYR_EXCEPTIONS.forEach(function (d) {
      var k = cyrKey(d);
      var isOff = cyrExc.off.some(function (o) { return cyrKey(o) === k; });
      ud.appendChild(li(cyrLabel(d), { checkbox: true, off: isOff, onToggle: function (on) {
        cyrExc.off = cyrExc.off.filter(function (o) { return cyrKey(o) !== k; });
        if (!on) cyrExc.off.push(d.split("=")[0]);
        saveCyr(); renderCyr(); preview();
      } }));
    });
    var active = OzbekAlifbo.DEFAULT_CYR_EXCEPTIONS.length - cyrExc.off.length;
    $("cyrDefCount").textContent = active + " / " + OzbekAlifbo.DEFAULT_CYR_EXCEPTIONS.length + " yoqilgan";
    $("cyrCount").textContent = (active + cyrExc.user.length) + " ta";
  }

  function addCyr() {
    var k = $("cyrKey").value.trim(), l = $("cyrLat").value.trim();
    if (!k) { cyrMsg("Kirill soʻzni kiriting.", "err"); return; }
    if (/\s/.test(k) || !/[\u0400-\u04FF]/.test(k)) { cyrMsg("Bitta kirill soʻz (boʻshliqsiz) kiriting.", "err"); return; }
    if (/[\u0400-\u04FF]/.test(l)) { cyrMsg("Lotin shaklida kirill harf boʻlmasin.", "err"); return; }
    var entry = l ? k + "=" + l : k;
    var key = cyrKey(entry);
    cyrExc.user = cyrExc.user.filter(function (e) { return cyrKey(e) !== key; });
    if (cyrExc.user.indexOf(entry) < 0) cyrExc.user.push(entry);
    saveCyr(); renderCyr(); preview();
    $("cyrKey").value = ""; $("cyrLat").value = "";
    cyrMsg("Qoʻshildi: " + cyrLabel(entry));
  }

  async function cyrFromSelection() {
    try {
      await Word.run(async function (ctx) {
        var sel = ctx.document.getSelection(); sel.load("text"); await ctx.sync();
        var t = (sel.text || "").trim().replace(/^[^\u0400-\u04FF]+|[^\u0400-\u04FF]+$/g, "");
        if (!t || /\s/.test(t)) { cyrMsg("Hujjatda bitta kirill soʻzni belgilang.", "err"); return; }
        $("cyrKey").value = t; $("cyrLat").focus();
        cyrMsg("Lotin shaklini yozing (yoki boʻsh qoldiring) va «Qoʻshish» ni bosing.");
      });
    } catch (e) { cyrMsg("Xatolik: " + (e && e.message ? e.message : e), "err"); }
  }

  // ---------------------------------------------------------------- UI
  function bindUi() {
    var chkE = $("chkEnglish"), chkT = $("chkTutuq");

    loadExc();
    if (store("en") !== null) chkE.checked = store("en") === "1";
    if (store("tutuq") !== null) chkT.checked = store("tutuq") === "1";

    chkE.addEventListener("change", function () { store("en", chkE.checked ? "1" : "0"); preview(); });
    chkT.addEventListener("change", function () { store("tutuq", chkT.checked ? "1" : "0"); preview(); testWord(); });
    var chkR = $("chkRussian");
    if (store("ru") !== null) chkR.checked = store("ru") === "1";
    chkR.addEventListener("change", function () { store("ru", chkR.checked ? "1" : "0"); preview(); });
    $("txtTry").addEventListener("input", preview);

    // istisnolar
    $("excAdd").addEventListener("click", function () { if (addWord($("excNew").value)) $("excNew").value = ""; });
    $("excNew").addEventListener("keydown", function (e) {
      if (e.key === "Enter") { e.preventDefault(); if (addWord($("excNew").value)) $("excNew").value = ""; }
    });
    $("excFromSel").addEventListener("click", addFromSelection);
    $("excSearch").addEventListener("input", renderExc);
    $("excTest").addEventListener("input", testWord);
    $("excBulkApply").addEventListener("click", function () {
      importLines($("excBulk").value.split(/\r?\n/), true);
      excChanged("Roʻyxat yangilandi: " + exc.user.length + " ta soʻz, " + exc.off.length + " ta standart oʻchirilgan.");
    });
    $("excBulkCopy").addEventListener("click", function () {
      var ta = $("excBulk"); ta.value = excLines().join("\n"); ta.select();
      var ok = false;
      try { ok = document.execCommand("copy"); } catch (e) { ok = false; }
      excMsg(ok ? "Nusxa olindi — istisnolar.txt ga joylashingiz mumkin." : "Matn belgilandi — Ctrl+C bosing.");
    });
    renderExc();

    loadCyr();
    $("cyrAdd").addEventListener("click", addCyr);
    $("cyrLat").addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); addCyr(); } });
    $("cyrKey").addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); addCyr(); } });
    $("cyrFromSel").addEventListener("click", cyrFromSelection);
    renderCyr();

    $("btnDoc").addEventListener("click", function () { runConvert(false); });
    $("btnSel").addEventListener("click", function () { runConvert(true); });
    $("btnCyrDoc").addEventListener("click", function () { runConvert(false, "cyr"); });
    $("btnCyrSel").addEventListener("click", function () { runConvert(true, "cyr"); });
    $("btnL2CDoc").addEventListener("click", function () { runConvert(false, "l2c"); });
    $("btnL2CSel").addEventListener("click", function () { runConvert(true, "l2c"); });
    $("btnN2ODoc").addEventListener("click", function () { runConvert(false, "n2o"); });
    $("btnN2OSel").addEventListener("click", function () { runConvert(true, "n2o"); });
    $("chkLive").addEventListener("change", async function (e) {
      var box = e.target;
      try {
        if (box.checked) box.checked = await liveOn();
        else await liveOff();
      } catch (err) {
        box.checked = false;
        setStatus("Xatolik: " + (err && err.message ? err.message : err), "err");
      }
    });
  }

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Word) {
      setStatus("Bu qoʻshimcha faqat Word uchun.", "warn");
      return;
    }
    bindUi();
  });
})();
