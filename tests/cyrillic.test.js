// Ishga tushirish: node tests/cyrillic.test.js
const assert = require("assert");
const { cyrToLatin, convertText } = require("../2-Office-Add-in/converter.js");

const cases = [
  ["Ўзбекистон Республикаси", "Özbekiston Respublikasi"],
  ["Шаҳар, ШАҲАР, Чўл", "Şahar, ŞAHAR, Çöl"],
  ["ер, ерга, поезд, мева, иероглиф", "yer, yerga, poyezd, meva, iyeroglif"],
  ["Ер юзи", "Yer yuzi"],
  ["цирк, лицей, концерт, Цех", "sirk, litsey, konsert, Sex"],
  ["Қуръон, маъно, объект, съезд", "Qurʼon, maʼno, obyekt, syezd"],
  ["компьютер, бульон, апрель", "kompyuter, bulyon, aprel"],
  ["Исҳоқ, Мусҳаф", "Isʼhoq, Musʼhaf"],
  ["Ёшлар, ЁШЛАР, ёз", "Yoşlar, YOŞLAR, yoz"],
  ["Юрист, ЮРИСТ, Ясси", "Yurist, YURIST, Yassi"],
  ["ғоя, тоғ, ҒАЛАБА", "ğoya, toğ, ĞALABA"],
  ["ящик, эълон, Эркин", "yaşik, eʼlon, Erkin"],
  ["Е", "Ye"],
  ["тонг, минг", "tong, ming"],
  ["«Тошкент» — Toshkent, 2026-йил", "«Toşkent» — Toshkent, 2026-yil"],
  ["Бу ахборот хавфсизлиги бўйича дарслик.", "Bu axborot xavfsizligi böyiça darslik."],
  // ruscha abzas — tegilmaydi
  ["Это было в Ташкенте и не только для студентов.", "Это было в Ташкенте и не только для студентов."],
];

let fail = 0;
for (const [inp, exp] of cases) {
  const out = cyrToLatin(inp).text;
  if (out !== exp) { fail++; console.log("XATO:", inp, "->", out, "| kutilgan", exp); }
  // natija lotin konvertori bilan qayta ishlansa ham buzilmasligi kerak
  const again = convertText(out, { skipEnglish: false }).text;
  if (again !== out && !/Toshkent/.test(inp)) { fail++; console.log("LOTIN BUZDI:", out, "->", again); }
}
assert.strictEqual(cyrToLatin("Это было в Ташкенте и не только.", { skipRussian: false }).text.startsWith("Eto"), true);
// ---- kirill istisnolari
const K = require("../2-Office-Add-in/converter.js");
const eq = (a, b) => { if (a !== b) { fail++; console.log("XATO:", a, "| kutilgan", b); } };
eq(K.cyrToLatin("Майкрософт ва Гугл").text, "Microsoft va Google");
eq(K.cyrToLatin("Майкрософтнинг маҳсулотлари").text, "Microsoftning mahsulotlari");
eq(K.cyrToLatin("ГУГЛГА").text, "GOOGLEGA");
eq(K.cyrToLatin("Ютубдан, Эпплга").text, "YouTubedan, Applega");
const ex = K.mergeCyrExceptions(["!Ворд", "Иванов", "Ташкент*", "Ворд=MS Word", "Гугл=GOOGLE"]);
eq(K.cyrToLatin("Иванов ва Ташкентский Ворд Гугл", { cyrExceptions: ex, skipRussian: false }).text,
   "Иванов va Ташкентский MS Word GOOGLE");
eq(K.cyrToLatin("Иванова", { cyrExceptions: ex, skipRussian: false }).text, "Ivanova");   // aniq so'z: faqat "Иванов"
eq(ex.filter(e => e.startsWith("Гугл")).length, 1);   // foydalanuvchi yozuvi standartni almashtiradi
eq(K.DEFAULT_EXCEPTIONS.includes("Schneier"), true);
// qoidadan chetga chiqadigan so'zlar (lug'at)
eq(K.cyrToLatin("21 октябрь, Сентябрда, АБЗАЦ, абзацлар, коэффициенти").text, "21 oktabr, Sentabrda, ABZAS, abzaslar, koeffitsienti");
eq(K.cyrToLatin("Айфон").text, "iPhone");                        // brend yozilishi saqlanadi
eq(K.convertText("Ishoq va Ishoqov mushafni o'qidi").text, "Ishoq va Ishoqov mushafni öqidi");
eq(K.convertText("Schneier va Shannon shifrlari").text, "Schneier va Shannon şifrlari");
eq(K.convertText("Shor tuproq").text, "Şor tuproq");
console.log(fail ? `${fail} ta xato` : `Kirill testlari o'tdi (${cases.length} ta + istisnolar)`);
process.exit(fail ? 1 : 0);
