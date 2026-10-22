// Ishga tushirish: node tests/latin2cyr.test.js
const { latToCyr, cyrToLatin } = require("../2-Office-Add-in/converter.js");

const cases = [
  // yangi alifbo
  ["Özbekiston Respublikasi şaharlari, çöl va toğlar", "Ўзбекистон Республикаси шаҳарлари, чўл ва тоғлар"],
  ["YOŞLAR, ĞALABA, Şahar", "ЁШЛАР, ҒАЛАБА, Шаҳар"],
  // amaldagi lotin (barcha apostrof turlari)
  ["Oʻzbekiston, O'ZBEKISTON, o‘qituvchi, g’oya, tog'", "Ўзбекистон, ЎЗБЕКИСТОН, ўқитувчи, ғоя, тоғ"],
  ["shahar, choʻl, ishchi", "шаҳар, чўл, ишчи"],
  // е / э
  ["Yer yuzi, ekran, Erkin, mevalar, aeroport, poeziya, bugun-erta", "Ер юзи, экран, Эркин, мевалар, аэропорт, поэзия, бугун-эрта"],
  // ye / yo / yu / ya
  ["poyezd, iyeroglif, obyekt, syezd, yoz, yoʻl, tayyor, aktyor, yulduz, Yangi, kelyapti", "поезд, иероглиф, объект, съезд, ёз, йўл, тайёр, актёр, юлдуз, Янги, келяпти"],
  // йо / йа / йе (lug'at)
  ["mayor, general-mayor, rayon, yogurt, poyabzal, konveyer, foye", "майор, генерал-майор, район, йогурт, пойабзал, конвейер, фойе"],
  // tutuq
  ["maʼno, Qurʼon, eʼlon, sanʼat", "маъно, Қуръон, эълон, санъат"],
  ["Isʼhoq, mus'haf, Ishoq, Ishoqov", "Исҳоқ, мусҳаф, Исҳоқ, Исҳоқов"],
  // ц
  ["informatsiya, funksiya, aksiya, operatsion, konstitutsiya, litsey, dotsent, sirk, sement, prinsip", "информация, функция, акция, операцион, конституция, лицей, доцент, цирк, цемент, принцип"],
  ["otsiz, pensiya, versiya, taksi", "отсиз, пенсия, версия, такси"],
  // ь (lug'at, qo'shimchalar bilan)
  ["oktabr, 21-oktabrda, sentabr, mebel, mebellar, kompyuter, avtomobillar, film, filmlar", "октябрь, 21-октябрда, сентябрь, мебель, мебеллар, компьютер, автомобиллар, фильм, фильмлар"],
  // tegilmaydi
  ["SHA-256, Microsoft, Windows, https://nuu.uz, info@nuu.uz, iPhone", "SHA-256, Microsoft, Windows, https://nuu.uz, info@nuu.uz, iPhone"],
  ["This is the research of hash functions and keys.", "This is the research of hash functions and keys."],
  ["‘kino’ va «Toshkent»", "‘кино’ ва «Тошкент»"],
  ["2026-yil, 3-band", "2026-йил, 3-банд"],
];

let fail = 0;
for (const [inp, exp] of cases) {
  const out = latToCyr(inp).text;
  if (out !== exp) { fail++; console.log("XATO:\n  ", inp, "\n  ->", out, "\n  kutilgan", exp); }
}
console.log(fail ? `${fail} ta xato` : `Lotin → kirill testlari o'tdi (${cases.length} ta)`);
process.exit(fail ? 1 : 0);
