// Ishga tushirish: node tests/converter.test.js
const assert = require("assert");
const { convertText } = require("../2-Office-Add-in/converter.js");

const cases = [
  // [kirish, kutilgan natija, opts]
  ["O'zbekiston", "Özbekiston"],
  ["Oʻzbekiston", "Özbekiston"],
  ["O‘ZBEKISTON RESPUBLIKASI", "ÖZBEKISTON RESPUBLIKASI"],
  ["o’qituvchi g‘oya", "öqituvçi ğoya"],
  ["tog' va bog'", "toğ va boğ"],
  ["Shahar, SHAHAR, shaxs", "Şahar, ŞAHAR, şaxs"],
  ["Choy ichish kerak.", "Çoy içiş kerak."],
  ["KIRISH", "KIRIŞ"],
  ["Toshkent shahri", "Toşkent şahri"],
  ["ishchi", "işçi"],
  ["tong, ming, dengiz", "tong, ming, dengiz"],            // ng o'zgarmaydi
  ["Is'hoq, mus'haf", "Is'hoq, mus'haf"],                    // s'h — alohida tovushlar
  ["ma'no, Qur'on", "ma'no, Qur'on"],                        // tutuq (sukut bo'yicha tegmaymiz)
  ["ma'no, Qur'on, Is'hoq", "maʼno, Qurʼon, Isʼhoq", { tutuq: true }],
  ["«O‘zbek tili»", "«Özbek tili»"],
  ["(shartnoma)", "(şartnoma)"],
  ["SHA-256 va SHA3 xesh, SSH protokoli", "SHA-256 va SHA3 xeş, SSH protokoli"],
  ["hash funksiyasi va hasharotlar", "hash funksiyasi va haşarotlar"],
  ["ChaCha20-Poly1305 shifri", "ChaCha20-Poly1305 şifri"],
  ["Microsoft Windows check", "Microsoft Windows check"],     // chet so'zlar (c, w)
  ["https://shop.uz/chat sahifasi", "https://shop.uz/chat sahifasi"],
  ["info@shahar.uz ga yozing", "info@shahar.uz ga yozing"],
  ["‘kino’ so‘zi", "‘kino’ sözi"],
  ["‘Bu o‘qish’ dedi", "‘Bu öqiş’ dedi"],
  ["boo'' deb yozilgan", "boo'' deb yozilgan"],             // ikki apostrof -> tegmaymiz
  ["This is the research of hash chains and shared keys", "This is the research of hash chains and shared keys"],
  ["Ö, Ğ, Ş, Ç", "Ö, Ğ, Ş, Ç"],                                // allaqachon o'girilgan
  ["‘Bu tog’’ dedi", "‘Bu toğ’ dedi"],                      // so'z oxirida gʻ + qo'shtirnoq
  // lotin so'z ichidagi kirill "o'xshash" harflari (TextLib qoidasi)
  ["Tоshkent shаhri", "Toşkent şahri"],
  ["Mаshinа chоp etildi", "Maşina çop etildi"],
  ["SHА-256 va Сох", "SHA-256 va Сох"],
  ["gо‘zаl", "gözal"],
  // kengaytirilgan apostroflar
  ["oʽzbek oˈqituvchi gꞌoz o‛g‛il", "özbek öqituvçi ğoz öğil"],
];

let fail = 0;
for (const [inp, exp, opts] of cases) {
  const out = convertText(inp, opts).text;
  if (out !== exp) { fail++; console.log("XATO:", JSON.stringify(inp), "->", JSON.stringify(out), "kutilgan", JSON.stringify(exp)); }
}
// idempotentlik
for (const [inp, , opts] of cases) {
  const once = convertText(inp, opts).text;
  assert.strictEqual(convertText(once, opts).text, once, "idempotent emas: " + inp);
}
console.log(fail ? `${fail} ta test xato` : `Hammasi o'tdi (${cases.length} ta)`);

// ---- istisnolar boshqaruvi
const A = require("../2-Office-Add-in/converter.js");
const ex1 = A.mergeExceptions(["!hash", "Schnorr", "; izoh", "phish*", "Schnorr"]);
assert.ok(!ex1.includes("hash") && ex1.includes("Schnorr") && ex1.includes("phish*"));
assert.strictEqual(ex1.filter(e => e === "Schnorr").length, 1);
assert.strictEqual(A.convertText("hash Schnorr shifr", { exceptions: ex1 }).text, "haş Schnorr şifr");
assert.strictEqual(A.explainWord("SHA-256").reason, "istisno: SHA#");
assert.strictEqual(A.explainWord("Windows").reason.startsWith("chet"), true);
assert.strictEqual(A.explainWord("shahar").output, "şahar");
assert.strictEqual(A.revertWord("Şamir"), "Shamir");
assert.strictEqual(A.revertWord("ŞAHAR"), "SHAHAR");
assert.strictEqual(A.revertWord("Özbek"), "O\u02BBzbek");
assert.strictEqual(A.revertWord("Ş"), "Sh");
assert.strictEqual(A.revertWord("çip"), "chip");
console.log("Istisnolar boshqaruvi testlari o'tdi");
process.exit(fail ? 1 : 0);
