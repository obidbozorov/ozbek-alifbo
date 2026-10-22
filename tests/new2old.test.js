// Ishga tushirish: node tests/new2old.test.js
const { newToOld, convertText } = require("../2-Office-Add-in/converter.js");
const cases = [
  ["Özbekiston Respublikasi şaharlari, çöl va toğlar", "Oʻzbekiston Respublikasi shaharlari, choʻl va togʻlar"],
  ["ŞAHAR, ÖZBEKISTON, Ş, çip", "SHAHAR, OʻZBEKISTON, Sh, chip"],
  ["«Özbek tili» — 2026-yil", "«Oʻzbek tili» — 2026-yil"],
  ["SHA-256, Microsoft, Ishoq, https://nuu.uz", "SHA-256, Microsoft, Ishoq, https://nuu.uz"],
];
let fail = 0;
for (const [inp, exp] of cases) {
  const out = newToOld(inp).text;
  if (out !== exp) { fail++; console.log("XATO:", inp, "->", out, "| kutilgan", exp); }
}
// aylanma: eski -> yangi -> eski
const old = "Oʻzbekiston Respublikasi shaharlari, choʻl va togʻlar, KIRISH, ishchi";
const back = newToOld(convertText(old).text).text;
if (back !== old) { fail++; console.log("AYLANMA XATO:", back); }
console.log(fail ? `${fail} ta xato` : `Yangi → eski lotin testlari o'tdi (${cases.length + 1} ta)`);
process.exit(fail ? 1 : 0);
