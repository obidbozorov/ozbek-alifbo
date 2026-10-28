# Özbekça Yozuv

Windows'ning **istalgan dasturida** (Telegram, brauzer, Bloknot, …) odatdagidek yozganingizda harflar
yangi alifboga özi ögiriladi:

| Yozasiz | Çiqadi |
|---|---|
| `o'` | ö |
| `g'` | ğ |
| `sh` | ş |
| `ch` | ç |

Apostrofning barça turlari (`'` `` ` `` `ʻ` `ʼ` `‘` `’`) qabul qilinadi.

## Örnatiş

1. `Özbekça-Yozuv-1.0.exe` ni yuklab oling va istalgan papkaga qöying (masalan, Hujjatlar).
2. Işga tuşiring — soat yonida **Ö** ikonkasi paydo böladi. Dastur Windows bilan birga özi işga tuşadi.

Öçirib taşlaş: ikonka → «Windows bilan birga işga tuşsin» belgisini oling → «Çiqiş» → exe faylni öçiring.

## Boşqaruv

- **Alt+Shift+Z** yoki ikonkani bosish — yoqiş / öçiriş (kulrang ikonka — öçirilgan).
- Ögirilmasin desangiz — harfni yana bir bor bosing: `s h h` → `sh`, `o ' '` → `o'`.
- **Istisno sözlar** (Photoshop, Chevrolet, English, …) söz tugagaç asl holiga qaytadi.
  Öz sözlaringiz: ikonka → «Istisno sözlar…» (Word plagini bilan umumiy `istisnolar.txt`).
- **Istisno dasturlar**: ikonka → «Istisno dasturlar…». Standart holda Word, Excel, PowerPoint
  (ularda Özbekça plaginining öz «Yozişda» rejimi bor), buyruq satri, dasturlaş muhitlari va parol menejerlari.

## Xavfsizlik

- **Parol maydonlarida işlamaydi** — dastur fokusdagi maydon parol ekanini aniqlab, unga tegmaydi.
- Administrator huquqi bilan işlayotgan oynalarga tegmaydi.
- Bosilgan tugmalar **hech qayerga yozilmaydi va yuborilmaydi**: dastur faqat oxirgi sözni xotirada uşlab turadi.
  Internetga ulanmaydi. Manba kodi şu papkada: [`src/`](src/).

> Dastur hozirça raqamli imzosiz. Windows SmartScreen «Noma'lum nashriyotçi» deb ogohlantirsa —
> «Batafsil» → «Baribir işga tuşiriş». Ba'zi antivirus dasturlari klaviatura bilan işlovçi dasturlarga
> şubha bilan qaraşi mumkin.

## Yig'ish

```sh
sh src/build.sh          # Mono mcs
cd test && mcs EngineTest.cs ../src/Engine.cs && mono EngineTest.exe   # testlar
```
