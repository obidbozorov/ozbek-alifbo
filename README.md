# Özbekça 1.0 — Word, Excel va PowerPoint uchun plagin (yangi oʻzbek alifbosi)

Lentada uchta asosiy tugma: **Yangi alifboga** (kirill va eski lotindan; har abzas oʻz yozuvi boʻyicha),
**Eskichaga**, **Kirillga**. Matn belgilangan boʻlsa — faqat u, aks holda butun hujjat oʻgiriladi.

**Muallif:** Bozorov Obidjon Norqobilovich — Mirzo Ulugʻbek nomidagi Oʻzbekiston Milliy universiteti,
Axborot xavfsizligi kafedrasi · o.bozorov@nuu.uz · [github.com/obidbozorov](https://github.com/obidbozorov)

Matnni yangi oʻzbek lotin alifbosiga oʻgiradi:

| Amaldagi | Yangi |
|---|---|
| Oʻ oʻ (o', o‘, o’, oʻ, o`) | **Ö ö** |
| Gʻ gʻ | **Ğ ğ** |
| Sh sh | **Ş ş** |
| Ch ch | **Ç ç** |
| Ng ng | oʻzgarmaydi (alifbodan harf sifatida chiqadi, n+g deb yoziladi) |

Kirill yozuvidagi matnni ham toʻgʻridan-toʻgʻri yangi lotinga oʻgiradi: `Ўзбекистон шаҳарлари → Özbekiston şaharlari`,
va teskarisi — lotindan kirillga: `Özbekiston / Oʻzbekiston → Ўзбекистон`.

## Eng oson yoʻl: `Özbekça-1.0.exe`

Oʻrnatuvchini **[Releases](https://github.com/obidbozorov/ozbek-alifbo/releases)** boʻlimidan yuklab oling.

1. Wordni yoping.
2. `4-Ornatuvchi-EXE/Özbekça.exe` ni ikki marta bosing → **Oʻrnatish**.
3. Wordni oching — yuqorida **«Özbekça»** lentasi paydo boʻladi.

Administrator huquqi kerak emas. Windows «Kompyuteringiz himoyalandi» (SmartScreen) desa —
**«Batafsil» → «Baribir ishga tushirish»** (dastur raqamli imzolanmagan, shuning uchun shunday deydi).
Oʻchirish: shu `.exe` ni qayta ochib → **Oʻchirish**.

<details><summary>Oʻrnatuvchi nima qiladi?</summary>

- Wordni yashirin ishga tushirib, `OzbekAlifbo.bas` makros modulidan `Ozbekca.dotm` shablonini yasaydi.
  Buning uchun bir necha soniyaga «VBA loyihasiga kirish» ruxsatini yoqadi va darhol avvalgi holiga qaytaradi.
- Shablonga «Özbekça» lentasi va oʻng tugma menyusini qoʻshadi.
- Shablonni Wordning `STARTUP` papkasiga (`%APPDATA%\Microsoft\Word\STARTUP`) qoʻyadi — Word uni har safar oʻzi yuklaydi.
  Bu papka Word uchun ishonchli joy, shuning uchun makros ogohlantirishi chiqmaydi.

Universitet boʻyicha tarqatish uchun buyruq qatori: `Özbekça.exe /S` — oynasiz oʻrnatish, `/U` — oynasiz oʻchirish.

Agar tashkilot siyosati VBA loyihasiga kirishni taqiqlagan boʻlsa, oʻrnatuvchi buni aytadi — unda VBA ni qoʻlda oʻrnating (pastda).
Manba kodi: `4-Ornatuvchi-EXE/src` (C#, .NET Framework 4 — Windows 10/11 da tayyor turadi).
</details>


### Excel va PowerPoint

`Özbekça.exe` kompyuterda qaysi biri bor bo'lsa — **Word, Excel, PowerPoint** — hammasiga bitta bosishda o'rnatadi
(`%APPDATA%\Microsoft\AddIns\Ozbekca.xlam` va `Ozbekca.ppam`). Lenta bir xil:
to'liq / belgilangan qismni o'girish, eskichaga, kirill ↔ lotin, istisno so'zlar (ro'yxat Word bilan umumiy).

| | Excel | PowerPoint |
|---|---|---|
| Nima o'giriladi | matnli katakchalar (formulalarga tegilmaydi), shakllar | sarlavha, matn bloklari, shakllar, guruhlar, jadvallar, ma'ruzachi izohlari |
| Formatlash | katakcha matni butunligicha almashadi; shakllarda saqlanadi | saqlanadi |
| Bekor qilish | Ctrl+Z (katakchalar) | Ctrl+Z |
| Tezkor tugmalar | Alt+A, E, K, C, I, O (bir nechta katakcha belgilangan bo'lsa — faqat ular, aks holda butun kitob) | yo'q — PowerPoint makrosga tugma bog'lashga ruxsat bermaydi |

Modullar avtomatik yig'iladi: `python3 tools/build_apps.py` → `1-VBA/OzbekAlifbo_Excel.bas`, `OzbekAlifbo_PowerPoint.bas`
va `customUI14_excel.xml` / `customUI14_powerpoint.xml` (konvertor yadrosi Word moduli bilan aynan bir xil).

## Barcha versiyalar

Hammasi bir xil qoidalar bilan ishlaydi:

| Papka | Nima | Qayerda ishlaydi | Oʻrnatish |
|---|---|---|---|
| `4-Ornatuvchi-EXE` | `.exe` oʻrnatuvchi (VBA versiyasini lenta bilan oʻrnatadi) | Windows Word 2010+ | ikki marta bosish |
| `1-VBA` | Makros (`.bas`) | Windows/Mac Word | 2 daqiqa, internetsiz |
| `2-Office-Add-in` | Veb qoʻshimcha | Word 365 (Windows, Mac, brauzer) | HTTPS hosting kerak (GitHub Pages) |
| `3-VSTO` | C# qoʻshimcha | Faqat Windows Word | Visual Studio bilan yigʻiladi, `setup.exe` |

---

## Qoidalar (barcha versiyalarda bir xil)

**Oʻgiriladi:** `Oʻzbekiston → Özbekiston`, `SHAHAR → ŞAHAR`, `ishchi → işçi`, `togʻ → toğ`, `KIRISH → KIRIŞ`.
Apostrofning har qanday koʻrinishi tushuniladi: `' ‘ ’ ʻ ʼ ` ´`.

**Oʻgirilmaydi:**
- `Isʼhoq`, `musʼhaf` — s va h orasida tutuq bor (alohida tovushlar).
- Inglizcha abzaslar (annotatsiya, adabiyotlar roʻyxati) — avtomatik aniqlanadi, sozlamada oʻchirish mumkin.
- Chet soʻzlar: tarkibida `w` yoki yolgʻiz `c` boʻlgan soʻzlar (`Windows`, `Microsoft`, `check`).
- Havolalar, e-pochta, fayl nomlari: `https://…`, `info@…`, `fayl.docx`.
- Monospace shriftdagi matn (Consolas, Courier New…) — dastur kodi.
- Istisno soʻzlar: `SHA-256`, `SSH`, `hash`, `phishing`, `search`, `chain`, `ChaCha20` va boshqalar.
- `boo''` — ikki apostrof ataylab yozilgan deb hisoblanadi.
- `‘kino’` — bitta qoʻshtirnoq ichidagi soʻz oxiridagi apostrof yopuvchi qoʻshtirnoq deb tushuniladi.

---

## Kirill → yangi lotin

| Kirill | Lotin | Qoida |
|---|---|---|
| ш ч ў ғ | ş ç ö ğ | — |
| қ ҳ х ж | q h x j | — |
| е | **ye** / e | soʻz boshida, unlidan va ъ/ь dan keyin `ye` (`ер → yer`, `поезд → poyezd`), qolganida `e` |
| ц | **s** / ts | soʻz boshida va undoshdan keyin `s` (`цирк → sirk`, `концерт → konsert`), unlidan keyin `ts` (`лицей → litsey`) |
| ё ю я | yo yu ya | — |
| ъ | ʼ | `маъно → maʼno`; е/ё/ю/я oldida tushib qoladi (`объект → obyekt`) |
| ь | — | tushib qoladi (`компьютер → kompyuter`); о oldida `y` (`бульон → bulyon`) |
| сҳ | sʼh | `Исҳоқ → Isʼhoq` (aks holda `ş` bilan adashadi) |
| нг | ng | oʻzgarmaydi |

Katta harflar saqlanadi: `ЁШЛАР → YOŞLAR`, `Ёшлар → Yoşlar`.

**Ruscha abzaslar** (iqtiboslar, adabiyotlar roʻyxati) avtomatik aniqlanadi va oʻgirilmaydi: abzasda ў, қ, ғ, ҳ harflari boʻlmasa va
«в, на, не, что, для…» kabi ruscha soʻzlar koʻp boʻlsa. Sozlamada oʻchirish mumkin.

Aralash hujjat (bir qismi kirillcha, bir qismi eski lotincha) boʻlsa — avval **kirilldan**, keyin **lotindan** oʻgiring.

| | VBA | Add-in | VSTO |
|---|---|---|---|
| Butun hujjat | `YangiAlifbo_Kirill` | «Hujjat: kirill → lotin» | «Kirilldan oʻgirish → Butun hujjat» |
| Belgilangan matn | `YangiAlifbo_KirillTanlangan` | «Belgilangan» | «Kirilldan oʻgirish → Belgilangan matn» |
| Ruscha abzasni oʻtkazish | `RUSCHA_ABZASNI_OTKAZISH` | Sozlamalar | Sozlamalar menyusi |

## Lotin → kirill

Yangi alifbodagi (`Özbekiston şaharlari`) ham, amaldagi lotindagi (`Oʻzbekiston shaharlari`) matn ham kirillga oʻgiriladi.
Har bir soʻzda yangi harf (ö ğ ş ç) boʻlsa — u yangi alifbo deb oʻqiladi (`sh` = s+h), aks holda amaldagi (`sh` = ш).

| Lotin | Kirill | Qoida |
|---|---|---|
| e | **э** / е | soʻz boshida va unlidan keyin `э` (`ekran → экран`, `aeroport → аэропорт`), qolganida `е` |
| ye, yo, yu, ya | е, ё, ю, я | `yer → ер`, `tayyor → тайёр`, `aktyor → актёр`; undoshdan keyin `ye → ъе` (`obyekt → объект`) |
| yoʻl | йўл | `y` + `oʻ` alohida |
| ʼ (tutuq) | ъ | `maʼno → маъно`; `sʼh → сҳ` (`Isʼhoq → Исҳоқ`) |
| -tsiya, -ksiya, -ensiya | -ция, -кция, -енция | `informatsiya → информация`, `funksiya → функция`, `konferensiya → конференция`, `konfidensial → конфиденциал` |
| ts + e / ts + i… | ц | `atseton → ацетон`, `vitse-prezident → вице-президент`, `defitsit → дефицит`; lekin `ketsa, otsiz, ketsin → кетса, отсиз, кетсин` |

**Lugʻat** (`istisnolar/standart.json` → `lotin_kirill`, 254 soʻz) — qoida bilan topib boʻlmaydigan soʻzlar: ц (`dotsent → доцент`,
`sirk → цирк`, `sement → цемент`), ь (`oktabr → октябрь`, `mebel → мебель`, `kompyuter → компьютер`), йо/йа (`mayor → майор`,
`rayon → район`, `poyabzal → пойабзал`). 5+ harfli soʻzlar qoʻshimchalari bilan ham taniladi: `21-oktabrda → 21-октябрда`,
`mebellar → мебеллар`. Kirill istisnolaringizdagi oddiy soʻzlar (`доцент=dotsent`) ham teskari yoʻnalishda ishlaydi.

**Oʻgirilmaydi:** istisno soʻzlar (`SHA-256`, `Microsoft`), chet soʻzlar (`w`, yolgʻiz `c`, `iPhone` kabi aralash registr),
havolalar, inglizcha abzaslar, kod shrifti.

**Aniqlik.** Oddiy oʻzbekcha matnda (gazeta, ilmiy maqola) kirill → lotin → kirill aylanmasi 106/106 soʻzda bir xil chiqdi.
Faqat oʻzlashma soʻzlardagi `ц`/`ь` ni (`vaksina → вакцина`, `albom → альбом`) qoida bilan bilib boʻlmaydi — bunday soʻzlarni
lugʻatga qoʻshing. Lugʻat boyigani sari aniqlik oshadi.

| | VBA / exe | Add-in | VSTO | Tugma |
|---|---|---|---|---|
| Lotin → kirill | `YangiAlifbo_LotinKirill`, lentada «Lotin → kirill» | «Hujjat: lotin → kirill» | «Lotin → kirill» | **Alt+C** |

---

## Yangi lotindan eski (amaldagi) lotinga

`Özbekiston → Oʻzbekiston`, `ŞAHAR → SHAHAR`, `çöl → choʻl`. Faqat yangi harf (Ö Ğ Ş Ç) bor soʻzlar oʻzgaradi,
qolgan matn (SHA-256, URL, inglizcha soʻzlar) tegilmaydi. Eski → yangi → eski aylanmasi matnni aynan qaytaradi.

| | VBA / exe | Add-in | VSTO | Tugma |
|---|---|---|---|---|
| Yangi → eski lotin | `YangiAlifbo_EskiLotin`, lentada «Eski lotinga» | «Hujjat: yangi → eski» | «Eski lotinga» | **Alt+E** |

---

## Tezkor tugmalar

| Tugmalar | Amal |
|---|---|
| **Alt+A** | Yangi alifboga (kirill ham, eski lotin ham) — matn belgilangan boʻlsa faqat uni, aks holda butun hujjatni |
| **Alt+E** | Eskichaga (amaldagi lotinga) — xuddi shunday |
| **Alt+C** | Kirillga (har qanday lotindan) — xuddi shunday |
| **Alt+Z** | Yozish paytida oʻgirishni yoqish / oʻchirish |
| **Alt+I** | Kursor turgan soʻzni istisnoga qoʻshish (kirill soʻz boʻlsa — kirill istisnosiga) |
| **Alt+O** | Istisnolarni boshqarish |

Harflar Wordning oʻz Alt-yorliqlari (H, N, P, S, M, R, W, L, F, Y, Q va h.k.) bilan toʻqnashmaydigan qilib tanlangan.

**Band tugmalar.** Oʻrnatuvchi har bir tugmani avval tekshiradi: uni Wordning oʻzi, `Normal.dotm` yoki boshqa
plagin (MathType, Zotero va h.k.) egallagan boʻlsa, navbatdagi boʻsh tugmani oladi:
`Alt+A → Alt+U → Alt+Shift+A → Alt+Shift+U → Ctrl+Alt+Shift+A` (boshqa amallar uchun ham shunday).
Tanlangan tugma lentadagi tugma nomida va **Yordam** oynasida koʻrinadi; keyinchalik biror tugmani boshqa
dastur egallab olsa, Yordam buni koʻrsatadi — oʻrnatuvchida **Yangilash** ni bosish kifoya.

**Oʻzingiz tanlash.** Lentadagi **Tezkor tugmalar** tugmasi har bir amal uchun joriy tugmani va uning holatini
koʻrsatadi. Raqamni kiriting va yangi tugmani yozing (`Alt+L`, `Alt+Shift+L`, `Ctrl+Alt+L`, `Alt+5`). Tugma band boʻlsa,
plagin kim egallaganini aytib, tasdiq soʻraydi. Tanlov `Normal.dotm` ga yoziladi (plagin shablonidan ustun turadi),
**Yangilash** paytida saqlanib qoladi, **Oʻchirish** paytida tozalanadi. Menyudagi oxirgi band hammasini
oʻrnatuvchi tanlagan holatga qaytaradi.

- `Özbekça.exe` ularni oʻrnatish paytida avtomatik qoʻyadi; lentadagi har bir tugma nomining ostida qavs ichida
  tezkor tugmasi yozilgan, masalan «Hujjatni oʻgirish (Alt+A)».
- VBA qoʻlda oʻrnatilgan boʻlsa — `YangiAlifbo_Tugmalar` makrosini bir marta ishga tushiring.
- VSTO da — **Sozlamalar → Tezkor tugmalar** (sukut boʻyicha yoqilgan).
- Office Add-in (veb) Word hujjatidagi tugmalarni ushlay olmaydi — u yerda paneldagi tugmalardan foydalaning.

## Ikonkalar

Lenta tugmalarining ikonkalari `tools/make_icons.py` bilan chiziladi (32×32 PNG, `4-Ornatuvchi-EXE/src/icons/`).
Oʻrnatuvchi ularni shablonga (`customUI/images`) joylaydi; VSTO da ular `icons/*.png` — *Embedded Resource* sifatida
qoʻshing. EXE ni qayta yigʻish: `4-Ornatuvchi-EXE/src/build.sh` (Windowsda — `csc` bilan xuddi shu parametrlar).

---

## Qoʻshimcha qoidalar

- **Apostrofning oʻxshash belgilari:** `' ‘ ’ ʻ ʼ ` ´ ′` dan tashqari `ʹ ʽ ʾ ʿ ˈ ˊ ՙ ՚ ׳ ´ ῾ ‛ ꞌ` ham — PDF, OCR va
  turli klaviaturalardan kelgan matnda `oˈ`, `gꞌ` ham Ö, Ğ boʻladi.
- **Lotin soʻz ichiga adashib kirgan kirill harflari** (koʻrinishi bir xil: а е о р с х у к А В Е К М Н О Р С Т Х)
  tuzatiladi: kirill «о» bilan yozilgan `Tоshkent` → `Toşkent`, `SHА-256` → `SHA-256`. Faqat lotin harfi bor va boshqa
  kirill harfi yoʻq soʻzlarda; haqiqiy kirill soʻzlarga tegilmaydi. (Ikkala qoida — TextLib loyihasidan.)

## Istisno soʻzlarni boshqarish

Uchala versiyada ham bir xil imkoniyatlar bor:

| Imkoniyat | VBA | Add-in | VSTO |
|---|---|---|---|
| Roʻyxatni koʻrish (mening soʻzlarim + standartlar) | ✓ (yangi hujjatda) | ✓ | ✓ |
| Soʻz qoʻshish / oʻchirish | ✓ | ✓ | ✓ (koʻp tanlab, Del) |
| Qidirish | — | ✓ | ✓ |
| Standart istisnoni oʻchirish / qayta yoqish | ✓ | ✓ (belgicha) | ✓ (belgicha) |
| Hujjatdagi soʻzni bir bosishda qoʻshish | ✓ | ✓ | ✓ |
| Sichqonchaning oʻng tugmasi menyusida | ✓ | — | ✓ |
| Soʻzni tekshirish («nega oʻgirilmadi?») | ✓ | ✓ | ✓ |
| Import / eksport | Bloknot orqali | nusxa/joylash | fayl orqali |

**Hujjatdagi soʻzni qoʻshish** — eng qulay yoʻl: soʻz allaqachon notoʻgʻri oʻgirilgan boʻlsa (`Şamir`),
plagin uni hujjatda asl holiga qaytaradi (`Shamir`) va istisnoga qoʻshadi. Keyingi safar bu soʻzga tegmaydi.

**Soʻzni tekshirish** — soʻz kiritsangiz, natijani va sababini koʻrsatadi:
`shahar → şahar`, `SHA-512 — oʻgirilmaydi, istisno: SHA#`, `Windows — chet soʻz`.

**Fayl formati** — `%APPDATA%\OzbekAlifbo\istisnolar.txt` (VBA va VSTO uchun umumiy; Add-in panelidagi «Import / eksport» ham shu format):
```
Schnorr          ← kichik harfli yozuv registrga sezgir emas, katta harf boʻlsa — sezgir
phish*           ← * = shu bilan boshlanadigan har qanday soʻz
SHA#             ← # = keyin faqat raqam, "-" yoki "/" (SHA-256, SHA3) — lekin SHAHAR emas
!hash            ← standart roʻyxatdagi "hash" ni oʻchirish (endi oʻgiriladi)
; izoh
```
Bir kompyuterdan boshqasiga oʻtkazish uchun shu faylni nusxalash kifoya.

**Standart lotin istisnolari** (90 ta, istalganini oʻchirish mumkin):
oʻzbekcha *s+h* soʻzlar (`Ishoq`, `Ishoqov`, `mushaf`, `ashob` — `ş` ga aylanmaydi),
kriptografiya (`SHA-256`, `SSH`, `ChaCha20`, `hash`, `phishing`, `Schnorr`, `Shamir`, `Shannon`, `Schneier`, `Shellshock`…),
IT (`bash`, `shell`, `push`, `fetch`, `Apache`, `PyTorch`, `Chrome`, `SharePoint`, `Photoshop`…),
brendlar va nomlar (`Chevrolet`, `Chery`, `Bosch`, `Shakespeare`, `Massachusetts`…).

## Kirill istisnolari

Kirilldan oʻgirishda alohida roʻyxat ishlaydi — `%APPDATA%\OzbekAlifbo\kirill-istisnolar.txt`:
```
Майкрософт=Microsoft    ← maxsus shakl; qoʻshimchalari odatdagidek: Майкрософтнинг → Microsoftning
Иванов                  ← oʻgirilmaydi (kirillcha qoladi)
Пушкин*                 ← shu bilan boshlanadigan barcha soʻzlar oʻgirilmaydi
!Ворд                   ← standart yozuvni oʻchirish
```
Katta harfli soʻzda shakl ham katta boʻladi: `ГУГЛГА → GOOGLEGA`.

**Standart yozuvlar** (28 ta): brendlar (`Майкрософт → Microsoft`, `Айфон → iPhone`, `Ворд → Word`…) va
qoidadan chetga chiqadigan soʻzlar: `октябрь → oktabr`, `сентябрь → sentabr`, `абзац → abzas`, `коэффициент → koeffitsient` va boshqalar
(qoʻshimchalari bilan ham: `Сентябрда → Sentabrda`, `абзацлар → abzaslar`).

Boshqarish: VBA — `YangiAlifbo_KirillIstisnolar` (yoki lentadagi «Kirill istisnolari»); Add-in — paneldagi «Kirill istisnolari» boʻlimi;
VSTO — «Kirilldan oʻgirish → Kirill istisnolari…». Hujjatdagi kirill soʻzni belgilab, bir bosishda qoʻshish mumkin.

**Tutuq belgisi** (ixtiyoriy, sukut boʻyicha oʻchiq): `maʼno`, `Qurʼon` dagi har xil apostroflarni bitta `ʼ` (U+02BC) ga keltiradi.

---

## 1. VBA makros (qoʻlda oʻrnatish)

`Özbekça.exe` aynan shu modulni oʻrnatadi. Qoʻlda oʻrnatish kerak boʻlsa (masalan, Mac da):

1. Word → `Alt+F11` → chapda **Normal** ni tanlang → **File → Import File…** → `OzbekAlifbo.bas`.
2. Saqlang (`Ctrl+S`) va VBA oynasini yoping.
3. Tugma qoʻshish: **Fayl → Parametrlar → Tez kirish paneli** → «Makroslar» → quyidagilarni qoʻshing:
   - `YangiAlifbo_Hujjat` — butun hujjat (kolontitul, izohlar, matn maydonlari bilan)
   - `YangiAlifbo_Tanlangan` — faqat belgilangan matn
   - `YangiAlifbo_Kirill` / `YangiAlifbo_KirillTanlangan` — kirill → yangi lotin
   - `YangiAlifbo_LotinKirill` / `YangiAlifbo_LotinKirillTanlangan` — lotin (yangi yoki amaldagi) → kirill
   - `YangiAlifbo_Yozishda` — yozish paytida oʻgirishni yoqish/oʻchirish
   - `YangiAlifbo_Istisnolar` — istisnolarni boshqarish menyusi (koʻrish, qoʻshish, oʻchirish, standartlar, tekshirish, Bloknot)
   - `YangiAlifbo_IstisnoQoshish` — kursor turgan soʻzni istisnoga qoʻshish
4. Oʻng tugma menyusi: `YangiAlifbo_KontekstMenyu` makrosini **bir marta** ishga tushiring — soʻz ustida oʻng tugmani bossangiz
   «Istisnolarga qoʻshish» bandi chiqadi. Yana ishga tushirsangiz, olib tashlanadi.

Butun oʻgirish bitta amal hisoblanadi — **Ctrl+Z** hammasini qaytaradi.

**Yozish paytida:** `h`, `Shift+H`, `'` va `` ` `` tugmalari makrosga bogʻlanadi. `s`+`h` yozsangiz darhol `ş` boʻladi.
Asl holini qoldirish kerak boʻlsa (masalan, inglizcha soʻz) — tugmani yana bosing: `ş`+`h` → `sh`, `ö`+`'` → `o'`. Yoki `Ctrl+Z`.
Klaviatura tili avtomatik aniqlanadi — rus/kirill tartibida ham odatdagidek yozadi.

**Hamkasblarga tarqatish:** yangi hujjat oching, `Alt+F11` da shu hujjatga import qiling, **Word makrosli shablon (.dotm)** sifatida
`%APPDATA%\Microsoft\Word\STARTUP` papkasiga saqlang. Shu papkadagi shablon har safar avtomatik yuklanadi.

Sozlamalar faylning boshida: `TUTUQNI_TOGRILASH`, `INGLIZCHA_ABZASNI_OTKAZISH`, `STANDART_ISTISNOLAR`.

---

## 2. Office Add-in (Word 365, brauzer, Mac)

Add-in — bu veb sahifa, shuning uchun fayllar HTTPS manzilda turishi kerak. Eng oson yoʻli — GitHub Pages:

1. GitHub da `ozbek-alifbo` nomli ochiq repozitoriy yarating.
2. `2-Office-Add-in` papkasidagi fayllarni repozitoriy ildiziga yuklang:
   `taskpane.html`, `taskpane.js`, `converter.js`, `assets/`.
3. **Settings → Pages → Branch: main / root → Save.** Bir necha daqiqadan soʻng sahifa
   `https://obidbozorov.github.io/ozbek-alifbo/taskpane.html` manzilida ochiladi.
   Boshqa manzil ishlatsangiz — `manifest.xml` dagi barcha `obidbozorov.github.io/ozbek-alifbo` larni almashtiring.
4. Wordga qoʻshish: **Qoʻshish (Insert) → Qoʻshimchalar (Add-ins) → Mening qoʻshimchalarim → Qoʻshimchani yuklash (Upload My Add-in)** → `manifest.xml`.
   Bosh sahifada **«Özbekça»** tugmasi paydo boʻladi.

Universitet boʻyicha hammaga oʻrnatish: Microsoft 365 administratori **Admin center → Settings → Integrated apps → Upload custom apps** orqali `manifest.xml` ni joylaydi.

Panelda: butun hujjat / belgilangan matn, sozlamalar, istisnolar boshqaruvi va «Sinab koʻrish» maydoni.
Add-in istisnolarni brauzer xotirasida saqlaydi — boshqa kompyuterga oʻtkazish uchun «Import / eksport» dan foydalaning.
**Yozish paytida oʻgirish** — tajribaviy, Word API 1.6 talab qiladi (yangi Word 365): siz yozishni toʻxtatganingizdan ~0,7 soniya keyin abzas oʻgiriladi.

Cheklov: Add-in soʻzlarni butunligicha almashtiradi — agar bitta soʻz ichida turli formatlash boʻlsa (masalan, faqat birinchi harf qalin), soʻz bitta formatga oʻtadi. VBA va VSTO da bunday cheklov yoʻq.

---

## 3. VSTO (C#) qoʻshimcha

Kerak: Visual Studio 2022 + **«Office/SharePoint development»** komponenti.

1. **New Project → Word VSTO Add-in** (.NET Framework 4.8), nomi: `OzbekAlifboAddIn`.
2. Shablon yaratgan `ThisAddIn.cs` ni shu papkadagi `ThisAddIn.cs` bilan almashtiring.
3. Qolgan fayllarni loyihaga qoʻshing (**Add → Existing Item**):
   `Converter.cs`, `DocumentConverter.cs`, `LiveTyping.cs`, `Hotkeys.cs`, `Settings.cs`, `ExceptionsForm.cs`,
   `CyrExceptionsForm.cs`, `AlifboRibbon.cs`, `AlifboRibbon.xml`.
4. `AlifboRibbon.xml` → Properties → **Build Action: Embedded Resource**.
5. **References → Add Reference → System.Windows.Forms** (odatda allaqachon bor).
6. `F5` — Word ochiladi, **Bosh sahifa** da «Özbekça» guruhi chiqadi: oʻgirish tugmalari, **Istisnolar** oynasi,
   **Soʻzni istisnoga** tugmasi. Oʻng tugma menyusida (imlo xatosi bilan belgilangan soʻzlarda ham) «Istisnolarga qoʻshish» bor.
7. Tarqatish: **Build → Publish** → `setup.exe` hosil boʻladi (ClickOnce). Imzolash sertifikati kerak boʻladi.

Yozish paytida oʻgirish VSTO da Word ning xabar oqimini kuzatadi: haqiqiy kiritilgan belgiga qaraydi, shuning uchun klaviatura tili va Wordning «aqlli qoʻshtirnoq» funksiyasi bilan muammosiz ishlaydi. Sozlamalar `HKCU\Software\OzbekAlifbo` da saqlanadi.

---

## Standart roʻyxatlarni oʻzgartirish

Barcha standart istisnolar bitta faylda: `istisnolar/standart.json`. Uni tahrirlab, soʻng
```
python3 tools/sync_lists.py      # JS, C# va VBA kodiga yozadi
python3 tools/check_vba.py       # VBA modulini Wordsiz tekshiradi
```
`Özbekça.exe` ni qayta yigʻish: `4-Ornatuvchi-EXE/src` (buyruq README oxirida).

**Manbalar.** Kirill istisnolari va qoidalar quyidagi ochiq loyihalar maʼlumotlari bilan solishtirildi:
[UzTransliterator](https://github.com/elmurod1202/UzTransliterator) (U. Salaev, E. Kuriyozov; MIT litsenziyasi) —
672 juftlikdan 630 tasi bizning qoidalar bilan mos keldi, farqlilari lugʻatga kiritildi;
[uzbek-transliterator](https://github.com/kodchi/uzbek-transliterator) (GPL-2) — faqat qoidalarni tekshirish uchun, maʼlumot koʻchirilmadi.

## Manbalar

- Lotin → kirill lugʻatining bir qismi ochiq soʻz roʻyxatlari asosida tekshirib toʻldirilgan:
  [kodchi/uzbek-words](https://github.com/kodchi/uzbek-words) va
  [kmashrab/uzbek-wordlist](https://github.com/kmashrab/uzbek-wordlist) (ikkalasi GPL-2). Roʻyxatlarning oʻzi
  koʻchirilmagan — faqat kirill↔lotin aylanmasida xato chiqqan alohida soʻzlarning yozilishi olingan.
- Imlo qoidalari: [imlo.uz](https://imlo.uz/collections/ozbek-tilining-asosiy-imlo-qoidalari).
- Apostrofning oʻxshash belgilari va lotin soʻz ichidagi kirill harflarini tuzatish qoidasi — muallifning
  TextLib / WordFileToLatin loyihasidan.

## Testlar

```
node tests/converter.test.js
node tests/cyrillic.test.js
node tests/latin2cyr.test.js
python3 tools/check_vba.py        # VBA modulini statik tekshirish
python3 tools/test_vba_lo.py      # VBA kodini LibreOffice Basic da ishga tushirib, JS bilan solishtirish
```
VBA konvertor qismi Wordsiz — LibreOffice Basic da haqiqatan ishga tushirib tekshiriladi (322 holat, JS bilan 100% mos).
28 ta holat: oddiy soʻzlar, katta harflar, `ng`, `sʼh`, tutuq, qoʻshtirnoqlar, istisnolar, havolalar, inglizcha abzas;
qoʻshimcha — istisnolarni birlashtirish (`!soʻz`), soʻzni tekshirish va asl holiga qaytarish.
17 ta kirill holati va lugʻat/istisno testlari: е/ц/ъ/ь qoidalari, katta harflar, `сҳ`, ruscha abzas; natija lotin konvertoridan qayta oʻtkazilganda buzilmasligi ham tekshiriladi.
C# konvertori ham shu holatlar bilan tekshirilgan.

## Eʼtibor bering

- Hujjatning **nusxasida** sinab koʻring yoki **Tuzatishlarni kuzatish** (Track Changes) ni yoqib qoʻying — har bir oʻzgarish koʻrinib turadi.
- Oʻzbek matni ichidagi alohida inglizcha soʻzlarni (`machine learning`, `push`) toʻliq aniqlab boʻlmaydi — ularni istisnolarga qoʻshing.
- Yangi harflar barcha shriftlarda bor (Times New Roman, Arial, Calibri…) — turk tili harflari bilan bir xil.

---

<details><summary>Özbekça.exe ni qayta yigʻish</summary>

Windows da (.NET Framework 4 bilan keladigan `csc.exe`) yoki Linux/Mac da Mono `mcs` bilan, `4-Ornatuvchi-EXE/src` papkasida:
```
mcs -target:winexe -out:../Özbekça.exe -win32icon:icon.ico \
  -resource:../../1-VBA/OzbekAlifbo.bas,OzbekAlifbo.bas -resource:customUI14.xml,customUI14.xml \
  -r:System.Windows.Forms.dll -r:System.Drawing.dll -r:System.IO.Compression.dll \
  -r:System.IO.Compression.FileSystem.dll -r:System.Xml.Linq.dll -r:System.Xml.dll Ornatuvchi.cs
```
</details>

## Litsenziya

[MIT](LICENSE) © 2026 Bozorov Obidjon Norqobilovich — Mirzo Ulugʻbek nomidagi Oʻzbekiston Milliy universiteti,
Axborot xavfsizligi kafedrasi.
