# Özbekça (yangi alifbo) — Linux klaviaturasi (XKB)

Oddiy AQSh (QWERTY) klaviaturasi + yangi harflar **AltGr** (o'ng Alt) bilan:

| Tugmalar | Natija |
|---|---|
| AltGr + O / G / S / C | ö ğ ş ç |
| Shift + AltGr + O / G / S / C | Ö Ğ Ş Ç |
| AltGr + ' | ʼ (tutuq: maʼno) |
| Shift + AltGr + ' | ʻ |

## O'rnatish

`Ozbekca-Linux.zip` ni oching va papka ichida terminalda:

```sh
sh ornatish.sh              # joriy foydalanuvchi uchun (GNOME / KDE, Wayland)
sudo sh ornatish.sh -s      # butun tizim uchun (X11, barcha foydalanuvchilar)
```

Keyin tizimdan chiqib qayta kiring va **Sozlamalar → Klaviatura → Kiritish manbalari → +** orqali
**«Özbekça (yangi alifbo)»** ni qo'shing (ro'yxatda "Uzbek" ichida yoki qidiruvda).

X11 da o'rnatmasdan sinab ko'rish: `setxkbmap -I ~/.config/xkb uz-yangi` (qaytarish: `setxkbmap us`).

## O'chirish

```sh
rm ~/.config/xkb/symbols/uz-yangi ~/.config/xkb/rules/evdev.xml      # foydalanuvchi o'rnatmasi
```
Tizim o'rnatmasida: `/usr/share/X11/xkb/symbols/uz-yangi` ni o'chiring va `rules/*.ozbekca.bak` nusxalarini qaytaring.

---
Muallif: Bozorov Obidjon Norqobilovich, O'zbekiston Milliy universiteti, Axborot xavfsizligi kafedrasi · MIT
