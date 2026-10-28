#!/bin/sh
# «Özbekça (yangi alifbo)» klaviaturasini o'rnatish (Linux).
#   sh ornatish.sh          — joriy foydalanuvchi uchun (GNOME/KDE Wayland; administrator kerak emas)
#   sudo sh ornatish.sh -s  — butun tizim uchun (X11 va barcha foydalanuvchilar)
set -e
DIR=$(cd "$(dirname "$0")" && pwd)

if [ "$1" = "-s" ]; then
    XKB=${XKB_ROOT:-/usr/share/X11/xkb}
    install -m 644 "$DIR/uz-yangi" "$XKB/symbols/uz-yangi"
    for f in evdev base; do
        x="$XKB/rules/$f.xml"
        if [ -f "$x" ] && ! grep -q "<name>uz-yangi</name>" "$x"; then
            cp "$x" "$x.ozbekca.bak"
            awk -v add="$(sed -n '/<layout>/,/<\/layout>/p' "$DIR/evdev.xml")" \
                '/<layoutList>/ && !done { print; print add; done=1; next } { print }' "$x.ozbekca.bak" > "$x"
        fi
        l="$XKB/rules/$f.lst"
        if [ -f "$l" ] && ! grep -q "^  uz-yangi " "$l"; then
            sed -i "/^! layout/a\\  uz-yangi        Özbekça (yangi alifbo)" "$l"
        fi
    done
    echo "Tizimga o'rnatildi."
else
    CFG="${XDG_CONFIG_HOME:-$HOME/.config}/xkb"
    mkdir -p "$CFG/symbols" "$CFG/rules"
    install -m 644 "$DIR/uz-yangi" "$CFG/symbols/uz-yangi"
    install -m 644 "$DIR/evdev.xml" "$CFG/rules/evdev.xml"
    echo "Foydalanuvchi uchun o'rnatildi: $CFG"
fi
echo "Tizimdan chiqib qayta kiring, keyin Sozlamalar -> Klaviatura -> Kiritish manbalari -> «Özbekça (yangi alifbo)» ni qo'shing."
echo "X11 da darhol sinash:  setxkbmap uz-yangi   (qaytarish: setxkbmap us)"
