#!/bin/bash
# Usage: ./reset_bot.sh <nomor_hp>
# Contoh: ./reset_bot.sh 6281228790091

PHONE=${1:-}
if [ -z "$PHONE" ]; then
    echo "Usage: $0 <nomor_hp>"
    echo "Contoh: $0 6281228790091"
    exit 1
fi

echo "=== Reset Bot Session: $PHONE ==="

# 1. Reset DB (hapus paket + specs, pertahankan email)
echo ""
echo "[1/2] Reset database..."
RESULT=$(curl -s -X POST https://adkivia.com/api/bot_save_order.php \
  -H "Content-Type: application/json" \
  -H "X-Bot-Key: adkivia-bot-2026" \
  -d "{\"phone\": \"$PHONE\", \"reset_session\": 1}")
echo "DB: $RESULT"

# 2. Cari dan hapus Redis chat history
echo ""
echo "[2/2] Cari Redis chat history..."
DELETED=0
CURSOR=0
PHONE_SHORT="${PHONE:2}"
PHONE_LOCAL="0${PHONE_SHORT}"
LIST_KEYS=()

while true; do
    READ=$(redis-cli SCAN $CURSOR COUNT 100)
    CURSOR=$(echo "$READ" | sed -n '1p')
    KEYS=$(echo "$READ" | sed -n '2,$p')

    while IFS= read -r KEY; do
        [ -z "$KEY" ] && continue
        TYPE=$(redis-cli TYPE "$KEY")
        [ "$TYPE" != "list" ] && continue

        LIST_KEYS+=("$KEY")
        CONTENT=$(redis-cli LRANGE "$KEY" 0 -1)
        if echo "$CONTENT" | grep -qE "$PHONE|$PHONE_SHORT|$PHONE_LOCAL"; then
            redis-cli DEL "$KEY" > /dev/null
            echo "Redis: Hapus key '$KEY'"
            DELETED=$((DELETED + 1))
        fi
    done <<< "$KEYS"

    [ "$CURSOR" = "0" ] && break
done

if [ "$DELETED" -eq 0 ]; then
    echo "Redis: Tidak ditemukan otomatis."
    if [ ${#LIST_KEYS[@]} -gt 0 ]; then
        echo ""
        echo "List key Redis yang ada (cek manual mana milik user ini):"
        for K in "${LIST_KEYS[@]}"; do
            echo ""
            echo "  Key: $K"
            redis-cli LRANGE "$K" 0 1 | sed 's/^/    /'
        done
        echo ""
        echo "Hapus manual: redis-cli DEL \"<key>\""
    else
        echo "Redis: Tidak ada chat history sama sekali."
    fi
else
    echo "Redis: $DELETED key dihapus"
fi

echo ""
echo "=== Selesai ==="
