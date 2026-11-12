#!/bin/bash
# Собирает приложение для коллег (Docker внутри, ярлыки снаружи) в dist/<версия>/:
#   «Конвертер Фоксфорда — Mac (M1–M5) <версия>.zip»
#   «Конвертер Фоксфорда — Mac (Intel) <версия>.zip»
#   «Конвертер Фоксфорда — Windows <версия>.zip»
# Запуск из корня проекта:  ./desktop/build.sh   (20–40 минут: версия для Windows и Mac Intel собирается в эмуляции)
# Настройки ИИ и контакт для «Сообщить автору» берутся из .env. Папка dist/ в git не попадает.
set -euo pipefail
cd "$(dirname "$0")/.."

VERSION=$(date +%Y.%m.%d-%H%M)
OUT="dist/$VERSION"
IMAGE="foxford-converter-desktop:$VERSION"
mkdir -p "$OUT/images"

# ---------- настройки приложения (из .env: ИИ и контакт; вход без пароля — LOCAL_MODE)
env_value() { grep -E "^$1=" .env 2>/dev/null | tail -1 | cut -d= -f2- | sed -E 's/^["'\'']//; s/["'\'']$//' || true; }
{
  echo "LOCAL_MODE=1"
  echo "JOB_TTL_HOURS=24"
  echo "MAX_UPLOAD_MB=100"
  echo "SUPPORT_CONTACT=$(env_value SUPPORT_CONTACT)"
  echo "ANTHROPIC_API_KEY=$(env_value ANTHROPIC_API_KEY)"
  echo "ANTHROPIC_BASE_URL=$(env_value ANTHROPIC_BASE_URL)"
  echo "ANTHROPIC_AUTH_TOKEN=$(env_value ANTHROPIC_AUTH_TOKEN)"
  echo "AI_DAILY_LIMIT_RUB=$(env_value AI_DAILY_LIMIT_RUB)"
} > "$OUT/images/settings.env"
echo "$VERSION" > "$OUT/images/version.txt"

# ---------- образы: arm64 — Mac на чипе Apple, amd64 — Mac на Intel и Windows
for arch in arm64 amd64; do
  echo "== Собираю образ $arch"
  docker buildx build --platform "linux/$arch" -t "$IMAGE" --load .
  docker save -o "$OUT/images/image-$arch.tar" "$IMAGE"
done
docker rmi "$IMAGE" >/dev/null  # у себя ставим так же, как коллеги, — через «Запустить»

# ---------- папки для коллег
pack() {  # pack <название> <arch> <папка ярлыков>
  local dir="$OUT/Конвертер Фоксфорда — $1 $VERSION"
  mkdir -p "$dir/files"
  cp desktop/Инструкция.html "$dir/"
  cp -R "desktop/$3/." "$dir/"
  cp "$OUT/images/settings.env" "$OUT/images/version.txt" "$dir/files/"
  ln "$OUT/images/image-$2.tar" "$dir/files/image.tar"  # жёсткая ссылка — без лишней копии на диске
  find "$dir" -name .DS_Store -delete
}
pack "Mac (M1–M5)" arm64 mac
pack "Mac (Intel)" amd64 mac
pack "Windows"     amd64 windows

# Windows: .bat — с переводом строк CRLF, .ps1 — UTF-8 с BOM (иначе Windows PowerShell путает русские буквы)
W="$OUT/Конвертер Фоксфорда — Windows $VERSION"
for f in "$W"/*.bat; do perl -pi -e 's/\r?\n/\r\n/' "$f"; done
for f in "$W"/files/*.ps1; do perl -0777 -pi -e 's/\r?\n/\r\n/g; $_ = "\xEF\xBB\xBF$_"' "$f"; done

# ---------- архивы (zip с UTF-8 именами и правами на запуск для Mac)
python3 - "$OUT" <<'PY'
import os, sys, zipfile
out = sys.argv[1]
for name in sorted(os.listdir(out)):
    src = os.path.join(out, name)
    if not name.startswith('Конвертер') or not os.path.isdir(src): continue
    with zipfile.ZipFile(src + '.zip', 'w', zipfile.ZIP_DEFLATED, compresslevel=1, allowZip64=True) as z:
        for root, _, files in os.walk(src):
            for f in sorted(files):
                p = os.path.join(root, f)
                z.write(p, os.path.relpath(p, out),
                        compress_type=zipfile.ZIP_STORED if f.endswith('.tar') else zipfile.ZIP_DEFLATED)
    print('  %s.zip — %.0f МБ' % (name, os.path.getsize(src + '.zip') / 2**20))
PY
rm -rf "$OUT/images"
echo "Готово: $OUT"
