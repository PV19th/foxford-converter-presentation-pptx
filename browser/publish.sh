#!/bin/bash
# Публикует собранную версию «в браузере» на GitHub Pages: копирует dist/browser/ в папку отдельного публичного
# репозитория (его копия на этом компьютере, склонированная через GitHub Desktop). Основной репозиторий остаётся закрытым.
#   первый раз:  ./browser/publish.sh "/путь/к/папке публичного репозитория"   (путь запоминается)
#   дальше:      ./browser/publish.sh
# После команды — Commit и Push в GitHub Desktop. Подробности — browser/README.md, «Публикация».
set -euo pipefail
cd "$(dirname "$0")/.."
SAVED="dist/browser-publish-path.txt"
TARGET="${1:-$(cat "$SAVED" 2>/dev/null || true)}"
TARGET="${TARGET%/}"

[ -f dist/browser/meta.json ] || { echo "Сначала соберите версию: ./browser/build.sh"; exit 1; }
[ -n "$TARGET" ] || { echo "Укажите папку публичного репозитория: ./browser/publish.sh \"/путь/к/папке\""; exit 1; }
[ -d "$TARGET/.git" ] || { echo "«$TARGET» — не папка репозитория. Склонируйте публичный репозиторий через GitHub Desktop и укажите его папку."; exit 1; }

# Защита от ошибки в пути: команда заменяет содержимое папки. Поэтому в ней должно быть пусто
# (кроме служебных файлов) или уже лежать прошлая публикация.
KEEP='\.git|\.gitattributes|\.gitignore|README\.md|LICENSE|CNAME|\.DS_Store'
if [ ! -f "$TARGET/pyodide/pyodide.mjs" ]; then
  extra=$(ls -A "$TARGET" | grep -vxE "$KEEP" || true)
  if [ -n "$extra" ]; then
    echo "В папке «$TARGET» лежат посторонние файлы — похоже, это не репозиторий сайта:"; echo "$extra" | head -5
    exit 1
  fi
fi

# CNAME (свой домен; его создаёт GitHub) и README репозитория не трогаем
rsync -a --delete --exclude .git --exclude .gitattributes --exclude .gitignore --exclude README.md --exclude LICENSE \
      --exclude CNAME --exclude .DS_Store dist/browser/ "$TARGET/"
echo "$TARGET" > "$SAVED"
VERSION=$(python3 -c "import json; print(json.load(open('dist/browser/meta.json'))['version'])")
echo "Скопировано в «$TARGET», версия $VERSION."
echo "Теперь в GitHub Desktop: выберите этот репозиторий → Commit → Push. Сайт обновится через 1–2 минуты."
