#!/bin/bash
# Собирает версию «в браузере» (сайт без серверной части) в dist/browser/. Подробности — browser/README.md.
# Запуск из корня проекта:  ./browser/build.sh   (первый раз скачивает Pyodide, ~15 МБ)
# Контакт для «Сообщить автору» берётся из .env (на GitHub — из переменной SUPPORT_CONTACT). Папка dist/ в git не попадает.
# Эту же команду запускает GitHub при публикации сайта (.github/workflows/browser-site.yml).
set -euo pipefail
cd "$(dirname "$0")/.."
docker compose build
docker compose run --rm -T -e SUPPORT_CONTACT -v "$PWD:/srv" app python3 browser/build.py
