#!/bin/bash
# Собирает версию «в браузере» (сайт без серверной части) в dist/browser/. Подробности — browser/README.md.
# Запуск из корня проекта:  ./browser/build.sh   (первый раз скачивает Pyodide, ~15 МБ)
# Контакт для «Сообщить автору» берётся из .env. Папка dist/ в git не попадает.
set -euo pipefail
cd "$(dirname "$0")/.."
docker compose build
docker compose run --rm -v "$PWD:/srv" app python3 browser/build.py
