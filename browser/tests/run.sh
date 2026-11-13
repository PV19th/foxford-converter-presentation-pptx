#!/bin/bash
# Проверяет версию «в браузере» на тестовых презентациях: отпечаток результата должен совпасть с эталоном сайта
# (tests/baseline/). Сначала соберите её: ./browser/build.sh
# Параметры: [профиль …] [--only урок], например  ./browser/tests/run.sh foxford_5-11_v01_26 --only 5.21
set -euo pipefail
cd "$(dirname "$0")/../.."
[ -f dist/browser/meta.json ] || { echo "Сначала соберите версию: ./browser/build.sh"; exit 1; }
docker run --rm -v "$PWD:/srv" -w /srv node:24-slim node browser/tests/regress.mjs "$@"
