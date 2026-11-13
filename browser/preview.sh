#!/bin/bash
# Открывает собранную версию «в браузере» на этом компьютере: http://localhost:8766  (остановить — Ctrl+C).
# Просто открыть dist/browser/index.html двойным щелчком нельзя: браузеру нужен адрес http://…
set -euo pipefail
cd "$(dirname "$0")/.."
[ -f dist/browser/index.html ] || { echo "Сначала соберите версию: ./browser/build.sh"; exit 1; }
echo "Откройте в браузере: http://localhost:8766"
docker run --rm -p 127.0.0.1:8766:8766 -v "$PWD/dist/browser:/site:ro" -w /site foxford-converter python3 -m http.server 8766
