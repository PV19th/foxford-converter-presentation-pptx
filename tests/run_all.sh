#!/usr/bin/env bash
# Прогоняет все презентации из test_inputs/old через конвертер в Docker.
# Результат: outputs/tests/index.html (превью рядом с эталоном методиста).
# Параметры передаются дальше: --only 5.21, --save-baseline
set -euo pipefail
cd "$(dirname "$0")/.."
docker compose build
docker compose run --rm -v "$PWD:/srv" app python3 tests/run_tests.py "$@"
