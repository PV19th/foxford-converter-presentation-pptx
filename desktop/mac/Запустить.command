#!/bin/bash
# Конвертер Фоксфорда на Mac: запускает Docker, ставит нужную версию конвертера и открывает его в браузере.
cd "$(dirname "$0")/files" 2>/dev/null || { echo "Не найдена папка files рядом с этим файлом. Распакуйте архив целиком."; read -r _; exit 1; }

VERSION=$(cat version.txt)
IMAGE="foxford-converter-desktop:$VERSION"
NAME=foxford-converter-desktop
URL=http://localhost:8765

say()  { echo; echo "  $*"; }
fail() { say "ОШИБКА: $*"; say "Если не получается — запустите «Лог для автора» и отправьте файл с рабочего стола автору."; echo; read -r -p "  Нажмите Enter, чтобы закрыть окно. " _; exit 1; }
docker_ok() { docker info >/dev/null 2>&1; }

# docker ставят Docker Desktop (/usr/local/bin, ~/.docker/bin) или Rancher Desktop (~/.rd/bin)
export PATH="$PATH:/usr/local/bin:/opt/homebrew/bin:$HOME/.docker/bin:$HOME/.rd/bin:/Applications/Docker.app/Contents/Resources/bin"
command -v docker >/dev/null || fail "Не найден Docker. Установите Docker Desktop (см. «Инструкция») и запустите снова."

say "Конвертер Фоксфорда, версия $VERSION"
if ! docker_ok; then
  say "Запускаю Docker… (до 3 минут)"
  open -a Docker 2>/dev/null || open -a "Rancher Desktop" 2>/dev/null
  for _ in $(seq 1 90); do docker_ok && break; sleep 2; done
  docker_ok || fail "Docker не запустился. Откройте Docker Desktop вручную, дождитесь надписи «Engine running» и запустите снова."
fi

if ! docker image inspect "$IMAGE" >/dev/null 2>&1; then
  say "Устанавливаю эту версию в Docker. Это 1–3 минуты и только при первом запуске…"
  docker load -i image.tar || fail "Не удалось установить конвертер в Docker."
fi

if [ "$(docker inspect -f '{{.Config.Image}}' $NAME 2>/dev/null)" != "$IMAGE" ] ||
   [ "$(docker inspect -f '{{.State.Running}}' $NAME 2>/dev/null)" != true ]; then
  docker rm -f $NAME >/dev/null 2>&1
  docker run -d --name $NAME --restart unless-stopped -p 127.0.0.1:8765:8000 \
    --env-file settings.env -v foxford-converter-data:/srv/data "$IMAGE" >/dev/null ||
    fail "Не удалось запустить конвертер. Возможно, порт 8765 занят другой программой."
fi

# старые версии больше не нужны — освобождаем место на диске
for old in $(docker images foxford-converter-desktop --format '{{.Repository}}:{{.Tag}}'); do
  [ "$old" != "$IMAGE" ] && docker rmi "$old" >/dev/null 2>&1
done

say "Жду, пока конвертер запустится…"
for _ in $(seq 1 90); do curl -s -o /dev/null "$URL/" && break; sleep 1; done
curl -s -o /dev/null "$URL/" || fail "Конвертер не ответил за 90 секунд."

open "$URL"
say "Готово: конвертер открыт в браузере — $URL"
say "Это окно можно закрыть. Конвертер работает в фоне, пока включён Docker."
echo
