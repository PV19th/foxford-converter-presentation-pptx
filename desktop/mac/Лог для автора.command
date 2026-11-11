#!/bin/bash
# Собирает журнал работы конвертера в файл на рабочем столе — его нужно отправить автору.
cd "$(dirname "$0")/files" 2>/dev/null
export PATH="$PATH:/usr/local/bin:/opt/homebrew/bin:$HOME/.docker/bin:$HOME/.rd/bin:/Applications/Docker.app/Contents/Resources/bin"
OUT="$HOME/Desktop/Конвертер Фоксфорда — лог.txt"
NAME=foxford-converter-desktop
{
  echo "Версия: $(cat version.txt 2>/dev/null)"
  echo "Время: $(date)"
  echo "Компьютер: Mac, $(uname -m), macOS $(sw_vers -productVersion 2>/dev/null)"
  echo
  if ! command -v docker >/dev/null; then echo "Docker не найден"
  elif ! docker info >/dev/null 2>&1; then echo "Docker не запущен"
  else
    echo "== Docker"; docker version --format '{{.Server.Version}}'; docker ps -a --filter name=$NAME
    echo; echo "== Журнал конвертера"; docker exec $NAME tail -n 3000 /srv/data/logs/app.log 2>&1
    echo; echo "== Журнал контейнера"; docker logs --tail 300 $NAME 2>&1
  fi
} > "$OUT" 2>&1
echo
echo "  Готово: на рабочем столе файл «Конвертер Фоксфорда — лог.txt»."
echo "  Отправьте его автору. Это окно можно закрыть."
echo
