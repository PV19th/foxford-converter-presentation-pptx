#!/bin/bash
# Останавливает конвертер — освобождает память. Снова включить — «Запустить».
export PATH="$PATH:/usr/local/bin:/opt/homebrew/bin:$HOME/.docker/bin:$HOME/.rd/bin:/Applications/Docker.app/Contents/Resources/bin"
docker stop foxford-converter-desktop >/dev/null 2>&1
echo
echo "  Конвертер остановлен. Это окно можно закрыть."
echo "  Docker Desktop можно закрыть отдельно (значок кита в строке меню → Quit Docker Desktop)."
echo
