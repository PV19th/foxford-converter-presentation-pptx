# Конвертер Фоксфорда на Windows: запускает Docker, ставит нужную версию конвертера и открывает его в браузере.
# Вызывается из «Запустить.bat».
Set-Location -LiteralPath $PSScriptRoot
$Version = (Get-Content -LiteralPath 'version.txt' -Raw).Trim()
$Image = "foxford-converter-desktop:$Version"
$Name = 'foxford-converter-desktop'
$Url = 'http://localhost:8765'

function Say($text) { Write-Host ''; Write-Host "  $text" }
function Fail($text) {
    Say "ОШИБКА: $text"
    Say 'Если не получается — запустите «Лог для автора» и отправьте файл с рабочего стола автору.'
    Write-Host ''
    Read-Host '  Нажмите Enter, чтобы закрыть окно' | Out-Null
    exit 1
}
function DockerOk { & docker info *> $null; return ($LASTEXITCODE -eq 0) }

# docker ставят Docker Desktop или Rancher Desktop
$env:Path = "$env:Path;$env:ProgramFiles\Docker\Docker\resources\bin;$env:LOCALAPPDATA\Programs\Rancher Desktop\resources\resources\win32\bin"
if (-not (Get-Command docker -ErrorAction SilentlyContinue)) {
    Fail 'Не найден Docker. Установите Docker Desktop (см. «Инструкция») и запустите снова.'
}

Say "Конвертер Фоксфорда, версия $Version"
if (-not (DockerOk)) {
    Say 'Запускаю Docker… (до 3 минут)'
    $dd = "$env:ProgramFiles\Docker\Docker\Docker Desktop.exe"
    $rd = "$env:LOCALAPPDATA\Programs\Rancher Desktop\Rancher Desktop.exe"
    if (Test-Path -LiteralPath $dd) { Start-Process -FilePath $dd } elseif (Test-Path -LiteralPath $rd) { Start-Process -FilePath $rd }
    for ($i = 0; $i -lt 90; $i++) { if (DockerOk) { break }; Start-Sleep -Seconds 2 }
    if (-not (DockerOk)) { Fail 'Docker не запустился. Откройте Docker Desktop вручную, дождитесь надписи «Engine running» и запустите снова.' }
}

& docker image inspect $Image *> $null
if ($LASTEXITCODE -ne 0) {
    Say 'Устанавливаю эту версию в Docker. Это 1–3 минуты и только при первом запуске…'
    & docker load -i image.tar
    if ($LASTEXITCODE -ne 0) { Fail 'Не удалось установить конвертер в Docker.' }
}

$current = & docker inspect -f '{{.Config.Image}}' $Name 2> $null
$running = & docker inspect -f '{{.State.Running}}' $Name 2> $null
if ($current -ne $Image -or $running -ne 'true') {
    & docker rm -f $Name *> $null
    & docker run -d --name $Name --restart unless-stopped -p 127.0.0.1:8765:8000 --env-file settings.env -v foxford-converter-data:/srv/data $Image | Out-Null
    if ($LASTEXITCODE -ne 0) { Fail 'Не удалось запустить конвертер. Возможно, порт 8765 занят другой программой.' }
}

# старые версии больше не нужны — освобождаем место на диске
& docker images foxford-converter-desktop --format '{{.Repository}}:{{.Tag}}' |
    Where-Object { $_ -and $_ -ne $Image } | ForEach-Object { & docker rmi $_ *> $null }

Say 'Жду, пока конвертер запустится…'
$ok = $false
for ($i = 0; $i -lt 90; $i++) {
    try { Invoke-WebRequest -UseBasicParsing -TimeoutSec 2 -Uri "$Url/" | Out-Null; $ok = $true; break }
    catch { Start-Sleep -Seconds 1 }
}
if (-not $ok) { Fail 'Конвертер не ответил за 90 секунд.' }

Start-Process $Url
Say "Готово: конвертер открыт в браузере — $Url"
Say 'Окно закроется само через 10 секунд. Конвертер работает в фоне, пока включён Docker.'
Start-Sleep -Seconds 10
