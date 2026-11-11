# Собирает журнал работы конвертера в файл на рабочем столе — его нужно отправить автору.
Set-Location -LiteralPath $PSScriptRoot
$env:Path = "$env:Path;$env:ProgramFiles\Docker\Docker\resources\bin;$env:LOCALAPPDATA\Programs\Rancher Desktop\resources\resources\win32\bin"
$Name = 'foxford-converter-desktop'
$Out = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Конвертер Фоксфорда — лог.txt'

$lines = @(
    "Версия: $((Get-Content -LiteralPath 'version.txt' -Raw -ErrorAction SilentlyContinue))".Trim()
    "Время: $(Get-Date)"
    "Компьютер: Windows $([Environment]::OSVersion.Version)"
    ''
)
if (-not (Get-Command docker -ErrorAction SilentlyContinue)) { $lines += 'Docker не найден' }
else {
    # через cmd — чтобы сообщения Docker попали в файл обычным текстом
    $lines += '== Docker'
    $lines += cmd /c "docker version --format {{.Server.Version}} 2>&1"
    $lines += cmd /c "docker ps -a --filter name=$Name 2>&1"
    $lines += ''; $lines += '== Журнал конвертера'
    $lines += cmd /c "docker exec $Name tail -n 3000 /srv/data/logs/app.log 2>&1"
    $lines += ''; $lines += '== Журнал контейнера'
    $lines += cmd /c "docker logs --tail 300 $Name 2>&1"
}
$lines | Out-File -LiteralPath $Out -Encoding utf8

Write-Host ''
Write-Host '  Готово: на рабочем столе файл «Конвертер Фоксфорда — лог.txt». Отправьте его автору.'
Write-Host ''
Read-Host '  Нажмите Enter, чтобы закрыть окно' | Out-Null
