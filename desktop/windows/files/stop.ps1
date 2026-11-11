# Останавливает конвертер — освобождает память. Снова включить — «Запустить».
$env:Path = "$env:Path;$env:ProgramFiles\Docker\Docker\resources\bin;$env:LOCALAPPDATA\Programs\Rancher Desktop\resources\resources\win32\bin"
& docker stop foxford-converter-desktop *> $null
Write-Host ''
Write-Host '  Конвертер остановлен. Docker Desktop можно закрыть отдельно (значок кита у часов → Quit Docker Desktop).'
Write-Host '  Окно закроется само через 5 секунд.'
Start-Sleep -Seconds 5
