# Установка office-docs на Windows.
#
#   .\install.ps1                    # ядро
#   .\install.ps1 --with-markitdown  # + markitdown
#
# Вся логика в install.py; этот файл только находит Python и передаёт аргументы.
# Если PowerShell-скрипты запускать не хочется, то же самое делает:
#   py -3 install.py

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

$python = $null
foreach ($candidate in @("py", "python")) {
    if (Get-Command $candidate -ErrorAction SilentlyContinue) { $python = $candidate; break }
}
if (-not $python) {
    Write-Error @"
Python не найден. Нужен Python 3.10+:
  winget install Python.Python.3.12
после чего запустите:  py -3 install.py
"@
    exit 1
}

& $python "$root\install.py" @args
exit $LASTEXITCODE
