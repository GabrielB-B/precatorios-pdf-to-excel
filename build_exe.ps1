$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

python -m pip install --upgrade pyinstaller
python .\generate_brand_assets.py

pyinstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name ExtratorPrecatorios `
    --icon assets\extrator_precatorios.ico `
    --add-data "assets;assets" `
    --exclude-module torch `
    --exclude-module torchvision `
    --exclude-module torchaudio `
    --exclude-module tensorflow `
    --exclude-module transformers `
    --exclude-module sklearn `
    --exclude-module scipy `
    --exclude-module pytest `
    --exclude-module sympy `
    --exclude-module sqlalchemy `
    --exclude-module jinja2 `
    --exclude-module pygments `
    --exclude-module rich `
    --exclude-module httpx `
    --exclude-module anyio `
    --exclude-module trio `
    --exclude-module matplotlib `
    --exclude-module seaborn `
    app_gui.py

Write-Host ""
Write-Host "EXE gerado em: $projectRoot\\dist\\ExtratorPrecatorios.exe"
