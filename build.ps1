$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot
.\.venv\Scripts\python -m pip install -r requirements-dev.txt
if ($LASTEXITCODE -ne 0) { throw "pip install failed" }
.\.venv\Scripts\python tools\make_assets.py
if ($LASTEXITCODE -ne 0) { throw "asset generation failed" }
.\.venv\Scripts\python -m PyInstaller --noconfirm --clean --onefile --windowed `
  --name "PDF Converter" --icon assets\icon.ico --splash assets\splash.png `
  --add-data "app\ui;app\ui" --collect-all webview run.py
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed" }
