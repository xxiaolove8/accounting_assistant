$ErrorActionPreference = 'Stop'
$projectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $projectDir

if (-not (Test-Path -LiteralPath '.venv\Scripts\python.exe')) {
    py -m venv .venv
}

& '.\.venv\Scripts\python.exe' -m pip install --disable-pip-version-check -r requirements-build.txt
if ($LASTEXITCODE -ne 0) { throw '安装构建依赖失败。' }

& '.\.venv\Scripts\python.exe' -m PyInstaller `
    --noconfirm --clean --onefile --windowed `
    --name ExcelFilterAssistant `
    --distpath dist --workpath build `
    main.py
if ($LASTEXITCODE -ne 0) { throw 'EXE 打包失败。' }

Write-Host "已生成：$projectDir\dist\ExcelFilterAssistant.exe"
