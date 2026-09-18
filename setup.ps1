param(
    [string]$Python = "D:\ProgramData\anaconda3\python.exe"
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

if (-not (Test-Path -LiteralPath $Python)) {
    $Python = (Get-Command python -ErrorAction Stop).Source
}

if (-not (Test-Path -LiteralPath ".venv\Scripts\python.exe")) {
    & $Python -m venv --system-site-packages .venv
}

& ".venv\Scripts\python.exe" -c "import pdfplumber, pypdf, win32com.client, tkinter; print('依赖检查通过')"
Write-Host "虚拟环境已就绪：$projectRoot\.venv"
Write-Host "运行图形界面：.\run.bat"
