param(
    [switch]$SkipInstall
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

$templatePath = Join-Path $projectRoot "模板.xls"
if (-not (Test-Path -LiteralPath $templatePath)) {
    throw "缺少本地模板：$templatePath。模板不会提交到 Git，请先放入项目根目录。"
}

$bootstrapPython = Join-Path $projectRoot ".venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $bootstrapPython)) {
    & (Join-Path $projectRoot "setup.ps1")
}
if (-not (Test-Path -LiteralPath $bootstrapPython)) {
    throw "项目虚拟环境创建失败：$bootstrapPython"
}

$buildVenv = Join-Path $projectRoot ".venv-build"
$python = Join-Path $buildVenv "Scripts\python.exe"
if (-not (Test-Path -LiteralPath $python)) {
    & $bootstrapPython -m venv $buildVenv
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

if (-not $SkipInstall) {
    & $python -m pip install --disable-pip-version-check -r requirements.txt -r requirements-build.txt
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
} else {
    & $python -c "import PyInstaller, pdfplumber, pypdf, win32com.client"
    if ($LASTEXITCODE -ne 0) {
        throw "构建环境依赖不完整，请不要使用 -SkipInstall"
    }
}

$releaseDir = Join-Path $projectRoot "release"
$workDir = Join-Path $projectRoot ".build\pyinstaller"
$specDir = Join-Path $projectRoot ".build\spec"

foreach ($directory in @($workDir, $specDir)) {
    if (Test-Path -LiteralPath $directory) {
        Remove-Item -LiteralPath $directory -Recurse -Force
    }
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
}
New-Item -ItemType Directory -Path $releaseDir -Force | Out-Null

$exePath = Join-Path $releaseDir "HITFapiao.exe"
$zipPath = Join-Path $releaseDir "HITFapiao-Windows-x64.zip"
foreach ($file in @($exePath, $zipPath)) {
    if (Test-Path -LiteralPath $file) {
        Remove-Item -LiteralPath $file -Force
    }
}

& $python -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --noupx `
    --name HITFapiao `
    --distpath $releaseDir `
    --workpath $workDir `
    --specpath $specDir `
    --hidden-import pythoncom `
    --hidden-import pywintypes `
    --hidden-import win32com.client `
    --hidden-import win32timezone `
    app.py
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Copy-Item -LiteralPath $templatePath -Destination $releaseDir -Force
Copy-Item -LiteralPath (Join-Path $projectRoot "README.md") -Destination $releaseDir -Force

& $python (Join-Path $projectRoot "tools\package_release.py") `
    --release-dir $releaseDir `
    --output $zipPath
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

$exeHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $exePath).Hash
$zipHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $zipPath).Hash
Write-Host "Release 构建完成："
Write-Host "  EXE: $exePath"
Write-Host "  ZIP: $zipPath"
Write-Host "  EXE SHA256: $exeHash"
Write-Host "  ZIP SHA256: $zipHash"
