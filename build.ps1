$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$env:PYTHONUTF8 = '1'
$projectDir = $PSScriptRoot
$pythonExe = Join-Path $projectDir '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $pythonExe)) {
    throw '请先运行 python -m venv .venv 并安装 requirements-dev.txt。'
}
Push-Location -LiteralPath $projectDir
try {
    & $pythonExe -m unittest discover -s tests -v
    if ($LASTEXITCODE -ne 0) { throw '测试未通过，已停止打包。' }
    & $pythonExe -m PyInstaller --noconfirm --distpath 'dist\v5' --workpath 'build\v5' 'qqmusic_converter.spec'
    if ($LASTEXITCODE -ne 0) { throw '打包失败。' }
    $optionalFfmpeg = Join-Path $projectDir 'ffmpeg.exe'
    if (Test-Path -LiteralPath $optionalFfmpeg) {
        Copy-Item -LiteralPath $optionalFfmpeg -Destination (Join-Path $projectDir 'dist\v5\ffmpeg.exe')
    }
    Write-Host "已生成: $projectDir\dist\v5\QQMusicConverter.exe"
} finally {
    Pop-Location
}
