$ErrorActionPreference = 'Stop'
$exe = Join-Path $PSScriptRoot '青序.exe'
if (-not (Test-Path -LiteralPath $exe)) {
    $exe = Join-Path (Split-Path -Parent $PSScriptRoot) 'outputs\native-win-x64\青序.exe'
}
if (-not (Test-Path -LiteralPath $exe)) {
    throw '找不到 青序.exe。请先完整解压发布包，或在源码目录运行 scripts\build_native.ps1 -Publish。'
}
$target = Join-Path ([Environment]::GetFolderPath('Desktop')) '青序-原生版.lnk'
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($target)
$shortcut.TargetPath = $exe
$shortcut.WorkingDirectory = Split-Path -Parent $exe
$shortcut.IconLocation = $exe + ',0'
$shortcut.Description = '青序 · C# 原生版（不依赖 Python / Node）'
$shortcut.Save()
Write-Host ('已创建：' + $target)
