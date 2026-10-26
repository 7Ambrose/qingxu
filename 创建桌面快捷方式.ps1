$ErrorActionPreference='Stop'
$exe=Join-Path $PSScriptRoot '今日日程\启动青序.exe'
if(-not (Test-Path -LiteralPath $exe)){throw '请先完整解压发布包，或运行 scripts/build_release.ps1'}
$shell=New-Object -ComObject WScript.Shell
$shortcut=$shell.CreateShortcut((Join-Path ([Environment]::GetFolderPath('Desktop')) '青序-PowerShell版.lnk'))
$shortcut.TargetPath=$exe;$shortcut.WorkingDirectory=Split-Path -Parent $exe
$shortcut.IconLocation=(Join-Path $PSScriptRoot '今日日程\icon.ico')+',0'
$shortcut.Description='青序 PowerShell 旧版（不是原生实验版）';$shortcut.Save()
