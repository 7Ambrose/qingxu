param([string]$Version='0.1.0')
$ErrorActionPreference='Stop'
if($Version -notmatch '^\d+\.\d+\.\d+([-.][a-zA-Z0-9.-]+)?$'){throw 'Invalid release version'}
$root=Split-Path -Parent $PSScriptRoot
$app=Join-Path $root '今日日程'
$compiler=Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if(-not (Test-Path -LiteralPath $compiler)){$compiler=Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe'}
if(-not (Test-Path -LiteralPath $compiler)){throw '.NET Framework compiler not found'}
& $compiler /nologo /target:winexe /platform:anycpu /reference:System.Windows.Forms.dll ("/win32icon:"+(Join-Path $app 'icon.ico')) ("/out:"+(Join-Path $app '启动青序.exe')) (Join-Path $app '启动青序.cs')
if($LASTEXITCODE -ne 0){throw 'PowerShell launcher compilation failed'}
$dist=Join-Path $root 'dist';[void][IO.Directory]::CreateDirectory($dist)
$name='Qingxu-PowerShell-v'+$Version;$stage=Join-Path (Join-Path $dist ('build-'+[guid]::NewGuid().ToString('N'))) $name
[void][IO.Directory]::CreateDirectory((Join-Path $stage '今日日程'))
[void][IO.Directory]::CreateDirectory((Join-Path $stage 'templates'))
[void][IO.Directory]::CreateDirectory((Join-Path $stage 'outputs\study-plan'))
$files=@('controls.ps1','pages.ps1','desktop.ps1','今日日程.ps1','view.xaml','taskbar.cs','read_plan.py','reader_worker.ps1','excel_links.py','course_links.py','sync_review.py','conflicts.ps1','sync_worker.ps1','write_batch.mjs','write_plan.mjs','icon.ico','启动青序.exe')
foreach($file in $files){Copy-Item -LiteralPath (Join-Path $app $file) -Destination (Join-Path $stage ('今日日程\'+$file))}
Copy-Item -LiteralPath (Join-Path $root 'README.md') -Destination $stage
Copy-Item -LiteralPath (Join-Path $root 'templates\计划模板.xlsx') -Destination (Join-Path $stage 'templates')
Copy-Item -LiteralPath (Join-Path $root '创建桌面快捷方式.ps1') -Destination $stage
$manifest=Get-ChildItem -LiteralPath $stage -Recurse -File|ForEach-Object { (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash+'  '+$_.FullName.Substring($stage.Length+1) }
[IO.File]::WriteAllLines((Join-Path $stage 'SHA256SUMS.txt'),$manifest,[Text.Encoding]::UTF8)
$zip=Join-Path $dist ($name+'.zip')
Compress-Archive -LiteralPath $stage -DestinationPath $zip -CompressionLevel Optimal -Force
Get-FileHash -LiteralPath $zip -Algorithm SHA256|Format-List
