$ErrorActionPreference='Stop'
$qa=Join-Path (Split-Path -Parent $PSScriptRoot) 'outputs\study-plan\qa'
$fixture=Join-Path $qa ('launcher-'+[guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($fixture)
$launcher=Join-Path $fixture 'launcher.exe'
Copy-Item -LiteralPath (Join-Path $PSScriptRoot '启动青序.exe') -Destination $launcher
$scriptPath=Join-Path $fixture '今日日程.ps1'
$logPath=Join-Path $fixture '启动错误.txt'
function Invoke-TestLauncher([string]$policy){
 # 通过父进程环境变量传递给子进程：本机 PowerShell 5.1 下 ProcessStartInfo.EnvironmentVariables 会返回空引用
 $previous=$env:PSExecutionPolicyPreference;$env:PSExecutionPolicyPreference=$policy
 try{
  $start=New-Object Diagnostics.ProcessStartInfo
  $start.FileName=$launcher;$start.Arguments='--diagnose';$start.UseShellExecute=$false;$start.CreateNoWindow=$true
  $process=[Diagnostics.Process]::Start($start)
  if(-not $process.WaitForExit(15000)){$process.Kill();throw 'Launcher did not exit within 15 seconds'}
  $code=$process.ExitCode;$process.Dispose();return $code
 }finally{$env:PSExecutionPolicyPreference=$previous}
}
try{
 if((Invoke-TestLauncher 'RemoteSigned') -eq 0 -or -not (Test-Path -LiteralPath $logPath)){throw 'Missing script failed silently'}
 [IO.File]::WriteAllText($scriptPath,"Write-Output 'fixture-runtime-error';exit 7",(New-Object Text.UTF8Encoding $true))
 if((Invoke-TestLauncher 'RemoteSigned') -ne 7 -or [IO.File]::ReadAllText($logPath) -notmatch 'fixture-runtime-error'){throw 'Runtime error was not captured with its exit code'}
 [IO.File]::WriteAllText($scriptPath,'exit 0',(New-Object Text.UTF8Encoding $true))
 if((Invoke-TestLauncher 'RemoteSigned') -ne 0){throw 'Valid script failed to launch'}
 if((Invoke-TestLauncher 'Restricted') -eq 0 -or [IO.File]::ReadAllText($logPath) -notmatch 'PSSecurityException|UnauthorizedAccess|ExecutionPolicy'){throw 'Execution policy rejection was not logged'}
 'PASS: launcher reports missing scripts, runtime failures and policy rejection; valid scripts exit successfully.'
}finally{
 Get-ChildItem -LiteralPath $fixture -File|Remove-Item -Force
 Remove-Item -LiteralPath $fixture
}
