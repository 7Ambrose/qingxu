param(
 [Parameter(Mandatory=$true)][string]$Book,
 [Parameter(Mandatory=$true)][string]$Python,
 [Parameter(Mandatory=$true)][string]$InputFile,
 [Parameter(Mandatory=$true)][string]$ResultFile
)
$ErrorActionPreference='Stop'
$reader=Join-Path $PSScriptRoot 'read_plan.py'
function Invoke-Reader([string]$Mode,[string]$Body=''){
 $info=New-Object Diagnostics.ProcessStartInfo;$info.FileName=$Python
 $info.Arguments='-X utf8 "'+$reader+'" --book "'+$Book+'" '+$Mode
 $info.UseShellExecute=$false;$info.CreateNoWindow=$true;$info.RedirectStandardOutput=$true;$info.RedirectStandardError=$true;$info.RedirectStandardInput=$true;$info.StandardOutputEncoding=[Text.Encoding]::UTF8
 $process=[Diagnostics.Process]::Start($info)
 # 两路管道并行抽干，避免子进程写满 stderr 缓冲区时与父进程互相等待
 $outTask=$process.StandardOutput.ReadToEndAsync();$errTask=$process.StandardError.ReadToEndAsync()
 $exitCode=0;$output='';$errorOutput=''
 try{
  if($Body){$bytes=[Text.Encoding]::UTF8.GetBytes($Body);$process.StandardInput.BaseStream.Write($bytes,0,$bytes.Length)}
  $process.StandardInput.Close();$process.WaitForExit();$exitCode=$process.ExitCode
  $output=$outTask.GetAwaiter().GetResult();$errorOutput=$errTask.GetAwaiter().GetResult()
 }finally{if(-not $process.HasExited){$process.Kill()};$process.Dispose()}
 if($exitCode -ne 0){throw '同步工具启动失败，请检查 Python 和运行依赖。'};$result=$output|ConvertFrom-Json;if($result.error){$script:readerErrorCode=$result.code;throw $result.error};return $result
}
try{
 $check=Invoke-Reader '--check'
 $input=Get-Content -LiteralPath $InputFile -Raw -Encoding UTF8|ConvertFrom-Json
 $body=$input|ConvertTo-Json -Depth 10 -Compress
 $stamp=Invoke-Reader '--stamp'
 # 版本号与内容摘要同时比对：低时间戳精度的磁盘上，仅靠 mtime 可能漏判外部改动
 $stale=([string]$stamp.version -ne [string]$input.version) -or ($input.digest -and [string]$stamp.digest -ne [string]$input.digest)
 if($stale){$response=@{ok=$false;code='VERSION_CONFLICT';error='Excel 已有更新，尚未写入。';review=(Invoke-Reader '--review-sync' $body)}}
 else{$result=Invoke-Reader '--batch-write' $body;$response=@{ok=$true;mappings=$result.mappings;version=$result.version;digest=$result.digest}}
}catch{
 $response=@{ok=$false;error=$_.Exception.Message}
 if($script:readerErrorCode -eq 'VERSION_CONFLICT'){
  try{$response.code='VERSION_CONFLICT';$response.review=Invoke-Reader '--review-sync' $body}catch{$response.error='Excel 又有更新或被占用，请保存并关闭后重试。'}
 }
}
$temp=$ResultFile+'.pending'
[IO.File]::WriteAllText($temp,($response|ConvertTo-Json -Depth 8 -Compress),[System.Text.UTF8Encoding]::new($false))
Move-Item -LiteralPath $temp -Destination $ResultFile -Force
