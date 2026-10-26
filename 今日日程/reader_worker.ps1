function Stop-AgendaReader {
 $process=$script:readerProcess;$script:readerProcess=$null
 if($process){
  try{$process.StandardInput.Close();if(-not $process.WaitForExit(200)){$process.Kill();$process.WaitForExit()}}catch{}
  finally{$process.Dispose()}
 }
 $script:readerErrorTask=$null;$script:readerBook=$null
}
function Invoke-AgendaReader([string]$date,[switch]$Responsive){
 if($script:readerRequestBusy){throw '已有 Excel 读取请求进行中。'}
 $script:readerRequestBusy=$true
 try{
  if($script:readerProcess -and ($script:readerProcess.HasExited -or $script:readerBook -ne $script:planPath)){Stop-AgendaReader}
  if(-not $script:readerProcess){
   if($script:readerPath.Contains('"') -or $script:planPath.Contains('"')){throw '工作簿路径无效。'}
   $info=New-Object Diagnostics.ProcessStartInfo;$info.FileName=$script:pythonPath
   $info.Arguments='-X utf8 "'+$script:readerPath+'" --book "'+$script:planPath+'" --serve'
   $info.UseShellExecute=$false;$info.CreateNoWindow=$true;$info.RedirectStandardInput=$true;$info.RedirectStandardOutput=$true;$info.RedirectStandardError=$true;$info.StandardOutputEncoding=[Text.Encoding]::UTF8
   $script:readerProcess=[Diagnostics.Process]::Start($info);$script:readerBook=$script:planPath
   $script:readerErrorTask=$script:readerProcess.StandardError.ReadToEndAsync()
  }
  $request=(@{date=$date}|ConvertTo-Json -Compress)+"`n";$bytes=[Text.Encoding]::UTF8.GetBytes($request)
  $script:readerProcess.StandardInput.BaseStream.Write($bytes,0,$bytes.Length);$script:readerProcess.StandardInput.BaseStream.Flush()
  $response=$script:readerProcess.StandardOutput.ReadLineAsync();$clock=[Diagnostics.Stopwatch]::StartNew()
  while(-not $response.IsCompleted){
   if($clock.Elapsed.TotalSeconds -gt 30){throw '读取 Excel 超时，本地修改仍保留，可稍后重试。'}
   if($Responsive){[Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action]{},[Windows.Threading.DispatcherPriority]::Background)}
   Start-Sleep -Milliseconds 5
  }
  $line=$response.GetAwaiter().GetResult()
  if(-not $line){throw 'Excel 读取进程已退出，可点击刷新重试。'}
  return ($line|ConvertFrom-Json)
 }catch{Stop-AgendaReader;throw}
 finally{$script:readerRequestBusy=$false}
}
