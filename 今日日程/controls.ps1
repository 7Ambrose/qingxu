function Invoke-Backend([string]$mode,[string]$body='',[switch]$Responsive){
 if($mode -match '^--range --date (\d{4}-\d{2}-\d{2})$'){return (Invoke-AgendaReader $Matches[1] -Responsive:$Responsive)}
 $info=New-Object Diagnostics.ProcessStartInfo;$info.FileName=$script:pythonPath
 $info.Arguments='-X utf8 "'+$script:readerPath+'" --book "'+$script:planPath+'" '+$mode
 $info.UseShellExecute=$false;$info.CreateNoWindow=$true;$info.RedirectStandardOutput=$true;$info.RedirectStandardError=$true;$info.RedirectStandardInput=$true;$info.StandardOutputEncoding=[Text.Encoding]::UTF8
 $p=[Diagnostics.Process]::Start($info)
 # 两路管道必须并行抽干：先把 stdout 读到结尾再读 stderr，子进程写满 stderr 缓冲区时会互相等待
 $outTask=$p.StandardOutput.ReadToEndAsync();$errTask=$p.StandardError.ReadToEndAsync()
 $exitCode=0;$raw='';$err=''
 try{
  if($body){$bytes=[Text.Encoding]::UTF8.GetBytes($body);$p.StandardInput.BaseStream.Write($bytes,0,$bytes.Length)}
  $p.StandardInput.Close()
  if($Responsive){
   # 解释器启动约 0.75 秒，这段等待必须让出消息循环，否则整个窗口在读取期间无响应
   while(-not $p.HasExited){
    [Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action]{},[Windows.Threading.DispatcherPriority]::Background)
    Start-Sleep -Milliseconds 10
   }
  }else{$p.WaitForExit()}
  $exitCode=$p.ExitCode
  $raw=$outTask.GetAwaiter().GetResult();$err=$errTask.GetAwaiter().GetResult()
 }finally{if(-not $p.HasExited){$p.Kill()};$p.Dispose()}
 if($exitCode -ne 0){throw $(if($err.Trim()){$err}else{'读取程序异常退出（代码 '+$exitCode+'）。'})}
 if(-not $raw.Trim()){throw '读取程序没有返回内容。'}
 return ($raw|ConvertFrom-Json)
}
. (Join-Path $PSScriptRoot 'reader_worker.ps1')
if(-not $script:cachePath){$script:cachePath=$script:planPath+'.agenda.json'}
if(-not $script:pendingOps){$script:pendingOps=@()}
function Set-LocalValue($object,[string]$name,$value){if($object.PSObject.Properties[$name]){$object.$name=$value}else{$object|Add-Member -NotePropertyName $name -NotePropertyValue $value}}
function Write-AgendaCache {
 $cache=@{migrations=@{weeklyUnschedule=1};date=$script:lastDate;version=$(if($script:baseVersion){$script:baseVersion}else{$script:data.version});digest=$(if($script:baseDigest){$script:baseDigest}else{$script:data.digest});data=$script:data;pendingOps=@($script:pendingOps)}|ConvertTo-Json -Depth 16
 $temp=$script:cachePath+'.pending';[IO.File]::WriteAllText($temp,$cache,[System.Text.UTF8Encoding]::new($false));Move-Item -LiteralPath $temp -Destination $script:cachePath -Force
}
function Add-PendingOperation([string]$id,[string]$date,$fields,$snapshot=$null){
 $existing=@($script:pendingOps|Where-Object {$_.id -eq $id -and $_.date -eq $date})
 if($existing.Count){foreach($k in $fields.Keys){if($existing[0].fields -is [System.Collections.IDictionary]){$existing[0].fields[$k]=$fields[$k]}else{Set-LocalValue $existing[0].fields $k $fields[$k]}};if($snapshot){Set-LocalValue $existing[0] 'snapshot' $snapshot}}
 else{$script:pendingOps+=@([PSCustomObject]@{id=$id;date=$date;version=$(if($script:baseVersion){$script:baseVersion}else{$script:data.version});fields=$fields;snapshot=$snapshot})}
}
function Remove-AgendaItem($item,[bool]$confirm=$true){
 Complete-AgendaCheckboxClick
 if($script:busy -or $item.id -eq 'new'){return $false}
 $unschedule=$item.id -like 'weekly:*' -and $script:page -ne 3
 if($confirm){$message=if($unschedule){'取消这一天的日期和时间安排，任务仍保留在本周任务中。确定取消安排吗？'}elseif($item.id -like 'class:*'){'只删除这节课的课内任务，保留课程本身。确定删除吗？'}else{'删除后将在下次同步时从 Excel 中移除。确定删除这个任务吗？'};$answer=[Windows.MessageBox]::Show($message,$(if($unschedule){'取消安排'}else{'删除任务'}),'YesNo','Warning');if($answer -ne [Windows.MessageBoxResult]::Yes){return $false}}
 if($unschedule){return [bool](Save-Item $item @{scheduleDate='';start=''})}
 $script:busy=$true
 try{
  $id=[string]$item.id;$writeDate=if($id -like 'weekly:*'){Get-AgendaWeekStart $item}elseif($item.agendaDate){$item.agendaDate}else{$script:lastDate}
  if($id -like 'class:*'){
   $current=@($script:data.schedule|Where-Object {$_.id -eq $id})|Select-Object -First 1
   if($current){Set-LocalValue $current 'taskTitle' '';Set-LocalValue $current 'detail' '';Set-LocalValue $current 'done' $false;Set-LocalValue $current 'incomplete' $false;Set-LocalValue $current 'highlight' $false;Set-LocalValue $current 'planned' '';Set-LocalValue $current 'actual' ''}
  }else{
   $script:data.schedule=@($script:data.schedule|Where-Object {$_.id -ne $id});$script:data.tasks=@($script:data.tasks|Where-Object {$_.id -ne $id})
   if($id -like 'weekly:*'){$script:data.weekly=@($script:data.weekly|Where-Object {$_.id -ne $id})}
  }
  if($id -like 'local:*' -and -not ($script:syncIds -and $script:syncIds.ContainsKey($id))){$script:pendingOps=@($script:pendingOps|Where-Object {$_.id -ne $id})}
  else{$fields=@{delete=$true};if($id -like 'weekly:*'){$fields.deleteScope='weekly'};Add-PendingOperation $id $writeDate $fields $item}
  Sort-LocalSchedule;Write-AgendaCache;$script:statusLabel.Text='本地已保存 · 待同步 '+$script:pendingOps.Count+' 项';if(Get-Command Render-Page -ErrorAction SilentlyContinue){Render-Page};return $true
 }catch{[void][Windows.MessageBox]::Show($_.Exception.Message,'删除失败','OK','Information');return $false}
 finally{$script:busy=$false}
}
function Move-AgendaTask($item,$target,[bool]$after){
 if($item.id -notmatch '^(daily|local|habit|weekly|event):' -or $target.id -notmatch '^(daily|local|habit|weekly|event):' -or $item.id -eq $target.id){return $false}
 $weekly=$item.id -like 'weekly:*';if($weekly -ne ($target.id -like 'weekly:*')){return $false}
 if($item.agendaDate -and $target.agendaDate -and $item.agendaDate -ne $target.agendaDate){return $false}
 $sourceItems=if($weekly){$script:data.weekly}else{@($script:data.tasks)+@($script:data.schedule|Where-Object {$_.id -match '^(daily|local|habit|event):'})}
 $items=New-Object 'System.Collections.Generic.List[object]';foreach($entry in $sourceItems){if($weekly){$items.Add($entry);continue};if($entry.id -notmatch '^(daily|local|habit|event):'){continue};$entryDate=if($entry.agendaDate){$entry.agendaDate}else{$script:lastDate};$movingDate=if($item.agendaDate){$item.agendaDate}else{$script:lastDate};if($entryDate -eq $movingDate){$items.Add($entry)}}
 $sourceIndex=-1;$targetIndex=-1
 for($index=0;$index -lt $items.Count;$index++){if($items[$index].id -eq $item.id){$sourceIndex=$index};if($items[$index].id -eq $target.id){$targetIndex=$index}}
 if($sourceIndex -lt 0 -or $targetIndex -lt 0){return $false}
 $moving=$items[$sourceIndex];$items.RemoveAt($sourceIndex);$targetIndex=-1
 for($index=0;$index -lt $items.Count;$index++){if($items[$index].id -eq $target.id){$targetIndex=$index;break}}
 $insertAt=$targetIndex+[int]$after;$items.Insert($insertAt,$moving);$ordered=@();foreach($entry in $items){$ordered+=@($entry)}
 if($weekly){$script:data.weekly=$ordered}else{
  $schedulePattern='^(daily|local|habit|event):'
  $day=if($item.agendaDate){$item.agendaDate}else{$script:lastDate}
  $script:data.tasks=@($script:data.tasks|Where-Object {$entryDate=if($_.agendaDate){$_.agendaDate}else{$script:lastDate};$_.id -notmatch $schedulePattern -or $entryDate -ne $day})+@($ordered|Where-Object {-not $_.start})
  $script:data.schedule=@($script:data.schedule|Where-Object {$entryDate=if($_.agendaDate){$_.agendaDate}else{$script:lastDate};$_.id -notmatch $schedulePattern -or $entryDate -ne $day})+@($ordered|Where-Object {$_.start}|Sort-Object @{Expression={$_.agendaDate}},@{Expression={if($_.time){($_.time -split '—')[0]}else{'99:99'}}})
 }
 $rank=0
 foreach($entry in $ordered){$rank++;Set-LocalValue $entry 'order' $rank;if($weekly){Add-PendingOperation $entry.id (Get-AgendaWeekStart $entry) @{order=$rank} $entry}else{$writeDate=if($entry.agendaDate){$entry.agendaDate}else{$script:lastDate};Add-PendingOperation $entry.id $writeDate @{order=$rank} $entry}}
 if($weekly){Sort-LocalWeekly}
 Write-AgendaCache;$script:statusLabel.Text='本地已保存 · 待同步 '+$script:pendingOps.Count+' 项';if(Get-Command Render-Page -ErrorAction SilentlyContinue){Render-Page};return $true
}
function Sort-LocalSchedule($target=$script:data){$target.schedule=@($target.schedule|Sort-Object @{Expression={$_.agendaDate}},@{Expression={if($_.time){($_.time -split '—')[0]}else{'99:99'}}})}
function Sort-LocalTasks($target=$script:data){
 $pattern='^(daily|local|habit|event):'
 $ordered=@($target.tasks|Where-Object {$_.id -match $pattern}|Sort-Object @{Expression={$_.agendaDate}},@{Expression={if($null -ne $_.order){[double]$_.order}elseif($_.id -like 'habit:*'){[int](($_.id -split ':')[3])-69}else{[double]::MaxValue}}})
 $index=0;$target.tasks=@(foreach($item in $target.tasks){if($item.id -match $pattern){if($index -lt $ordered.Count){$ordered[$index]};$index++}else{$item}})
}
function Sort-LocalWeekly($target=$script:data){
 $target.weekly=@($target.weekly|Sort-Object {[double]$_.order})
 $ordered=@($target.tasks|Where-Object {$_.id -like 'weekly:*'}|Sort-Object @{Expression={$_.agendaDate}},@{Expression={[double]$_.order}})
 $index=0;$target.tasks=@(foreach($item in $target.tasks){if($item.id -like 'weekly:*'){if($index -lt $ordered.Count){$ordered[$index]};$index++}else{$item}})
}
function Get-TimeRange($item){
 if($item.time -notmatch '^(\d{1,2}):(\d{2})(?:—(\d{1,2}):(\d{2}))?$'){return $null}
 $start=[int]$Matches[1]*60+[int]$Matches[2];$end=if($Matches[3]){[int]$Matches[3]*60+[int]$Matches[4]}else{$start+30}
 if($start -ge 1440 -or $end -le $start -or $end -gt 1440){return $null}
 return [PSCustomObject]@{start=$start;end=$end}
}
function Get-AgendaTargetDate($item){
 if($item.id -like 'weekly:*' -and $script:page -eq 3 -and $script:dragScheduleDate){return $script:dragScheduleDate}
 if($item.agendaDate){return $item.agendaDate}
 $page=if($null -ne $script:page){[int]$script:page}else{0}
 if($script:data.days -and $page -ge 0 -and $page -lt 3 -and @($script:data.days).Count -gt $page){return $script:data.days[$page].date}
 return $script:lastDate
}
function ConvertTo-AgendaMinutes([string]$value,[bool]$allowMidnight=$false){
 if($value -notmatch '^(\d{1,2}):(\d{2})$'){throw '时间请使用 HH:mm 格式。'}
 $hour=[int]$Matches[1];$minute=[int]$Matches[2]
 if($minute -gt 59 -or $hour -gt 23){if($allowMidnight -and $hour -eq 24 -and $minute -eq 0){return 1440};throw '时间须在 00:00—23:59 内，结束时间可填 24:00。'}
 return $hour*60+$minute
}
function Get-AgendaEndTime([string]$start,$duration){
 $total=[int][Math]::Floor((ConvertTo-AgendaMinutes $start)+[double]$duration)
 return ('{0:D2}:{1:D2}' -f [int][Math]::Floor($total/60),[int]($total%60))
}
function Get-AgendaWeekStart($item){
 if($item.weekStart){return $item.weekStart}
 $week=($item.id -split ':')[1]
 if($week -match '^秋(\d+)周$'){return ([datetime]'2026-10-19').AddDays(([int]$Matches[1]-6)*7).ToString('yyyy-MM-dd')}
 if($week -match '^冬(\d+)周$'){return ([datetime]'2026-11-09').AddDays(([int]$Matches[1]-1)*7).ToString('yyyy-MM-dd')}
 return $script:data.weeklyDate
}
function Get-ValidatedAgendaFields($item,$fields){
 $validated=@{};foreach($key in $fields.Keys){$validated[$key]=$fields[$key]}
 if($validated.ContainsKey('incomplete')){$validated.incomplete=[bool]$validated.incomplete;if($validated.incomplete){$validated.done=$false}}
 if($validated.ContainsKey('done')){if($validated.done -or -not $validated.ContainsKey('incomplete')){$validated.incomplete=$false}}
 $date=if($item.id -like 'weekly:*'){Get-AgendaWeekStart $item}elseif($item.agendaDate){$item.agendaDate}else{$script:lastDate}
 if($date -and ($date -lt '2026-10-18' -or $date -gt '2027-01-17')){throw '本版本只支持安排 2026-10-18 至 2027-01-17 的任务。'}
 if($validated.ContainsKey('title')){$validated.title=([string]$validated.title).Trim();if(-not $validated.title){throw '请填写任务名称；删除任务请使用删除按钮。'}}
 foreach($key in @('planned','actual')){if($validated.ContainsKey($key) -and $validated[$key] -ne '' -and $null -ne $validated[$key]){$minutes=0.0;if(-not [double]::TryParse([string]$validated[$key],[ref]$minutes) -or [double]::IsNaN($minutes) -or [double]::IsInfinity($minutes) -or $minutes -lt 0 -or $minutes -gt 1440){throw '时长须在 0—1440 分钟内。'};$validated[$key]=$minutes}}
 $start=if($validated.ContainsKey('start')){[string]$validated.start}else{[string]$item.start}
 if($start){$begin=ConvertTo-AgendaMinutes $start;$length=if($validated.ContainsKey('planned')){$validated.planned}else{$item.planned};if([double]$length -le 0 -or $begin+[double]$length -gt 1440){throw '结束时间必须晚于开始时间，且不能超过当天 24:00。'};if($validated.ContainsKey('start')){$validated.start=('{0:D2}:{1:D2}' -f [int][Math]::Floor($begin/60),[int]($begin%60))}}
 if($item.id -like 'weekly:*' -and $validated.ContainsKey('scheduleDate')){
  $date=[string]$validated.scheduleDate
  if($date){$parsed=[datetime]::MinValue;if(-not [datetime]::TryParseExact($date,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$parsed)){throw '日期请使用有效的 YYYY-MM-DD。'};$weekStart=[datetime]::ParseExact((Get-AgendaWeekStart $item),'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture);if($parsed -lt $weekStart -or $parsed -gt $weekStart.AddDays(6)){throw '周任务只能安排在所属周的周一至周日。'}}
  elseif($start){throw '给周任务安排时间时，请填写日期。'}
 }
 if($item.id -like 'event:*' -and $validated.ContainsKey('scheduleDate')){$date=[string]$validated.scheduleDate;$parsed=[datetime]::MinValue;if(-not [datetime]::TryParseExact($date,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$parsed) -or $date -lt '2026-10-18' -or $date -gt '2027-01-17'){throw '事项日期须在2026-10-18至2027-01-17内。'}}
 return $validated
}
function Test-AgendaSlotAvailable($item,[string]$start,[int]$duration=0,[string]$date=''){
 try{$begin=ConvertTo-AgendaMinutes $start}catch{return $false}
 $length=if($duration -gt 0){$duration}elseif([double]$item.planned -gt 0){[int][double]$item.planned}else{30};$finish=$begin+$length
 if($begin -lt 420 -or $finish -gt 1440){return $false}
 if(-not $date){$date=Get-AgendaTargetDate $item}
 foreach($existing in $script:data.schedule){$existingDate=if($existing.agendaDate){$existing.agendaDate}else{$script:lastDate};if($existing.id -eq $item.id -or $existingDate -ne $date){continue};$range=Get-TimeRange $existing;if($range -and $begin -lt $range.end -and $finish -gt $range.start){return $false}}
 return $true
}
function Set-AgendaTaskTime($item,[string]$start){
 $duration=if([double]$item.planned -gt 0){[int][double]$item.planned}else{30}
 if(-not (Test-AgendaSlotAvailable $item $start $duration)){return $false}
 $fields=@{start=$start;planned=$duration}
 if($item.id -like 'weekly:*'){$fields.scheduleDate=Get-AgendaTargetDate $item}
 return [bool](Save-Item $item $fields)
}
function Add-AgendaDayTargets {
 $panel=$script:window.FindName('DaySlotPanel');if(-not $panel){return};$panel.Children.Clear();$panel.Visibility='Visible'
 try{$monday=[datetime]::ParseExact($script:data.weeklyDate,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture)}catch{return}
 $selected=$script:dragScheduleDate;if(-not $selected -or $selected -lt $monday.ToString('yyyy-MM-dd') -or $selected -gt $monday.AddDays(6).ToString('yyyy-MM-dd')){$selected=$script:data.weeklyDate};$script:dragScheduleDate=$selected
 for($offset=0;$offset -lt 7;$offset++){
  $day=$monday.AddDays($offset);$date=$day.ToString('yyyy-MM-dd');$target=New-Object Windows.Controls.Button;$target.Content=('周'+'一二三四五六日'[$offset]+' '+$day.ToString('MM/dd'));$target.Tag=$date;$target.Padding='5,3';$target.Margin='2';$target.FontSize=10;$target.AllowDrop=$true;$target.ToolTip='拖到这里，把周任务安排到当天'
  $target.Add_Click({param($sender,$e)$script:dragScheduleDate=$sender.Tag;foreach($button in $script:window.FindName('DaySlotPanel').Children){$button.Background='#F1F7F3'};$sender.Background='#D6EBDD'})
  $target.Add_DragEnter({param($sender,$e)if($e.Data.GetDataPresent('AgendaTask') -and $e.Data.GetData('AgendaTask').id -like 'weekly:*'){$e.Effects=[Windows.DragDropEffects]::Move;$sender.Background='#D6EBDD'}else{$e.Effects=[Windows.DragDropEffects]::None};$e.Handled=$true})
  $target.Add_DragOver({param($sender,$e)if($e.Data.GetDataPresent('AgendaTask') -and $e.Data.GetData('AgendaTask').id -like 'weekly:*'){$e.Effects=[Windows.DragDropEffects]::Move}else{$e.Effects=[Windows.DragDropEffects]::None};$e.Handled=$true})
  $target.Add_Drop({param($sender,$e)$task=$e.Data.GetData('AgendaTask');if($task -and (Save-Item $task @{scheduleDate=$sender.Tag})){$script:dragScheduleDate=$sender.Tag;$e.Effects=[Windows.DragDropEffects]::Move};$e.Handled=$true})
  if($date -eq $selected){$target.Background='#D6EBDD'}
  [void]$panel.Children.Add($target)
 }
}
function Add-AgendaTimeTargets {
 $panel=$script:window.FindName('TimeSlotPanel');$panel.Children.Clear()
 $dayPanel=$script:window.FindName('DaySlotPanel');$hint=$script:window.FindName('ScheduleHint')
 if($script:page -eq 3){Add-AgendaDayTargets;if($hint){$hint.Text='点日期后拖到时间点；也可直接拖到周几'}}
 else{if($dayPanel){$dayPanel.Children.Clear();$dayPanel.Visibility='Collapsed'};$script:dragScheduleDate=Get-AgendaTargetDate ([PSCustomObject]@{id='new'});if($hint){$hint.Text='拖到时间点，为当前日期安排任务'}}
 for($minute=7*60;$minute -le 22*60;$minute+=30){
  $start='{0:D2}:{1:D2}' -f [int][Math]::Floor($minute/60),[int]($minute%60)
  $target=New-Object Windows.Controls.Button;$target.Content=$start;$target.Tag=$start;$target.Width=66;$target.Padding='3,2';$target.Margin='2';$target.FontSize=10;$target.Foreground='#547562';$target.Background='#F1F7F3';$target.BorderBrush='#DDEAE1';$target.AllowDrop=$true
  $target.Add_DragEnter({param($sender,$e)if($e.Data.GetDataPresent('AgendaTask') -and (Test-AgendaSlotAvailable $e.Data.GetData('AgendaTask') $sender.Tag)){$e.Effects=[Windows.DragDropEffects]::Move;$sender.Background='#D6EBDD'}else{$e.Effects=[Windows.DragDropEffects]::None};$e.Handled=$true})
  $target.Add_DragOver({param($sender,$e)if($e.Data.GetDataPresent('AgendaTask') -and (Test-AgendaSlotAvailable $e.Data.GetData('AgendaTask') $sender.Tag)){$e.Effects=[Windows.DragDropEffects]::Move}else{$e.Effects=[Windows.DragDropEffects]::None};$e.Handled=$true})
  $target.Add_DragLeave({param($sender,$e)$sender.Background='#F1F7F3'})
  $target.Add_Drop({param($sender,$e)$task=$e.Data.GetData('AgendaTask');if($task){if(Set-AgendaTaskTime $task $sender.Tag){$e.Effects=[Windows.DragDropEffects]::Move}else{$script:statusLabel.Text='该时段与已有安排冲突'}};$e.Handled=$true})
  [void]$panel.Children.Add($target)
 }
}
function Get-AgendaTaskDetail($item){
 $detail=[string]$item.detail
 if($item.id -match '^(daily|local|habit|weekly|event):'){
  if($item.start -and $item.time){return [string]$item.time}
  if($detail.Trim() -match '^\d{1,2}:\d{2}(?:—\d{1,2}:\d{2})?$'){
   if($item.id -like 'weekly:*'){if($item.plannedDate){return '安排：'+([datetime]::ParseExact($item.plannedDate,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture).ToString('MM月dd日'))};return '本周未安排'}
   return ''
  }
 }
 return $detail
}
function Save-Item($item,$fields){
 if($script:busy){return $false};$script:busy=$true
 try{
  $fields=Get-ValidatedAgendaFields $item $fields
  if($item.id -ne 'new'){$current=@(@($script:data.schedule)+@($script:data.tasks)+@($script:data.weekly)|Where-Object {$_.id -eq $item.id})|Select-Object -First 1;if($current){$item=$current}}
  if($item.id -eq 'new'){
   $item=[PSCustomObject]@{id=('local:'+([guid]::NewGuid().ToString('N')));title='';detail='';done=$false;highlight=$false;start='';planned='';actual='';subject='';agendaDate=$(if($item.agendaDate){$item.agendaDate}else{$script:lastDate})}
   $maxOrder=(@($script:data.tasks|Where-Object {$_.agendaDate -eq $item.agendaDate}|ForEach-Object {[double]$_.order})|Measure-Object -Maximum).Maximum;Set-LocalValue $item 'order' ([Math]::Max(0,[double]$maxOrder)+1);$fields.order=$item.order
   $script:data.tasks=@($script:data.tasks)+@($item)
  }
  $writeDate=if($item.id -like 'weekly:*'){Get-AgendaWeekStart $item}elseif($item.agendaDate){$item.agendaDate}else{$script:lastDate}
  $f=@{};foreach($k in $fields.Keys){$f[$k]=$fields[$k];if($item.id -like 'class:*' -and $k -eq 'title'){Set-LocalValue $item 'taskTitle' $fields[$k];if($fields[$k]){Set-LocalValue $item 'detail' '课内安排'}}elseif($k -eq 'scheduleDate'){Set-LocalValue $item 'plannedDate' $fields[$k];Set-LocalValue $item 'detail' $(if($fields[$k]){'安排：'+([datetime]::ParseExact($fields[$k],'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture).ToString('MM月dd日'))}else{'本周未安排'})}else{Set-LocalValue $item $k $fields[$k]}}
  if($fields.ContainsKey('start')){
   if($fields.start){Set-LocalValue $item 'time' ($fields.start+'—'+(Get-AgendaEndTime $fields.start $fields.planned));Set-LocalValue $item 'kind' 'task'}else{Set-LocalValue $item 'time' '';Set-LocalValue $item 'kind' ''}
   Set-LocalValue $item 'detail' (Get-AgendaTaskDetail $item)
   $script:data.schedule=@($script:data.schedule|Where-Object {$_.id -ne $item.id});$script:data.tasks=@($script:data.tasks|Where-Object {$_.id -ne $item.id})
   if($fields.start){$script:data.schedule+=@($item)}elseif($item.id -match '^(daily|local|habit|event):'){$script:data.tasks+=@($item)}
  }
  if($item.id -like 'weekly:*'){
   $script:data.weekly=@($script:data.weekly|ForEach-Object {if($_.id -eq $item.id){$item}else{$_}})
   $script:data.schedule=@($script:data.schedule|Where-Object {$_.id -ne $item.id});$script:data.tasks=@($script:data.tasks|Where-Object {$_.id -ne $item.id})
   $targetDay=@($script:data.days|Where-Object {$_.date -eq $item.plannedDate})
   if($targetDay.Count){Set-LocalValue $item 'agendaDate' $item.plannedDate;if($item.start){Set-LocalValue $item 'kind' 'task';Set-LocalValue $item 'time' ($item.start+'—'+(Get-AgendaEndTime $item.start $item.planned));$script:data.schedule+=@($item)}else{$script:data.tasks+=@($item)}}
  }
  if($item.id -like 'event:*' -and $fields.ContainsKey('scheduleDate')){
   Set-LocalValue $item 'agendaDate' $fields.scheduleDate
   $script:data.schedule=@($script:data.schedule|Where-Object {$_.id -ne $item.id});$script:data.tasks=@($script:data.tasks|Where-Object {$_.id -ne $item.id})
   if(@($script:data.days|Where-Object {$_.date -eq $item.agendaDate}).Count){if($item.start){$script:data.schedule+=@($item)}else{$script:data.tasks+=@($item)}}
  }
  Add-PendingOperation $item.id $writeDate $f $item;Sort-LocalSchedule;Sort-LocalTasks;Write-AgendaCache
  $script:statusLabel.Text='本地已保存 · 待同步 '+$script:pendingOps.Count+' 项'
  $stateOnly=@($fields.Keys|Where-Object {$_ -notin @('done','incomplete','highlight')}).Count -eq 0
  if(-not ($stateOnly -and (Update-AgendaTaskVisual $item)) -and (Get-Command Render-Page -ErrorAction SilentlyContinue)){Render-Page}
  return $true
 }catch{[void][Windows.MessageBox]::Show($_.Exception.Message,'本地保存失败','OK','Information');return $false}
 finally{$script:busy=$false}
}
function New-TimePicker([string]$value){
 $box=New-Object Windows.Controls.ComboBox;$box.IsEditable=$true;$box.Width=120;$box.Margin='0,6,8,12'
 for($m=7*60;$m -le 23*60+30;$m+=15){[void]$box.Items.Add(('{0:D2}:{1:D2}' -f [int][Math]::Floor($m/60),[int]($m%60)))}
 $box.Text=$value;return $box
}
function Edit-Item($item){
 Complete-AgendaCheckboxClick
 $dialog=New-Object Windows.Window;$dialog.Title=if($item.id -eq 'new'){'新增今日任务'}elseif($item.id -like 'class:*'){'编辑课内安排'}else{'编辑任务'}
 $dialog.Width=355;$dialog.SizeToContent='Height';$dialog.ResizeMode='NoResize';$dialog.WindowStartupLocation='CenterOwner';$dialog.Owner=$script:window;$dialog.FontFamily='Microsoft YaHei';$dialog.Background='#FCFEFD'
 $panel=New-Object Windows.Controls.StackPanel;$panel.Margin='18';$dialog.Content=$panel
 $label=New-Object Windows.Controls.TextBlock;$label.Text='任务名称／完成标准';[void]$panel.Children.Add($label)
 $title=New-Object Windows.Controls.TextBox;$title.Margin='0,7,0,12';$title.TextWrapping='Wrap';$title.MinHeight=55;$title.AcceptsReturn=$true
 $title.Text=if($item.id -like 'class:*'){$item.taskTitle}else{$item.title};[void]$panel.Children.Add($title)
 $timed=$item.id -eq 'new' -or $item.id -match '^(daily|local|habit|weekly|event):' -or $item.id -like 'class:*';$start=$null;$end=$null;$scheduleDate=$null
 if($timed){
  if($item.id -match '^(weekly|event):'){
   $label=New-Object Windows.Controls.TextBlock;$label.Text=if($item.id -like 'event:*'){'安排日期'}else{'安排日期（可留空）'};[void]$panel.Children.Add($label)
   $scheduleDate=New-Object Windows.Controls.TextBox;$scheduleDate.Margin='0,6,0,10';$scheduleDate.Text=$item.plannedDate;[void]$panel.Children.Add($scheduleDate)
  }
  if($item.id -notlike 'class:*'){
   $label=New-Object Windows.Controls.TextBlock;$label.Text='安排时间段（可手动输入 HH:mm；周任务请先填日期）';[void]$panel.Children.Add($label)
   $timeRow=New-Object Windows.Controls.StackPanel;$timeRow.Orientation='Horizontal'
   $start=New-TimePicker $item.start
   $endValue='';if($item.start -and $item.planned){try{$endValue=Get-AgendaEndTime $item.start $item.planned}catch{}}
   $end=New-TimePicker $endValue
   [void]$timeRow.Children.Add($start);[void]$timeRow.Children.Add($end);[void]$panel.Children.Add($timeRow)
  }
 }
 $done=New-Object Windows.Controls.CheckBox;$done.Content='已完成';$done.IsChecked=[bool]$item.done;$done.Margin='0,5,0,9';[void]$panel.Children.Add($done)
 $incomplete=New-Object Windows.Controls.CheckBox;$incomplete.Content='未完成（明确标记）';$incomplete.IsChecked=[bool]$item.incomplete;$incomplete.Margin='0,0,0,9';$incomplete.Foreground='#658477';[void]$panel.Children.Add($incomplete)
 $done.Tag=$incomplete;$incomplete.Tag=$done
 $done.Add_Checked({param($sender,$e)$sender.Tag.IsChecked=$false})
 $incomplete.Add_Checked({param($sender,$e)$sender.Tag.IsChecked=$false})
 $highlight=New-Object Windows.Controls.CheckBox;$highlight.Content='强调显示';$highlight.IsChecked=[bool]$item.highlight;$highlight.Margin='0,3,0,14';[void]$panel.Children.Add($highlight)
 $save=New-Object Windows.Controls.Button;$save.Content='保存任务';$save.Padding='10,7';[void]$panel.Children.Add($save)
 $save.Tag=@{item=$item;title=$title;start=$start;end=$end;scheduleDate=$scheduleDate;done=$done;incomplete=$incomplete;highlight=$highlight;dialog=$dialog;timed=$timed}
 $save.Add_Click({param($sender,$e)
  $v=$sender.Tag;$f=@{title=$v.title.Text;done=[bool]$v.done.IsChecked;incomplete=[bool]$v.incomplete.IsChecked;highlight=[bool]$v.highlight.IsChecked}
  if($v.timed){
   if($v.start){
    $st=$v.start.Text.Trim();$et=$v.end.Text.Trim()
    if(($st -and -not $et) -or ($et -and -not $st)){[void][Windows.MessageBox]::Show('开始和结束时间需要同时填写。','时间段未完整','OK','Information');return}
    if($st){
     try{$begin=ConvertTo-AgendaMinutes $st;$finish=ConvertTo-AgendaMinutes $et $true}catch{[void][Windows.MessageBox]::Show($_.Exception.Message,'时间格式错误','OK','Information');return}
     $mins=$finish-$begin;if($mins -le 0){[void][Windows.MessageBox]::Show('结束时间必须晚于开始时间。','时间段不正确','OK','Information');return}
     $f.start=('{0:D2}:{1:D2}' -f [int][Math]::Floor($begin/60),[int]($begin%60));$f.planned=[int]$mins
    }else{$f.start='';$f.planned=''}
    if($v.scheduleDate){$dateText=$v.scheduleDate.Text.Trim();if($dateText -and $dateText -notmatch '^202[6-7]-\d{2}-\d{2}$'){[void][Windows.MessageBox]::Show('日期请使用 YYYY-MM-DD。','日期格式错误','OK','Information');return};if($st -and -not $dateText){[void][Windows.MessageBox]::Show('给周任务安排时间时，请填写日期。','缺少日期','OK','Information');return};$f.scheduleDate=$dateText}
   }
  }
  if(Save-Item $v.item $f){$v.dialog.Close()}
 })
 [void]$dialog.ShowDialog()
}
function Complete-AgendaCheckboxClick {
 $pending=$script:checkboxClick
 if(-not $pending){return}
 $script:checkboxClick=$null;$pending.timer.Stop()
 [void](Save-Item $pending.item @{done=$pending.done})
}
function Invoke-AgendaCheckboxClick($check,[bool]$doubleClick=$false){
 if($doubleClick){
  if($script:checkboxClick -and $script:checkboxClick.item.id -ne $check.Tag.id){Complete-AgendaCheckboxClick}
  if($script:checkboxClick){$script:checkboxClick.timer.Stop();$script:checkboxClick=$null}
  [void](Save-Item $check.Tag @{incomplete=$true})
  return
 }
 Complete-AgendaCheckboxClick
 if($check.Tag.incomplete){$check.IsChecked=$true}
 if(-not $script:doubleClickMs){Add-Type -AssemblyName System.Windows.Forms;$script:doubleClickMs=[Windows.Forms.SystemInformation]::DoubleClickTime}
 $timer=New-Object Windows.Threading.DispatcherTimer
 $timer.Interval=[TimeSpan]::FromMilliseconds($script:doubleClickMs)
 $timer.Add_Tick({param($sender,$e)if($script:checkboxClick -and [object]::ReferenceEquals($script:checkboxClick.timer,$sender)){Complete-AgendaCheckboxClick}else{$sender.Stop()}})
 $script:checkboxClick=@{item=$check.Tag;done=[bool]$check.IsChecked;timer=$timer}
 $timer.Start()
}
function Get-AgendaCheckboxTemplate {
 if(-not $script:agendaCheckboxTemplate){
  $script:agendaCheckboxTemplate=[Windows.Markup.XamlReader]::Parse(@'
<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" TargetType="CheckBox">
 <Grid Background="Transparent"><Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
  <Border Name="Box" Width="14" Height="14" Margin="0,3,7,0" VerticalAlignment="Top" BorderThickness="1" BorderBrush="#91A298" Background="White">
   <Grid><Path Name="Check" Data="M 2,6 L 5,9 L 10,2" Stroke="#547D65" StrokeThickness="1.6" Visibility="Collapsed"/><Path Name="Slash" Data="M 2,10 L 10,2" Stroke="#969C98" StrokeThickness="1.5" Visibility="Collapsed"/></Grid>
  </Border>
  <ContentPresenter Grid.Column="1" VerticalAlignment="Center"/>
 </Grid>
 <ControlTemplate.Triggers>
  <Trigger Property="IsChecked" Value="True"><Setter TargetName="Check" Property="Visibility" Value="Visible"/><Setter TargetName="Box" Property="Background" Value="#EDF5EF"/></Trigger>
  <Trigger Property="IsChecked" Value="{x:Null}" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><Setter TargetName="Slash" Property="Visibility" Value="Visible"/><Setter TargetName="Box" Property="Background" Value="#ECEEEC"/><Setter TargetName="Box" Property="BorderBrush" Value="#A7ADA9"/></Trigger>
  <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Box" Property="BorderBrush" Value="#547D65"/></Trigger>
 </ControlTemplate.Triggers>
</ControlTemplate>
'@)
 }
 return $script:agendaCheckboxTemplate
}
function Update-AgendaTaskVisual($item){
 if(-not $script:content){return $false}
 $found=$false
 foreach($border in $script:content.Children){
  if($border -isnot [Windows.Controls.Border] -or $border.Tag.id -ne $item.id){continue}
  $found=$true;$card=$border.Tag
  foreach($field in @('done','incomplete','highlight')){Set-LocalValue $card $field ([bool]$item.$field)}
  $border.Background=if($card.incomplete){'#F2F3F2'}elseif($card.done){'#F0F8F3'}elseif($card.highlight){'#FFF7DF'}else{'#FFFFFF'}
  $row=$border.Child;$check=@($row.Children|Where-Object {[Windows.Controls.Grid]::GetColumn($_) -eq 1})[0].Children[0]
  $check.IsChecked=if($card.incomplete){$null}else{[bool]$card.done}
  $check.Content.Foreground=if($card.incomplete){'#969C98'}elseif($card.done){'#8AA197'}else{'#2D463C'}
  if($card.done){$check.Content.TextDecorations=[Windows.TextDecorations]::Strikethrough}else{$check.Content.TextDecorations=$null}
  $check.Content.FontWeight=if($card.highlight){[Windows.FontWeights]::Bold}else{[Windows.FontWeights]::Normal}
  $star=@($row.Children|Where-Object {[Windows.Controls.Grid]::GetColumn($_) -eq 2})[0]
  $star.IsChecked=[bool]$card.highlight;$star.Content=if($card.highlight){'★'}else{'☆'};$star.Foreground=if($card.highlight){'#B68B31'}else{'#A7BBB0'}
 }
 return $found
}
function Add-Task($item){
 $border=New-Object Windows.Controls.Border;$border.Tag=$item;$border.Padding='9,6';$border.Margin='0,0,0,3';$border.CornerRadius='8';$border.BorderBrush='#EBF1EE';$border.BorderThickness='0,0,0,1'
 if($item.id -match '^(daily|local|weekly|habit|event):'){
  $border.AllowDrop=$true
  $border.Add_DragEnter({param($sender,$e)$dragged=$e.Data.GetData('AgendaTask');if($dragged -and (($dragged.id -like 'weekly:*' -and $sender.Tag.id -like 'weekly:*') -or ($dragged.id -match '^(daily|local|habit|event):' -and $sender.Tag.id -match '^(daily|local|habit|event):'))){$e.Effects=[Windows.DragDropEffects]::Move;$sender.Background='#E6F2EB'}else{$e.Effects=[Windows.DragDropEffects]::None};$e.Handled=$true})
  $border.Add_DragOver({param($sender,$e)$dragged=$e.Data.GetData('AgendaTask');if($dragged -and (($dragged.id -like 'weekly:*' -and $sender.Tag.id -like 'weekly:*') -or ($dragged.id -match '^(daily|local|habit|event):' -and $sender.Tag.id -match '^(daily|local|habit|event):'))){$e.Effects=[Windows.DragDropEffects]::Move}else{$e.Effects=[Windows.DragDropEffects]::None};$e.Handled=$true})
  $border.Add_DragLeave({param($sender,$e)$item=$sender.Tag;$sender.Background=if($item.incomplete){'#F2F3F2'}elseif($item.done){'#F0F8F3'}elseif($item.highlight){'#FFF7DF'}else{'#FFFFFF'}})
  $border.Add_Drop({param($sender,$e)$dragged=$e.Data.GetData('AgendaTask');$after=$e.GetPosition($sender).Y -ge ($sender.ActualHeight/2);if($dragged -and (Move-AgendaTask $dragged $sender.Tag $after)){$e.Effects=[Windows.DragDropEffects]::Move};$e.Handled=$true})
 }
 $border.Background=if($item.incomplete){'#F2F3F2'}elseif($item.done){'#F0F8F3'}elseif($item.highlight){'#FFF7DF'}else{'#FFFFFF'}
 $row=New-Object Windows.Controls.Grid
 foreach($width in @('Auto','*','Auto','Auto','Auto')){$cd=New-Object Windows.Controls.ColumnDefinition;$cd.Width=$width;$row.ColumnDefinitions.Add($cd)}
 $stack=New-Object Windows.Controls.StackPanel;$check=New-Object Windows.Controls.CheckBox;$check.Template=Get-AgendaCheckboxTemplate;$check.IsChecked=if($item.incomplete){$null}else{[bool]$item.done};$check.Tag=$item
 $txt=New-Object Windows.Controls.TextBlock;$txt.Text=$item.title;$txt.TextWrapping='Wrap';$txt.MaxWidth=235;$txt.Foreground=if($item.incomplete){'#969C98'}elseif($item.done){'#8AA197'}else{'#2D463C'};$txt.Tag=$item
 if($item.id -match '^(daily|local|weekly|habit|event):'){
  $handle=New-Object Windows.Controls.Border;$handle.Background='Transparent';$handle.Width=26;$handle.MinHeight=34;$handle.VerticalAlignment='Top';$handle.Margin='0,0,4,0';$handle.Tag=$item;$handle.Cursor=[Windows.Input.Cursors]::SizeAll;$handle.ToolTip='按住拖动：调顺序或安排到某天/时间'
  $glyph=New-Object Windows.Controls.TextBlock;$glyph.Text='⠿';$glyph.FontSize=18;$glyph.Foreground='#9AAFA3';$glyph.HorizontalAlignment='Center';$glyph.VerticalAlignment='Top';$handle.Child=$glyph
  $handle.Add_PreviewMouseLeftButtonDown({param($sender,$e)$script:dragOrigin=$e.GetPosition($sender)})
  $handle.Add_PreviewMouseLeftButtonUp({$script:dragOrigin=$null})
  $handle.Add_PreviewMouseMove({param($sender,$e)
   if($e.LeftButton -ne [Windows.Input.MouseButtonState]::Pressed -or -not $script:dragOrigin){return}
   $position=$e.GetPosition($sender);if([Math]::Abs($position.X-$script:dragOrigin.X) -lt [Windows.SystemParameters]::MinimumHorizontalDragDistance -and [Math]::Abs($position.Y-$script:dragOrigin.Y) -lt [Windows.SystemParameters]::MinimumVerticalDragDistance){return}
   $script:dragOrigin=$null;$payload=New-Object Windows.DataObject;$payload.SetData('AgendaTask',$sender.Tag);[void][Windows.DragDrop]::DoDragDrop($sender,$payload,[Windows.DragDropEffects]::Move)
  })
  [Windows.Controls.Grid]::SetColumn($handle,0);[void]$row.Children.Add($handle)
 }
 if($item.highlight){$txt.FontWeight='Bold'};if($item.done){$txt.TextDecorations=[Windows.TextDecorations]::Strikethrough};$check.Content=$txt
 $check.Add_Click({param($sender,$e)Invoke-AgendaCheckboxClick $sender});[void]$stack.Children.Add($check)
 $check.Add_PreviewMouseLeftButtonDown({param($sender,$e)if($e.ClickCount -eq 2){Invoke-AgendaCheckboxClick $sender $true;$e.Handled=$true}})
 $check.ToolTip='单击完成／取消完成；双击标记未完成（灰色斜线）'
 $check.ContextMenu=New-Object Windows.Controls.ContextMenu
 foreach($choice in @(@{label='标记未完成';fields=@{incomplete=$true}},@{label='标记已完成';fields=@{done=$true}},@{label='清除状态';fields=@{done=$false}})){
  $option=New-Object Windows.Controls.MenuItem;$option.Header=$choice.label;$option.Foreground='#658477';$option.Tag=@{item=$item;fields=$choice.fields}
  $option.Add_Click({param($sender,$e)Complete-AgendaCheckboxClick;[void](Save-Item $sender.Tag.item $sender.Tag.fields)})
  [void]$check.ContextMenu.Items.Add($option)
 }
 $detailText=Get-AgendaTaskDetail $item
 if($detailText){$detail=New-Object Windows.Controls.TextBlock;$detail.Text=$detailText;$detail.FontSize=11;$detail.Foreground='#7B8795';$detail.Margin='17,4,0,0';[void]$stack.Children.Add($detail)}
 [Windows.Controls.Grid]::SetColumn($stack,1);[void]$row.Children.Add($stack)
 $star=New-Object Windows.Controls.Primitives.ToggleButton;$star.Content=if($item.highlight){'★'}else{'☆'};$star.ToolTip='强调';$star.FontSize=16;$star.Foreground=if($item.highlight){'#B68B31'}else{'#A7BBB0'};$star.Padding='3,0';$star.Margin='3,0';$star.Background='Transparent';$star.BorderThickness=0;$star.IsChecked=[bool]$item.highlight;$star.Tag=$item;$star.VerticalAlignment='Top';[Windows.Controls.Grid]::SetColumn($star,2)
 $star.Add_Click({param($sender,$e)[void](Save-Item $sender.Tag @{highlight=[bool]$sender.IsChecked})});[void]$row.Children.Add($star)
 $edit=New-Object Windows.Controls.Button;$edit.Content='⋯';$edit.ToolTip='编辑';$edit.Foreground='#94AAA0';$edit.FontSize=16;$edit.Padding='4,0';$edit.BorderThickness=0;$edit.Background='Transparent';$edit.Tag=$item;[Windows.Controls.Grid]::SetColumn($edit,3);$edit.VerticalAlignment='Top';$edit.Add_Click({param($sender,$e)Edit-Item $sender.Tag});[void]$row.Children.Add($edit)
 if($item.id -ne 'new'){$delete=New-Object Windows.Controls.Button;$unschedule=$item.id -like 'weekly:*' -and $script:page -ne 3;$delete.Content=if($unschedule){'撤'}else{'删'};$delete.ToolTip=if($unschedule){'取消安排，保留本周任务'}else{'删除任务'};$delete.Foreground='#658477';$delete.FontSize=11;$delete.Padding='4,3';$delete.Margin='1,2,0,0';$delete.BorderThickness=0;$delete.Background='Transparent';$delete.Tag=$item;[Windows.Controls.Grid]::SetColumn($delete,4);$delete.VerticalAlignment='Top';$delete.Add_Click({param($sender,$e)[void](Remove-AgendaItem $sender.Tag)});[void]$row.Children.Add($delete)}
 $border.Child=$row;[void]$script:content.Children.Add($border)
}
