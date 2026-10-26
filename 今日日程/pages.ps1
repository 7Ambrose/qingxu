function Render-Page {
 if(-not $script:data){return}
 $data=$script:data;$script:content.Children.Clear()
 $script:window.FindName('TodayTab').IsChecked=($script:page -eq 0)
 $script:window.FindName('TomorrowTab').IsChecked=($script:page -eq 1)
 $script:window.FindName('AfterTomorrowTab').IsChecked=($script:page -eq 2)
 $script:window.FindName('WeekTab').IsChecked=($script:page -eq 3)
 $script:window.FindName('NewButton').Visibility=if($script:page -lt 3){'Visible'}else{'Collapsed'}
 $days=@($data.days);if(-not $days.Count){$days=@([PSCustomObject]@{date=$data.date;title=$data.title})}
 $today=[DateTime]::UtcNow.AddHours(8).ToString('yyyy-MM-dd')
 for($index=0;$index -lt 3;$index++){$button=$script:window.FindName(@('TodayTab','TomorrowTab','AfterTomorrowTab')[$index]);if($index -lt $days.Count){$date=[datetime]::ParseExact($days[$index].date,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture);$button.Content=if($days[$index].date -eq $today){'今天'}else{$date.ToString('MM/dd')}}}
 $script:window.FindName('DatePickerButton').Content='选日期⌄'
 $script:dateLabel.Text=if($script:page -lt 3){if($script:page -lt $days.Count){$days[$script:page].title}else{$data.title}}else{$data.weekTitle}
 if($script:page -eq 3){
  if(-not $data.PSObject.Properties['weekSchedule']){
   Set-LocalValue $data 'weekSchedule' @()  # 先占位，避免读取让出消息循环时被重入
   try{$part=Invoke-Backend ('--date '+$data.weeklyDate) -Responsive;if($part.error){throw $part.error};Set-LocalValue $data 'weekSchedule' @($part.weekSchedule)}catch{Set-LocalValue $data 'weekSchedule' @()}
  }
  if($data.weekSchedule.Count){Add-Section '本周课程';foreach($course in $data.weekSchedule){$date=[datetime]::ParseExact($course.date,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture);$row=New-Object Windows.Controls.StackPanel;$row.Tag=$course;$row.Margin='7,1,0,5';$line=New-Object Windows.Controls.TextBlock;$line.Text=$course.weekday+' '+$date.ToString('MM/dd')+' · '+$course.time;$line.FontSize=10;$line.Foreground='#8EA79A';[void]$row.Children.Add($line);$name=New-Object Windows.Controls.TextBlock;$name.Text=$course.title;$name.TextWrapping='Wrap';$name.Foreground='#314E3E';$name.Margin='0,1,0,0';[void]$row.Children.Add($name);[void]$script:content.Children.Add($row)}}
  if($data.weekly.Count){Add-Section '每周任务 · 可拖动排序或拖到日期安排';foreach($item in $data.weekly){Add-Task $item}}
 }else{
  if($script:page -lt $days.Count){$day=$days[$script:page]
   $daySchedule=@($data.schedule|Where-Object {$_.agendaDate -eq $day.date})
   $dayTasks=@($data.tasks|Where-Object {$_.agendaDate -eq $day.date})
   Add-Section $day.title
   if($daySchedule.Count){
    Add-Section '日程'
    foreach($item in $daySchedule){
     if($item.kind -eq 'task'){Add-Task $item;continue}
     $row=New-Object Windows.Controls.DockPanel;$row.Margin='0,3,0,3'
     if($item.id -like 'class:*'){$edit=New-Object Windows.Controls.Button;$edit.Content='⋯';$edit.ToolTip='课内安排';$edit.FontSize=16;$edit.Padding='4,1';$edit.Tag=$item;[Windows.Controls.DockPanel]::SetDock($edit,'Right');$edit.Add_Click({param($sender,$e)Edit-Item $sender.Tag});[void]$row.Children.Add($edit)}
     $stack=New-Object Windows.Controls.StackPanel;$stack.Margin='8,2,0,4'
     $time=New-Object Windows.Controls.TextBlock;$time.Text=$item.time;$time.FontSize=10;$time.Foreground='#8EA79A';[void]$stack.Children.Add($time)
     $title=New-Object Windows.Controls.TextBlock;$title.Text=$item.title;$title.TextWrapping='Wrap';$title.Margin='0,2,0,0';$title.Foreground='#314E3E';[void]$stack.Children.Add($title);[void]$row.Children.Add($stack);[void]$script:content.Children.Add($row)
     if($item.detail -and $item.id -like 'class:*'){Add-Task ([PSCustomObject]@{id=$item.id;title=$item.taskTitle;taskTitle=$item.taskTitle;subject=$item.subject;planned=$item.planned;actual=$item.actual;done=$item.done;incomplete=$item.incomplete;highlight=$item.highlight;detail='课内安排';agendaDate=$day.date})}elseif($item.detail){Add-Text $item.detail 11 '#8DA295'}
    }
   }
   if($dayTasks.Count){Add-Section '待办';foreach($item in $dayTasks){Add-Task $item}}
   if(-not $daySchedule.Count -and -not $dayTasks.Count){Add-Text '当天暂无安排' 12 '#8DA295'}
  }
 }
 $pageDate=if($script:page -ge 0 -and $script:page -lt $days.Count){$days[$script:page].date}else{''}
 $canSchedule=if($script:page -eq 3){@($data.weekly|Where-Object {$_.id -like 'weekly:*'}).Count -gt 0}elseif($pageDate){@($data.tasks|Where-Object {$_.agendaDate -eq $pageDate -and $_.id -match '^(daily|local|habit|event):'}).Count -or @($data.schedule|Where-Object {$_.agendaDate -eq $pageDate -and $_.id -match '^(daily|local|habit|weekly|event):'}).Count}else{$false}
 $scheduleButton=$script:window.FindName('ScheduleButton');$scheduleButton.Visibility=if($canSchedule){'Visible'}else{'Collapsed'}
 if($canSchedule){Add-AgendaTimeTargets}else{$script:window.FindName('TimeSlotPopup').IsOpen=$false;$script:window.FindName('DaySlotPanel').Children.Clear();$script:window.FindName('DaySlotPanel').Visibility='Collapsed'}
 if(-not $script:syncProcess){
  $updated=if($data.updated){[string]$data.updated}else{''}
  $script:statusLabel.Text=if($script:pendingOps.Count){'本地已保存 · 待同步 '+$script:pendingOps.Count+' 项'}elseif($updated.Length -ge 5){'已同步 '+$updated.Substring(0,5)}else{'已与 Excel 同步'}
 }
 $script:content.Measure((New-Object Windows.Size 350,([double]::PositiveInfinity)))
 $script:window.Height=[Math]::Min(720,[Math]::Min(([Windows.SystemParameters]::WorkArea.Height-80),[Math]::Max(300,$script:content.DesiredSize.Height+215)))
}
function Set-Page([int]$value){$script:page=[Math]::Max(0,[Math]::Min(3,$value));Render-Page;$script:window.FindName('BodyScroll').ScrollToTop()}
function Normalize-CachedHabits($target){
 if(-not $target.PSObject.Properties['schedule'] -or -not $target.PSObject.Properties['tasks']){return}
 $byId=[ordered]@{}
 foreach($item in @($target.schedule|Where-Object {$_.id -like 'habit:*'})+@($target.tasks|Where-Object {$_.id -like 'habit:*'})){$byId[$item.id]=$item}
 if(-not $byId.Count){return}
 $target.schedule=@($target.schedule|Where-Object {$_.id -notlike 'habit:*'});$target.tasks=@($target.tasks|Where-Object {$_.id -notlike 'habit:*'})
 foreach($item in $byId.Values){
  if(-not $item.PSObject.Properties['planned'] -or [double]$item.planned -le 0){Set-LocalValue $item 'planned' 30}
  if(-not $item.start -and $item.time -match '^(\d{1,2}:\d{2})—(\d{1,2}:\d{2})$'){
   $start=[timespan]::Parse($Matches[1]);$finish=[timespan]::Parse($Matches[2]);Set-LocalValue $item 'start' $Matches[1]
   if($finish -gt $start){Set-LocalValue $item 'planned' ($finish.TotalMinutes-$start.TotalMinutes)}
  }
  if($item.start){
   Set-LocalValue $item 'time' ($item.start+'—'+(Get-AgendaEndTime $item.start $item.planned));Set-LocalValue $item 'detail' $item.time;Set-LocalValue $item 'kind' 'task';$target.schedule+=@($item)
  }else{Set-LocalValue $item 'time' '';Set-LocalValue $item 'detail' '';Set-LocalValue $item 'kind' ''; $target.tasks+=@($item)}
 }
 Sort-LocalSchedule $target;Sort-LocalTasks $target
}
function Restore-PendingAgenda($target,$previous,[object[]]$operations){
 if(-not $operations.Count){return}
 foreach($operation in $operations){
  $id=$operation.id
  $item=$operation.snapshot
  if(-not $item -and $previous){$item=@($previous.schedule)+@($previous.tasks)+@($previous.weekly)|Where-Object {$_.id -eq $id}|Select-Object -First 1}
  if($id -notlike 'class:*'){
   $current=@(@($target.schedule)+@($target.tasks)+@($target.weekly)|Where-Object {$_.id -eq $id})|Select-Object -First 1
   if($current){$item=$current}
   if($item){
    $keys=if($operation.fields -is [System.Collections.IDictionary]){@($operation.fields.Keys)}else{@($operation.fields.PSObject.Properties.Name)}
    foreach($key in $keys){if($key -in @('delete','deleteScope')){continue};if($key -eq 'scheduleDate'){Set-LocalValue $item 'plannedDate' $operation.fields.$key;if($id -like 'event:*'){Set-LocalValue $item 'agendaDate' $operation.fields.$key}}else{Set-LocalValue $item $key $operation.fields.$key}}
    if($item.start){Set-LocalValue $item 'time' ($item.start+'—'+(Get-AgendaEndTime $item.start $item.planned));Set-LocalValue $item 'kind' 'task'}else{Set-LocalValue $item 'time' '';Set-LocalValue $item 'kind' ''}
    Set-LocalValue $item 'detail' (Get-AgendaTaskDetail $item)
   }
  }
  if($id -like 'class:*'){
   $current=@($target.schedule|Where-Object {$_.id -eq $id})|Select-Object -First 1
   if($current){
    $item=$current;$fields=@{}
    if($operation.fields -is [System.Collections.IDictionary]){foreach($key in $operation.fields.Keys){$fields[$key]=$operation.fields[$key]}}else{foreach($property in $operation.fields.PSObject.Properties){$fields[$property.Name]=$property.Value}}
    foreach($key in $fields.Keys){
     if($key -eq 'title'){
      $value=[string]$fields[$key]
      if(-not $value.Trim()){$value=[string]$item.taskTitle}
      if($value.Trim()){Set-LocalValue $item 'taskTitle' $value;Set-LocalValue $item 'detail' ('课内：'+$value);Set-LocalValue $operation.fields 'title' $value}
      else{Set-LocalValue $item 'taskTitle' '';Set-LocalValue $item 'detail' '';if($operation.fields -is [System.Collections.IDictionary]){$operation.fields.Remove('title')}else{$operation.fields.PSObject.Properties.Remove('title')}}
     }else{Set-LocalValue $item $key $fields[$key]}
    }
   }
  }
  if($operation.fields.delete -and $id -like 'class:*'){Set-LocalValue $item 'taskTitle' '';Set-LocalValue $item 'detail' '';Set-LocalValue $item 'done' $false;Set-LocalValue $item 'incomplete' $false;Set-LocalValue $item 'highlight' $false;Set-LocalValue $item 'planned' '';Set-LocalValue $item 'actual' ''}
  elseif($operation.fields.delete){$target.schedule=@($target.schedule|Where-Object {$_.id -ne $id});$target.tasks=@($target.tasks|Where-Object {$_.id -ne $id});if($id -like 'weekly:*'){$target.weekly=@($target.weekly|Where-Object {$_.id -ne $id})};continue}
  if(-not $item){continue}
  $target.schedule=@($target.schedule|Where-Object {$_.id -ne $id});$target.tasks=@($target.tasks|Where-Object {$_.id -ne $id})
  if($id -like 'weekly:*'){
   $itemWeek=($id -split ':')[1];$belongsToWeek=if($target.weekName){$target.weekName -eq $itemWeek}else{$target.weeklyDate -eq (Get-AgendaWeekStart $item)}
   if($belongsToWeek){$found=$false;$target.weekly=@($target.weekly|ForEach-Object {if($_.id -eq $id){$found=$true;$item}else{$_}});if(-not $found){$target.weekly+=@($item)}}
   $date=$item.plannedDate
   if($date -and @($target.days|Where-Object {$_.date -eq $date}).Count){Set-LocalValue $item 'agendaDate' $date;if($item.start){Set-LocalValue $item 'time' ($item.start+'—'+(Get-AgendaEndTime $item.start $item.planned));Set-LocalValue $item 'kind' 'task';$target.schedule+=@($item)}else{$target.tasks+=@($item)}}
  }else{
   $date=if($item.agendaDate){$item.agendaDate}else{$operation.date}
   if($id -match '^(local|daily|class|habit|semester|event):'){
    if(@($target.days|Where-Object {$_.date -eq $date}).Count){Set-LocalValue $item 'agendaDate' $date;if($id -like 'class:*' -or $item.start){$target.schedule+=@($item)}else{$target.tasks+=@($item)}}
   }
  }
 }
 Sort-LocalSchedule $target;Sort-LocalTasks $target;Sort-LocalWeekly $target
}
function Move-View([int]$Offset){
 try{
  if($script:page -eq 3){$current=[datetime]::ParseExact($script:anchorDate,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture);$target=$current.AddDays($Offset*7).ToString('yyyy-MM-dd');$script:anchorDate=$target;$script:followToday=($target -eq [DateTime]::UtcNow.AddHours(8).ToString('yyyy-MM-dd'));$script:window.FindName('DatePicker').SelectedDate=[datetime]::ParseExact($target,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture);Refresh-Agenda $target;return}
  $next=$script:page+$Offset;if($next -ge 0 -and $next -le 2){Set-Page $next;return}
  $current=[datetime]::ParseExact($script:anchorDate,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture);$delta=if($Offset -gt 0){3}else{-3};$target=$current.AddDays($delta).ToString('yyyy-MM-dd');$script:anchorDate=$target;$script:followToday=($target -eq [DateTime]::UtcNow.AddHours(8).ToString('yyyy-MM-dd'));$script:page=if($Offset -gt 0){0}else{2};$script:window.FindName('DatePicker').SelectedDate=[datetime]::ParseExact($target,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture);Refresh-Agenda $target
 }catch{$script:statusLabel.Text='日期切换失败'}
}
function Repair-LegacyWeeklyScheduleDeletes($target,[object[]]$operations){
 $repaired=0
 foreach($operation in $operations){
  if($operation.id -notlike 'weekly:*' -or -not $operation.fields.delete -or $operation.fields.deleteScope -or -not $operation.snapshot.plannedDate){continue}
  $fields=@{};if($operation.fields -is [System.Collections.IDictionary]){foreach($key in $operation.fields.Keys){$fields[$key]=$operation.fields[$key]}}else{foreach($property in $operation.fields.PSObject.Properties){$fields[$property.Name]=$property.Value}}
  $fields.Remove('delete');$fields.scheduleDate='';$fields.start='';Set-LocalValue $operation 'fields' $fields
  foreach($name in @('plannedDate','start','time','kind')){Set-LocalValue $operation.snapshot $name ''};Set-LocalValue $operation.snapshot 'detail' '本周未安排';$repaired++
 }
 if($repaired){Restore-PendingAgenda $target $target $operations}
 return $repaired
}
function Get-AgendaExcelVersion {
 return (([IO.File]::GetLastWriteTimeUtc($script:planPath).Ticks-621355968000000000L)*100L).ToString([Globalization.CultureInfo]::InvariantCulture)
}
function Refresh-AgendaIfChanged {
 if(-not $script:data -or $script:busy -or $script:refreshing -or $script:syncProcess){return}
 # 跨过午夜时 Excel 往往一个字节都没变，只靠版本比对永远不会刷新，必须单独判断日期
 $today=[DateTime]::UtcNow.AddHours(8).ToString('yyyy-MM-dd')
 if($script:followToday -and $script:anchorDate -ne $today){Refresh-Agenda;return}
 try{if([string]$script:data.version -ne (Get-AgendaExcelVersion)){Refresh-Agenda $script:anchorDate -ReloadExcel}}catch{$script:statusLabel.Text='Excel 暂不可读取 · 本地修改已保留'}
}
function Get-AgendaAnchorDate([string]$RequestedDate=''){
 # 跟随今天（默认）时忽略外部传来的旧锚点：常驻跨过午夜后 anchorDate 不会自己前进，
 # 不纠正的话一整天都在显示昨天的数据，周任务也停在上一周
 $today=[DateTime]::UtcNow.AddHours(8).ToString('yyyy-MM-dd')
 if($PreviewDate){if($RequestedDate){return $RequestedDate};if($script:anchorDate){return $script:anchorDate};return $PreviewDate}
 if($script:followToday){return $today}
 if($RequestedDate){return $RequestedDate}
 if($script:anchorDate){return $script:anchorDate}
 return $today
}
function Refresh-Agenda([string]$RequestedDate='',[switch]$ReloadExcel) {
 if($script:refreshing){$script:queuedRefresh=@{date=$RequestedDate;reload=[bool]$ReloadExcel};return}
 Complete-AgendaCheckboxClick
 $script:refreshing=$true
 try{
  $day=Get-AgendaAnchorDate $RequestedDate;$script:anchorDate=$day;$script:lastDate=$day
  $datePicker=$script:window.FindName('DatePicker');$selectedDate=[datetime]::ParseExact($day,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture);if($datePicker.SelectedDate -ne $selectedDate){$datePicker.SelectedDate=$selectedDate}
  $loadedCache=$false
  $repairedWeekly=0
  if(-not $ReloadExcel -and (Test-Path -LiteralPath $script:cachePath)){$cache=Get-Content -LiteralPath $script:cachePath -Raw -Encoding UTF8|ConvertFrom-Json;if($cache.migrations.weeklyUnschedule -ne 1){$repairedWeekly=Repair-LegacyWeeklyScheduleDeletes $cache.data @($cache.pendingOps)};if($repairedWeekly){$backup=$script:cachePath+'.before-weekly-unschedule.json';if(-not (Test-Path -LiteralPath $backup)){Copy-Item -LiteralPath $script:cachePath -Destination $backup}};if($cache.pendingOps.Count){$script:pendingOps=@($cache.pendingOps);$script:baseVersion=if($cache.version){$cache.version}else{$cache.data.version};$script:baseDigest=if($cache.digest){$cache.digest}else{$cache.data.digest}};if($cache.date -eq $day){$script:data=$cache.data;$script:pendingOps=@($cache.pendingOps);if(-not $script:baseVersion){$script:baseVersion=if($cache.version){$cache.version}else{$cache.data.version};$script:baseDigest=if($cache.digest){$cache.digest}else{$cache.data.digest}};$script:lastDate=$day;$loadedCache=$true}}
  if($loadedCache -and $script:pendingOps.Count -and [string]$script:data.version -eq (Get-AgendaExcelVersion)){Normalize-CachedHabits $script:data;if($repairedWeekly -or $cache.migrations.weeklyUnschedule -ne 1){Write-AgendaCache};Render-Page;$script:statusLabel.Text='本地缓存 · 待同步 '+$script:pendingOps.Count+' 项';return}
  $anchor=[datetime]::ParseExact($day,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture);$days=@();$merged=$null
  $script:statusLabel.Text='正在读取 Excel…'
  [Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action]{},[Windows.Threading.DispatcherPriority]::Render)
  $range=Invoke-Backend ('--range --date '+$day) -Responsive;if($range.error){throw $range.error}
  if($script:queuedRefresh){return}
  Complete-AgendaCheckboxClick
  for($offset=0;$offset -lt 3;$offset++){
   $date=$anchor.AddDays($offset).ToString('yyyy-MM-dd');$part=$range.parts[$offset];if(-not $part -or $part.error){throw '未能读取连续三天的安排。'};$partSchedule=@($part.schedule);$partTasks=@($part.tasks);$partHabits=@($part.habits)
   if(-not $merged){$merged=$part;$merged.schedule=@();$merged.tasks=@();$merged.habits=@();Set-LocalValue $merged 'days' @()}
   $days+=@([PSCustomObject]@{date=$date;title=$part.title})
   foreach($entry in $partSchedule){Set-LocalValue $entry 'agendaDate' $date;$merged.schedule+=@($entry)}
   foreach($entry in $partTasks){Set-LocalValue $entry 'agendaDate' $date;$merged.tasks+=@($entry)}
   foreach($entry in $partHabits){Set-LocalValue $entry 'agendaDate' $date;$merged.habits+=@($entry)}
  }
  $merged.date=$day;$merged.title=$days[0].title;Set-LocalValue $merged 'days' $days;Set-LocalValue $merged 'rangeTitle' ($days[0].title+' — '+$days[2].title)
  # 读取期间（消息循环已让出）可能又产生了新的本地修改，必须以完成时刻的队列为准回放，不能用开始时的快照
  $currentOps=@($script:pendingOps)
  if(-not $currentOps.Count){$script:baseVersion=$merged.version;$script:baseDigest=$merged.digest}
  Restore-PendingAgenda $merged $script:data $currentOps;$script:data=$merged;$script:pendingOps=$currentOps;Write-AgendaCache;Render-Page
 }catch{$script:statusLabel.Text='刷新未完成 · 本地修改已保留';$script:statusLabel.ToolTip=$_.Exception.Message;[IO.File]::WriteAllText((Join-Path $script:baseDir '错误.txt'),($_|Out-String),[Text.Encoding]::UTF8)}
 finally{
  $script:refreshing=$false
  if($script:queuedRefresh){$request=$script:queuedRefresh;$script:queuedRefresh=$null;Refresh-Agenda $request.date -ReloadExcel:$request.reload}
 }
}
function Get-AgendaOperationSignature($operation){
 $fields=[ordered]@{}
 $keys=if($operation.fields -is [System.Collections.IDictionary]){@($operation.fields.Keys)}else{@($operation.fields.PSObject.Properties.Name)}
 foreach($key in @($keys|Sort-Object)){$fields[$key]=$operation.fields.$key}
 return ([ordered]@{id=$operation.id;date=$operation.date;fields=$fields}|ConvertTo-Json -Depth 8 -Compress)
}
function Complete-AgendaSync($result){
 $script:pendingOps=@($script:pendingOps|Where-Object {-not $script:syncSignatures.ContainsKey((Get-AgendaOperationSignature $_))})
 $mappings=@{}
 if($result.mappings -is [System.Collections.IDictionary]){foreach($key in $result.mappings.Keys){$mappings[$key]=$result.mappings[$key]}}
 elseif($result.mappings){foreach($property in $result.mappings.PSObject.Properties){$mappings[$property.Name]=$property.Value}}
 foreach($item in @($script:data.schedule)+@($script:data.tasks)+@($script:data.habits)+@($script:data.weekly)+@($script:pendingOps|ForEach-Object {$_.snapshot})){
  if($item -and $mappings.ContainsKey([string]$item.id)){$item.id=$mappings[[string]$item.id]}
 }
 foreach($operation in $script:pendingOps){if($mappings.ContainsKey([string]$operation.id)){$operation.id=$mappings[[string]$operation.id]}}
 if($result.version){$script:baseVersion=[string]$result.version;$script:baseDigest=if($result.digest){[string]$result.digest}else{''};Set-LocalValue $script:data 'version' $script:baseVersion;Set-LocalValue $script:data 'digest' $script:baseDigest;foreach($operation in $script:pendingOps){Set-LocalValue $operation 'version' $script:baseVersion}}
}
. (Join-Path $script:baseDir 'conflicts.ps1')
function Start-AgendaSync {
 Complete-AgendaCheckboxClick
 if($script:refreshing){$script:statusLabel.Text='正在读取 Excel，请稍候';return}
 if($script:syncProcess -and -not $script:syncProcess.HasExited){$script:statusLabel.Text='后台同步进行中…';return}
 if(-not $script:pendingOps.Count){Start-Process -FilePath $script:planPath;return}
 try{
  $token=[guid]::NewGuid().ToString('N');$inputPath=$script:planPath+'.sync-'+$token+'.json';$resultPath=$script:planPath+'.sync-'+$token+'.result.json'
  $ops=@();$script:syncSignatures=@{};$script:syncIds=@{};$expectedVersions=@($script:pendingOps|ForEach-Object {if($_.version){[string]$_.version}else{[string]$script:baseVersion}}|Select-Object -Unique);if($expectedVersions.Count -ne 1 -or -not $expectedVersions[0]){throw '本地修改缺少一致的 Excel 基线版本，请先重新载入计划并核对。'};$expectedDigest=if($script:baseDigest){[string]$script:baseDigest}else{[string]$script:data.digest};foreach($operation in $script:pendingOps){$fields=@{};if($operation.fields -is [System.Collections.IDictionary]){foreach($key in $operation.fields.Keys){$fields[$key]=$operation.fields[$key]}}else{foreach($property in $operation.fields.PSObject.Properties){$fields[$property.Name]=$property.Value}};$ops+=@(@{id=$operation.id;date=$operation.date;fields=$fields});$script:syncSignatures[(Get-AgendaOperationSignature $operation)]=$true;$script:syncIds[$operation.id]=$true}
  [IO.File]::WriteAllText($inputPath,(@{version=$expectedVersions[0];digest=$expectedDigest;ops=$ops}|ConvertTo-Json -Depth 8 -Compress),[System.Text.UTF8Encoding]::new($false))
  $worker=Join-Path $script:baseDir 'sync_worker.ps1';$powershell=(Get-Command powershell.exe -ErrorAction Stop).Source
  $arguments='-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "'+$worker+'" -Book "'+$script:planPath+'" -Python "'+$script:pythonPath+'" -InputFile "'+$inputPath+'" -ResultFile "'+$resultPath+'"'
  $script:syncInputPath=$inputPath;$script:syncResultPath=$resultPath;$script:window.FindName('OpenPlanButton').IsEnabled=$false;$script:statusLabel.Text='正在后台同步 Excel…'
  $script:syncStarted=[DateTime]::UtcNow;$script:syncProcess=Start-Process -FilePath $powershell -ArgumentList $arguments -PassThru -WindowStyle Hidden
  $script:syncTimer=New-Object Windows.Threading.DispatcherTimer;$script:syncTimer.Interval=[TimeSpan]::FromMilliseconds(350)
  $script:syncTimer.Add_Tick({
   if(-not $script:syncProcess.HasExited){$script:statusLabel.Text='后台同步 Excel · 已用 '+[int]([DateTime]::UtcNow-$script:syncStarted).TotalSeconds+' 秒';$script:statusLabel.ToolTip='正在完整保存工作簿，期间可继续安排任务；新的修改会保留到下次同步。';return}
   $script:syncTimer.Stop()
   $conflict=$false
   try{
    if(-not (Test-Path -LiteralPath $script:syncResultPath)){throw '后台同步进程未返回结果。'}
    $result=Get-Content -LiteralPath $script:syncResultPath -Raw -Encoding UTF8|ConvertFrom-Json
    if($result.code -eq 'VERSION_CONFLICT'){$conflict=$true}
    else{
     if(-not $result.ok){throw $result.error}
     Complete-AgendaSync $result
     Write-AgendaCache;$script:statusLabel.Text=if($script:pendingOps.Count){'本次已同步 · 新改动待同步 '+$script:pendingOps.Count+' 项'}else{'已与 Excel 同步 '+(Get-Date).ToString('HH:mm')}
     if(-not $script:pendingOps.Count){Start-Process -FilePath $script:planPath}
    }
   }catch{$script:statusLabel.Text='同步未完成 · 本地修改已保留';[void][Windows.MessageBox]::Show($_.Exception.Message+"`n本地 JSON 已保留；请核对 Excel 与本地内容后再重试。",'Excel 同步冲突或失败','OK','Warning')}
   finally{$script:window.FindName('OpenPlanButton').IsEnabled=$true;Remove-Item -LiteralPath $script:syncInputPath,$script:syncResultPath -ErrorAction SilentlyContinue;$script:syncProcess.Dispose();$script:syncProcess=$null;$script:syncIds=@{}}
   if($conflict){Show-AgendaSyncReview}
  })
  $script:syncTimer.Start()
 }catch{$script:syncIds=@{};$script:window.FindName('OpenPlanButton').IsEnabled=$true;$script:statusLabel.Text='同步未启动';[void][Windows.MessageBox]::Show($_.Exception.Message,'同步未启动','OK','Information')}
}
