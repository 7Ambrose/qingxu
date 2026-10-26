$ErrorActionPreference='Stop'
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase
$script:baseDir=$PSScriptRoot
$qaPath=Join-Path (Split-Path -Parent $PSScriptRoot) 'outputs\study-plan\qa'
[void][IO.Directory]::CreateDirectory($qaPath)
$script:planPath=Join-Path $qaPath ('qingxu-regression-'+[guid]::NewGuid().ToString('N')+'.xlsx')
$script:cachePath=$script:planPath+'.agenda.json'
. (Join-Path $PSScriptRoot 'controls.ps1')
. (Join-Path $PSScriptRoot 'pages.ps1')
function Render-Page {}
$script:statusLabel=New-Object Windows.Controls.TextBlock
$failures=New-Object 'System.Collections.Generic.List[string]'
function Reset-TestAgenda {
 $script:pendingOps=@();$script:syncIds=@{};$script:baseVersion='test-version';$script:busy=$false;$script:page=0;$script:lastDate='2026-10-19';$script:dragScheduleDate='';$script:followToday=$true
 $script:data=[PSCustomObject]@{version='test-version';schedule=@();tasks=@();habits=@();weekly=@();weeklyDate='2026-10-19';weekName='秋6周';days=@([PSCustomObject]@{date='2026-10-19'},[PSCustomObject]@{date='2026-10-20'},[PSCustomObject]@{date='2026-10-21'})}
}
function Test-Regression([string]$name,[scriptblock]$body){
 Reset-TestAgenda
 try{& $body;Write-Output ('PASS: '+$name)}catch{$failures.Add($name+': '+$_.Exception.Message);Write-Output ('FAIL: '+$name)}
}
try{
 Test-Regression 'Class checkbox updates the course-backed assignment' {
  $course=[PSCustomObject]@{id='class:6';title='示例课程';taskTitle='课程作业';detail='课内：课程作业';done=$false;highlight=$false;agendaDate='2026-10-19'};$script:data.schedule=@($course)
  $card=[PSCustomObject]@{id=$course.id;title=$course.taskTitle;taskTitle=$course.taskTitle;detail='课内安排';done=$false;highlight=$false;agendaDate=$course.agendaDate}
  $saved=Save-Item $card @{done=$true};if(-not $saved -or -not $script:data.schedule[0].done){throw ('The rendered course would still show unchecked: '+($script:data|ConvertTo-Json -Depth 4 -Compress)+' saved='+$saved)}
 }
 Test-Regression 'Weekly time drop retains the visible daily page date' {
  $script:page=1;$task=[PSCustomObject]@{id='weekly:秋6周:50';title='周任务';planned=30;start='';plannedDate='2026-10-20';agendaDate='2026-10-20'};$script:data.weekly=@($task);$script:data.tasks=@($task)
  $saved=Set-AgendaTaskTime $task '13:30';if(-not $saved -or $task.plannedDate -ne '2026-10-20'){throw ('A second-day drop was scheduled on the anchor day: '+($task|ConvertTo-Json -Compress)+' saved='+$saved)}
 }
 Test-Regression 'Weekly slot conflict uses the selected destination date' {
  $script:page=3;$script:dragScheduleDate='2026-10-20';$task=[PSCustomObject]@{id='weekly:秋6周:50';title='周任务';planned=30;start='';agendaDate='2026-10-19'}
  $script:data.schedule=@([PSCustomObject]@{id='class:target';agendaDate='2026-10-20';time='13:30—14:30'})
  if(Test-AgendaSlotAvailable $task '13:30'){throw 'The destination-day class conflict was ignored'}
 }
 Test-Regression 'Semester tasks stay editable after pending-cache recovery' {
  $task=[PSCustomObject]@{id='semester:5';title='学期任务';done=$true;agendaDate='2026-10-19'};$script:data.tasks=@($task)
  Restore-PendingAgenda $script:data $null @([PSCustomObject]@{id=$task.id;date=$task.agendaDate;fields=@{done=$true};snapshot=$task})
  if(@($script:data.tasks|Where-Object id -eq $task.id).Count -ne 1){throw 'Recovered task was converted into a non-editable schedule row'}
 }
 Test-Regression 'Deleting an in-flight new task retains a deletion operation' {
  $task=[PSCustomObject]@{id='local:in-flight';title='同步中新增任务';agendaDate='2026-10-19'};$script:data.tasks=@($task);Add-PendingOperation $task.id $task.agendaDate @{title=$task.title} $task;$script:syncIds[$task.id]=$true
  $script:syncSignatures=@{};$script:syncSignatures[(Get-AgendaOperationSignature $script:pendingOps[0])]=$true
  $saved=Remove-AgendaItem $task $false;if(-not $saved -or @($script:pendingOps|Where-Object {$_.id -eq $task.id -and $_.fields.delete}).Count -ne 1){throw ('The worker could recreate the deleted task without a pending deletion: '+($script:pendingOps|ConvertTo-Json -Depth 4 -Compress)+' saved='+$saved)}
  Complete-AgendaSync ([PSCustomObject]@{mappings=@{'local:in-flight'='daily:300'};version='after-sync'})
  if($script:pendingOps.Count -ne 1 -or $script:pendingOps[0].id -ne 'daily:300' -or -not $script:pendingOps[0].fields.delete){throw 'Deletion was not mapped to the newly persisted task'}
 }
 Test-Regression 'Sync maps edited tasks even after browsing away' {
  $task=[PSCustomObject]@{id='local:hidden';title='Before';agendaDate='2026-10-19'};Add-PendingOperation $task.id $task.agendaDate @{title='Before'} $task
  $script:syncSignatures=@{};$script:syncSignatures[(Get-AgendaOperationSignature $script:pendingOps[0])]=$true
  $script:pendingOps[0].fields.title='After';$task.title='After'
  Complete-AgendaSync ([PSCustomObject]@{mappings=@{'local:hidden'='daily:301'};version='after-sync'})
  if($script:pendingOps.Count -ne 1 -or $script:pendingOps[0].id -ne 'daily:301' -or $script:pendingOps[0].snapshot.id -ne 'daily:301' -or $script:pendingOps[0].version -ne 'after-sync'){throw 'An out-of-view edit would be inserted as a duplicate task'}
 }
 Test-Regression 'Input validation normalizes times and rejects invalid local operations' {
  $task=[PSCustomObject]@{id='local:validation';start='';planned=''};$fields=Get-ValidatedAgendaFields $task @{title='Task';start='8:30';planned=30}
  if($fields.start -ne '08:30'){throw 'A short hour would sort incorrectly'}
  foreach($invalid in @(@{title=' '},@{start='25:00';planned=30},@{start='23:30';planned=60},@{planned='NaN'})){$rejected=$false;try{[void](Get-ValidatedAgendaFields $task $invalid)}catch{$rejected=$true};if(-not $rejected){throw 'Invalid task fields reached the pending queue'}}
  $task|Add-Member -NotePropertyName agendaDate -NotePropertyValue '2027-01-18';$rejected=$false;try{[void](Get-ValidatedAgendaFields $task @{title='Out of range'})}catch{$rejected=$true};if(-not $rejected){throw 'An unsupported date reached the pending queue'}
  if((ConvertTo-AgendaMinutes '24:00' $true) -ne 1440){throw 'Midnight end time should remain valid'}
  if((Get-AgendaEndTime '23:30' 30) -ne '24:00'){throw 'The midnight end time wrapped back to 00:00'}
 }
 Test-Regression 'Weekly scheduling validates the actual week before mutating the task' {
  $task=[PSCustomObject]@{id='weekly:秋6周:50';start='';planned=30};$rejected=$false
  try{[void](Get-ValidatedAgendaFields $task @{scheduleDate='2026-10-26'})}catch{$rejected=$true}
  if(-not $rejected){throw 'A next-week date reached the pending queue'}
 }
 Test-Regression 'Routine order survives normalization and pending-cache recovery' {
  $habit=[PSCustomObject]@{id='habit:秋6周:B:67';title='背单词';planned=30;start='';order=2;agendaDate='2026-10-19'};$ordinary=[PSCustomObject]@{id='daily:15';title='普通任务';order=1;agendaDate='2026-10-19'};$script:data.tasks=@($habit,$ordinary)
  Normalize-CachedHabits $script:data
  if($script:data.tasks[0].id -ne $ordinary.id -or $script:data.tasks[1].id -ne $habit.id){throw 'Routine normalization changed the saved task order'}
 }
 Test-Regression 'Pending weekly edits stay out of another week view' {
  $script:data.weekName='秋7周';$script:data.weeklyDate='2026-10-26';$task=[PSCustomObject]@{id='weekly:秋6周:50';title='上周任务';order=1;plannedDate='2026-10-20';start=''}
  Restore-PendingAgenda $script:data $null @([PSCustomObject]@{id=$task.id;date='2026-10-19';fields=@{done=$true};snapshot=$task})
  if($script:data.weekly.Count){throw 'An edited task from the previous week appeared in the current weekly list'}
 }
 Test-Regression 'Weekly reorder survives pending-cache recovery' {
  $first=[PSCustomObject]@{id='weekly:秋6周:50';title='First';order=1};$second=[PSCustomObject]@{id='weekly:秋6周:51';title='Second';order=2};$script:data.weekly=@($first,$second)
  $moved=[PSCustomObject]@{id=$second.id;title=$second.title;order=1};$other=[PSCustomObject]@{id=$first.id;title=$first.title;order=2}
  Restore-PendingAgenda $script:data $null @([PSCustomObject]@{id=$moved.id;date='2026-10-19';fields=@{order=1};snapshot=$moved},[PSCustomObject]@{id=$other.id;date='2026-10-19';fields=@{order=2};snapshot=$other})
  if($script:data.weekly[0].id -ne $second.id){throw 'Browsing restored the old Excel row order over local weekly order'}
 }
 Test-Regression 'New task gets a valid order on a day containing only default routines' {
  $script:data.tasks=@([PSCustomObject]@{id='habit:秋6周:B:67';title='Words';order=-2;agendaDate='2026-10-19'})
  if(-not (Save-Item ([PSCustomObject]@{id='new';agendaDate='2026-10-19'}) @{title='First ordinary task';start='';planned=''})){throw 'Could not create the first ordinary task'}
  if($script:pendingOps[0].fields.order -lt 1){throw 'A default routine gave the new task an invalid Excel order'}
 }
 Test-Regression 'Clearing a dragged time removes its cached task-card subtitle' {
  $task=[PSCustomObject]@{id='daily:300';title='预习实验报告';start='21:30';planned=30;time='21:30—22:00';detail='21:30—22:00';kind='task';agendaDate='2026-10-19'};$script:data.schedule=@($task)
  if(-not (Save-Item $task @{start='';planned=''})){throw 'Could not clear the task time'}
  if($task.time -or $task.detail -eq '21:30—22:00'){throw 'The removed time remained in the task-card subtitle'}
  $script:content=New-Object Windows.Controls.StackPanel;Add-Task $task
  if($script:content.Children[0].Child.Children[1].Children.Count -ne 1){throw 'The untimed card still rendered a subtitle'}
 }
 Test-Regression 'Existing stale subtitles are hidden while weekly dates remain visible' {
  foreach($identity in @('daily:300','local:old-cache','habit:秋6周:B:67')){$task=[PSCustomObject]@{id=$identity;start='';time='';detail='21:30—22:00'};if(Get-AgendaTaskDetail $task){throw 'An existing cache still displayed its cancelled time'}}
  $weekly=[PSCustomObject]@{id='weekly:秋6周:50';start='';time='';detail='21:30—22:00';plannedDate='2026-10-19'}
  if((Get-AgendaTaskDetail $weekly) -ne '安排：10月19日'){throw 'Clearing a weekly time removed its scheduled date'}
  $weekly.start='21:30';$weekly.time='21:30—22:00';if((Get-AgendaTaskDetail $weekly) -ne $weekly.time){throw 'A timed task lost its current time subtitle'}
 }
 Test-Regression 'Daily removal unschedules weekly tasks without deleting them' {
  $task=[PSCustomObject]@{id='weekly:秋6周:50';title='每周练习';planned=30;start='13:30';plannedDate='2026-10-20';agendaDate='2026-10-20';done=$true};$script:data.weekly=@($task);$script:data.schedule=@($task);$script:page=1
  if(-not (Remove-AgendaItem $task $false)){throw 'Could not cancel the daily assignment'}
  if($script:data.weekly.Count -ne 1 -or $task.plannedDate -or $task.start -or -not $task.done -or $task.planned -ne 30){throw 'Unscheduling lost the underlying weekly task or its state'}
  if($script:data.schedule.Count -or $script:data.tasks.Count -or $script:pendingOps[0].fields.delete){throw 'Unscheduling deleted the weekly task or left a daily copy'}
  if(-not (Save-Item $task @{scheduleDate='2026-10-21';start='15:30'})){throw 'Could not reschedule the retained weekly task'}
  if($task.plannedDate -ne '2026-10-21' -or $script:data.weekly.Count -ne 1){throw 'Rescheduling duplicated or lost the task'}
 }
 Test-Regression 'Weekly list deletion remains explicit and is not migrated' {
  $script:page=3;$task=[PSCustomObject]@{id='weekly:秋6周:50';title='Delete';plannedDate='2026-10-20';start='13:30'};$script:data.weekly=@($task)
  if(-not (Remove-AgendaItem $task $false)){throw 'Could not delete from the weekly list'}
  if($script:data.weekly.Count -or -not $script:pendingOps[0].fields.delete -or $script:pendingOps[0].fields.deleteScope -ne 'weekly'){throw 'Weekly deletion was not marked explicitly'}
  if((Repair-LegacyWeeklyScheduleDeletes $script:data $script:pendingOps) -ne 0){throw 'A deliberate weekly deletion was restored'}
 }
 Test-Regression 'Legacy pending weekly deletions recover as unscheduling once' {
  $task=[PSCustomObject]@{id='weekly:秋6周:50';title='每周练习';plannedDate='2026-10-20';start='13:30';planned=30;done=$true}
  $operations=@([PSCustomObject]@{id=$task.id;date='2026-10-19';fields=@{delete=$true};snapshot=$task})
  if((Repair-LegacyWeeklyScheduleDeletes $script:data $operations) -ne 1){throw 'Legacy deletion was not repaired'}
  if($script:data.weekly.Count -ne 1 -or $script:data.weekly[0].plannedDate -or -not $script:data.weekly[0].done -or $operations[0].fields.delete){throw 'Recovered weekly task lost its state or remained deleted'}
  if((Repair-LegacyWeeklyScheduleDeletes $script:data $operations) -ne 0){throw 'Migration was not idempotent'}
 }
 Test-Regression 'Editing a class child updates its assignment but preserves the course title' {
  $course=[PSCustomObject]@{id='class:6';title='示例课程';taskTitle='旧作业';done=$false;agendaDate='2026-10-19'};$script:data.schedule=@($course)
  $card=[PSCustomObject]@{id=$course.id;title=$course.taskTitle;agendaDate=$course.agendaDate}
  if(-not (Save-Item $card @{title='新作业'})){throw 'Could not rename the assignment'}
  if($course.title -ne '示例课程' -or $course.taskTitle -ne '新作业'){throw 'Editing changed the course name or left a stale assignment'}
 }
 Test-Regression 'Fresh Excel fields survive pending checkbox recovery' {
  $fresh=[PSCustomObject]@{id='event:6';title='Excel new title';start='09:30';planned=120;agendaDate='2026-10-20';done=$false}
  $stale=[PSCustomObject]@{id='event:6';title='Old title';start='08:00';planned=30;agendaDate='2026-10-19';done=$true}
  $script:data.schedule=@($fresh)
  Restore-PendingAgenda $script:data $null @([PSCustomObject]@{id=$fresh.id;date='2026-10-19';fields=@{done=$true};snapshot=$stale})
  $actual=$script:data.schedule[0]
  if($actual.title -ne 'Excel new title' -or $actual.time -ne '09:30—11:30' -or $actual.agendaDate -ne '2026-10-20' -or -not $actual.done){throw 'A stale snapshot hid the changed Excel fields'}
 }
 Test-Regression 'Explicit incomplete state is exclusive and survives recovery' {
  $task=[PSCustomObject]@{id='daily:300';title='Status test';done=$true;agendaDate='2026-10-19'};$script:data.tasks=@($task)
  if(-not (Save-Item $task @{incomplete=$true})){throw 'Could not mark incomplete'}
  if($task.done -or -not $task.incomplete){throw 'Completion and incomplete are not exclusive'}
  $script:content=New-Object Windows.Controls.StackPanel;Add-Task $task
  $card=$script:content.Children[0].Child.Children[1]
  $check=$card.Children[0];$check.ApplyTemplate()|Out-Null
  if($card.Children.Count -ne 1 -or $null -ne $check.IsChecked -or $check.Template.FindName('Slash',$check).Visibility -ne 'Visible'){throw 'Incomplete should use a slash without extra text'}
  $fresh=[PSCustomObject]@{id=$task.id;title='Fresh title';done=$false;agendaDate=$task.agendaDate};$script:data.tasks=@($fresh)
  Restore-PendingAgenda $script:data $null $script:pendingOps
  if(-not $fresh.incomplete -or $fresh.done -or $fresh.title -ne 'Fresh title'){throw 'Recovery lost the explicit incomplete state'}
  if(-not (Save-Item $fresh @{done=$true}) -or $fresh.incomplete -or -not $fresh.done){throw 'Checking complete did not clear incomplete'}
  if(-not (Save-Item $fresh @{done=$false}) -or $fresh.incomplete -or $fresh.done){throw 'Clearing status failed'}
 }
 Test-Regression 'Double click cancels the pending single click and uses a gray slash' {
  $task=[PSCustomObject]@{id='daily:300';title='Double click test';done=$false;agendaDate='2026-10-19'};$script:data.tasks=@($task)
  $script:content=New-Object Windows.Controls.StackPanel;Add-Task $task
  $check=$script:content.Children[0].Child.Children[1].Children[0];$check.IsChecked=$true
  Invoke-AgendaCheckboxClick $check
  if(-not $script:checkboxClick -or $task.done){throw 'Single click was committed before the double-click window'}
  Invoke-AgendaCheckboxClick $check $true
  if($script:checkboxClick -or $task.done -or -not $task.incomplete){throw 'Double click failed to replace the pending single click'}
  Complete-AgendaCheckboxClick
  if($task.done -or -not $task.incomplete){throw 'A delayed single click overwrote incomplete'}
  $script:content.Children.Clear();Add-Task $task
  $border=$script:content.Children[0];$check=$border.Child.Children[1].Children[0];$check.ApplyTemplate()|Out-Null
  if($border.Background.ToString() -ne '#FFF2F3F2' -or $check.Template.FindName('Slash',$check).Visibility -ne 'Visible'){throw 'Gray slash state did not render'}
  Invoke-AgendaCheckboxClick $check;Complete-AgendaCheckboxClick
  if(-not $task.done -or $task.incomplete){throw 'Single click on incomplete did not complete it'}
 }
 Test-Regression 'Single click commits automatically after the double-click interval' {
  $task=[PSCustomObject]@{id='daily:300';title='Single click test';done=$false;agendaDate='2026-10-19'};$script:data.tasks=@($task)
  $script:content=New-Object Windows.Controls.StackPanel;Add-Task $task
  $check=$script:content.Children[0].Child.Children[1].Children[0];$check.IsChecked=$true;Invoke-AgendaCheckboxClick $check
  $script:clickFrame=New-Object Windows.Threading.DispatcherFrame
  $script:clickWait=New-Object Windows.Threading.DispatcherTimer;$script:clickWait.Interval=[TimeSpan]::FromMilliseconds([Windows.Forms.SystemInformation]::DoubleClickTime+150)
  $script:clickWait.Add_Tick({$script:clickWait.Stop();$script:clickFrame.Continue=$false});$script:clickWait.Start()
  [Windows.Threading.Dispatcher]::PushFrame($script:clickFrame)
  if(-not $task.done -or $script:checkboxClick){throw 'Single click was not saved by its timer'}
 }
 Test-Regression 'Midnight rollover re-anchors the view to today' {
  $midnightToday=[DateTime]::UtcNow.AddHours(8).ToString('yyyy-MM-dd')
  $staleDate=([datetime]::ParseExact($midnightToday,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture)).AddDays(-1).ToString('yyyy-MM-dd')
  $script:followToday=$true;$script:anchorDate=$staleDate
  $anchor=Get-AgendaAnchorDate $staleDate
  if($anchor -ne $midnightToday){throw ('A window left open overnight still anchors on '+$staleDate+' instead of '+$midnightToday+' (got '+$anchor+'), so the weekly list never advances')}
 }
 Test-Regression 'Manually selected date survives midnight rollover' {
  $manualToday=[DateTime]::UtcNow.AddHours(8).ToString('yyyy-MM-dd')
  $manualDate=([datetime]::ParseExact($manualToday,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture)).AddDays(-3).ToString('yyyy-MM-dd')
  $script:followToday=$false;$script:anchorDate=$manualDate
  if((Get-AgendaAnchorDate $manualDate) -ne $manualDate){throw 'A manually chosen date was dragged back to today'}
  if((Get-AgendaAnchorDate) -ne $manualDate){throw 'Refreshing without an argument dropped the manually chosen date'}
 }
 Test-Regression 'A delayed checkbox updates the current item after Excel refresh' {
  $stale=[PSCustomObject]@{id='daily:300';title='Old title';done=$false;agendaDate='2026-10-19'}
  $fresh=[PSCustomObject]@{id=$stale.id;title='Excel new title';done=$false;agendaDate=$stale.agendaDate};$script:data.tasks=@($fresh)
  if(-not (Save-Item $stale @{done=$true})){throw 'Checkbox could not save'}
  if(-not $fresh.done -or $script:pendingOps[0].snapshot.title -ne 'Excel new title'){throw 'Delayed click saved an obsolete item instead of current data'}
  $cached=Get-Content -LiteralPath $script:cachePath -Raw -Encoding UTF8|ConvertFrom-Json
  if(-not $cached.data.tasks[0].done){throw 'A completed checkbox became unchecked in the saved cache'}
 }
 Test-Regression 'Excel event cards expose the same drag handle as ordinary tasks' {
  $script:content=New-Object Windows.Controls.StackPanel
  Add-Task ([PSCustomObject]@{id='event:6';title='Event';done=$false;start='09:30';planned=60})
  $handles=@($script:content.Children[0].Child.Children|Where-Object {[Windows.Controls.Grid]::GetColumn($_) -eq 0})
  if($handles.Count -ne 1 -or $handles[0].Child.Text -ne '⠿'){throw 'Excel event has no usable drag handle'}
 }
 if($failures.Count){throw ($failures -join "`n")}
}finally{Remove-Item -LiteralPath $script:cachePath -ErrorAction SilentlyContinue}
