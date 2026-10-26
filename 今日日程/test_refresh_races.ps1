$ErrorActionPreference='Stop'
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase
$script:baseDir=$PSScriptRoot
$folder=Join-Path (Split-Path -Parent $PSScriptRoot) ('outputs\study-plan\qa\refresh-race-'+[guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($folder)
$script:planPath=Join-Path $folder 'test.xlsx';$script:cachePath=$script:planPath+'.agenda.json'
. (Join-Path $PSScriptRoot 'controls.ps1')
. (Join-Path $PSScriptRoot 'pages.ps1')
[xml]$xaml=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'view.xaml'))
$script:window=[Windows.Markup.XamlReader]::Load((New-Object Xml.XmlNodeReader $xaml))
$script:statusLabel=$script:window.FindName('StatusLabel');$script:content=$script:window.FindName('ContentPanel')
function Render-Page {}
function New-TestTask([string]$title){return [PSCustomObject]@{id='daily:300';title=$title;done=$false;incomplete=$false;agendaDate='2026-10-19';start='';planned=''}}
function New-TestData([string]$date){return [PSCustomObject]@{version='version';digest='digest';date=$date;title=$date;updated='00:00';schedule=@();tasks=@();habits=@();weekly=@();weekName='秋6周';weeklyDate='2026-10-19'}}
function Invoke-Backend([string]$mode,[string]$body='',[switch]$Responsive){
 $date=($mode -split ' ')[-1];$script:requests+=@($date)
 if($script:inject){$script:inject=$false;foreach($selected in @('2026-10-22','2026-10-23')){$script:anchorDate=$selected;Refresh-Agenda $selected -ReloadExcel}}
 if($script:injectClick){
  $script:injectClick=$false;$script:oldCheck=New-Object Windows.Controls.CheckBox;$script:oldCheck.Tag=$script:data.tasks[0];$script:oldCheck.IsChecked=$true
  Invoke-AgendaCheckboxClick $script:oldCheck
 }
 $anchor=[datetime]::ParseExact($date,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture)
 $parts=@();for($index=0;$index -lt 3;$index++){$parts+=@(New-TestData ($anchor.AddDays($index).ToString('yyyy-MM-dd')))}
 if($date -eq '2026-10-19'){$parts[0].tasks=@(New-TestTask 'Excel fresh title')}
 return [PSCustomObject]@{parts=$parts;version='version';digest='digest'}
}
try{
 $script:followToday=$false;$PreviewDate='';$script:pendingOps=@();$script:requests=@();$script:baseVersion='version';$script:baseDigest='digest'
 $script:data=New-TestData '2026-10-19';$script:data.tasks=@(New-TestTask 'Old title');$script:inject=$true
 Refresh-Agenda '2026-10-19' -ReloadExcel
 if($script:data.date -ne '2026-10-23' -or $script:anchorDate -ne '2026-10-23' -or $script:requests.Count -ne 2){throw 'Changing date while loading lost the latest selected date'}
 $script:data=New-TestData '2026-10-19';$script:data.tasks=@(New-TestTask 'Old title');$script:injectClick=$true;$script:pendingOps=@()
 Refresh-Agenda '2026-10-19' -ReloadExcel
 $fresh=@($script:data.tasks|Where-Object id -eq 'daily:300')[0]
 if(-not $fresh.done -or $fresh.title -ne 'Excel fresh title' -or $script:checkboxClick){throw 'A last-moment checkbox was not merged into refreshed data'}
 $cached=Get-Content -LiteralPath $script:cachePath -Raw -Encoding UTF8|ConvertFrom-Json
 if(-not $cached.data.tasks[0].done -or $cached.pendingOps.Count -ne 1){throw 'The refreshed checkbox was not durably saved'}
 'PASS: date changes during loading and last-moment checkbox edits survive responsive refresh.'
}finally{if($script:checkboxClick){$script:checkboxClick.timer.Stop();$script:checkboxClick=$null};Stop-AgendaReader;$script:window.Close()}
