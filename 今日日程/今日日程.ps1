param([string]$PreviewDate='', [string]$PreviewPath='', [string]$BookPath='', [int]$PreviewPage=0)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase
$script:baseDir=Split-Path -Parent $MyInvocation.MyCommand.Path
$settingsPath=Join-Path (Split-Path -Parent $script:baseDir) 'qingxu.local.json'
$settings=if(Test-Path -LiteralPath $settingsPath){Get-Content -LiteralPath $settingsPath -Raw -Encoding UTF8|ConvertFrom-Json}else{[PSCustomObject]@{}}
$script:pythonPath=if($env:AGENDA_PYTHON){$env:AGENDA_PYTHON}else{(Get-Command python.exe -ErrorAction Stop).Source}
$script:readerPath=Join-Path $script:baseDir 'read_plan.py'
$script:planPath=Join-Path (Split-Path -Parent $script:baseDir) 'outputs\study-plan\计划.xlsx'
if($BookPath){$script:planPath=$BookPath}
elseif($env:AGENDA_BOOK){$script:planPath=$env:AGENDA_BOOK}
elseif($settings.bookPath){$script:planPath=[string]$settings.bookPath}
$script:cachePath=$script:planPath+'.agenda.json';$script:pendingOps=@();$script:data=$null;$script:busy=$false
. (Join-Path $script:baseDir 'controls.ps1')
$script:instanceMutex=$null
$script:instanceRestore=$null
if(-not $PreviewPath){
 $restoreName=if($settings.restoreEvent){[string]$settings.restoreEvent}else{'Local\QingxuRestore'}
 $mutexName=if($settings.instanceMutex){[string]$settings.instanceMutex}else{'Local\QingxuAgenda'}
 $script:instanceRestore=New-Object Threading.EventWaitHandle($false,[Threading.EventResetMode]::AutoReset,$restoreName)
 $script:instanceMutex=New-Object Threading.Mutex($false,$mutexName)
 try{$owns=$script:instanceMutex.WaitOne(0)}catch [Threading.AbandonedMutexException] {$owns=$true}
 if(-not $owns){[void]$script:instanceRestore.Set();$script:instanceRestore.Dispose();$script:instanceMutex.Dispose();exit}
}
$script:lastStamp=''
$script:baseVersion=''
$script:baseDigest=''
$script:lastDate=''
$script:anchorDate=''
# 跟随今天：用户没有主动翻到别的日期时，跨过午夜要自动锚定到当天，否则会一直停在启动那天
$script:followToday=$true
. (Join-Path $script:baseDir 'desktop.ps1')
Set-AgendaApplicationIdentity
[xml]$xaml=[IO.File]::ReadAllText((Join-Path $script:baseDir 'view.xaml'))
$script:window=[Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
$iconPath=Join-Path $script:baseDir 'icon.ico';if(Test-Path -LiteralPath $iconPath){$script:window.Icon=[Windows.Media.Imaging.BitmapFrame]::Create([Uri]::new($iconPath))}
$script:content=$script:window.FindName('ContentPanel')
$script:dateLabel=$script:window.FindName('DateLabel')
$script:statusLabel=$script:window.FindName('StatusLabel')
function Add-Text([string]$text,[int]$size=13,[string]$color='#344B61'){
 $t=New-Object Windows.Controls.TextBlock;$t.Text=$text;$t.FontSize=$size;$t.Foreground=$color;$t.TextWrapping='Wrap';$t.Margin='0,3,0,5';[void]$script:content.Children.Add($t)
}
function Add-Section([string]$text){$t=New-Object Windows.Controls.TextBlock;$t.Text=$text;$t.FontWeight='SemiBold';$t.FontSize=12;$t.Foreground='#8AA394';$t.Margin='2,10,0,7';[void]$script:content.Children.Add($t)}
$script:page=$PreviewPage
. (Join-Path $script:baseDir 'pages.ps1')
$script:window.FindName('CloseButton').Add_Click({$script:window.Close()})
$script:window.FindName('PinButton').Add_Click({$script:window.Topmost=[bool]$script:window.FindName('PinButton').IsChecked})
$script:window.FindName('ScheduleButton').Add_Click({$script:window.FindName('TimeSlotPopup').IsOpen=$true})
$script:window.FindName('CloseSchedulePopup').Add_Click({$script:window.FindName('TimeSlotPopup').IsOpen=$false})
$script:window.FindName('OpenPlanButton').Add_Click({Start-AgendaSync})
$script:window.FindName('RefreshButton').Add_Click({if(-not $script:syncProcess){Refresh-Agenda $script:anchorDate -ReloadExcel}})
$script:window.Add_Activated({Refresh-AgendaIfChanged})
$script:window.FindName('TodayTab').Add_Click({Set-Page 0})
$script:window.FindName('TomorrowTab').Add_Click({Set-Page 1})
$script:window.FindName('AfterTomorrowTab').Add_Click({Set-Page 2})
$script:window.FindName('WeekTab').Add_Click({Set-Page 3})
$script:window.FindName('PrevButton').Add_Click({Move-View -Offset -1})
$script:window.FindName('NextButton').Add_Click({Move-View -Offset 1})
$script:window.FindName('DatePicker').Add_SelectedDateChanged({param($sender,$e)if($null -ne $sender.SelectedDate){$selected=([datetime]$sender.SelectedDate).ToString('yyyy-MM-dd');$script:followToday=($selected -eq [DateTime]::UtcNow.AddHours(8).ToString('yyyy-MM-dd'));if($selected -ne $script:anchorDate){$script:anchorDate=$selected;$script:page=0;Refresh-Agenda $selected}}})
$script:window.FindName('DatePickerButton').Add_Click({$picker=$script:window.FindName('DatePicker');$calendar=$script:window.FindName('DateCalendar');$selected=[datetime]::ParseExact($script:anchorDate,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture);$script:syncCalendar=$true;try{$calendar.DisplayDate=$selected;$calendar.SelectedDate=$selected}finally{$script:syncCalendar=$false};$script:window.FindName('DatePopup').IsOpen=$true})
$script:window.FindName('DateCalendar').Add_SelectedDatesChanged({param($sender,$e)if($script:syncCalendar -or $null -eq $sender.SelectedDate){return};$script:window.FindName('DatePicker').SelectedDate=$sender.SelectedDate;$script:window.FindName('DatePopup').IsOpen=$false})
$script:window.FindName('NewButton').Add_Click({$date=if($script:page -lt 3){$script:data.days[$script:page].date}else{$script:lastDate};Edit-Item ([PSCustomObject]@{id='new';title='';done=$false;highlight=$false;start='';planned='';actual='';subject='';agendaDate=$date})})
$script:window.FindName('Header').Add_MouseLeftButtonDown({param($sender,$e)if($e.GetPosition($sender).X -lt ($sender.ActualWidth-60)){$script:window.DragMove();$e.Handled=$true}})
$script:window.Add_Loaded({
 Refresh-Agenda
 if($PreviewPath){
  $script:window.UpdateLayout();$bmp=New-Object Windows.Media.Imaging.RenderTargetBitmap ([int]$script:window.ActualWidth),([int]$script:window.ActualHeight),96,96,([Windows.Media.PixelFormats]::Pbgra32);$bmp.Render($script:window);$encoder=New-Object Windows.Media.Imaging.PngBitmapEncoder;$encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bmp));$stream=[IO.File]::Create($PreviewPath);$encoder.Save($stream);$stream.Dispose();$script:window.Close()
 }else{$script:window.UpdateLayout();$area=[System.Windows.SystemParameters]::WorkArea;$script:window.Left=$area.Right-$script:window.ActualWidth-22;$script:window.Top=$area.Bottom-$script:window.ActualHeight-22}
})
$script:window.Add_Closing({param($sender,$e)Complete-AgendaCheckboxClick;if($script:syncProcess){$e.Cancel=$true;$script:statusLabel.Text='Excel 同步尚未结束，请完成后再关闭'}})
$script:window.Add_Closed({Stop-AgendaReader;if($script:instanceMutex){try{$script:instanceMutex.ReleaseMutex()}catch{};$script:instanceMutex.Dispose()};if($script:instanceRestore){$script:instanceRestore.Dispose()};[Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvokeShutdown([Windows.Threading.DispatcherPriority]::Background)})
$script:window.Add_ContentRendered({if(-not $PreviewPath){[IO.File]::WriteAllText((Join-Path $script:baseDir '状态.json'),(@{visible=$script:window.IsVisible;loaded=$script:window.IsLoaded;date=$script:dateLabel.Text;source=$script:planPath;sync='manual'}|ConvertTo-Json),[Text.Encoding]::UTF8)}})
if(-not $PreviewPath){Initialize-AgendaDesktop $script:instanceRestore;Initialize-AgendaExcelWatcher}
$script:window.Show()
[Windows.Threading.Dispatcher]::Run()
