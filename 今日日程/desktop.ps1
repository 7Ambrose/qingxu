function Set-AgendaApplicationIdentity {
 if(-not ('QingxuTaskbar' -as [type])){Add-Type -Path (Join-Path $script:baseDir 'taskbar.cs')}
 [QingxuTaskbar]::Initialize()
}
function Show-AgendaWindow {
 $script:window.ShowInTaskbar=$true
 $script:window.Show()
 $script:window.WindowState=[Windows.WindowState]::Normal
 [void]$script:window.Activate()
}
function Request-AgendaExit {
 $script:exitRequested=$true
 $script:window.Close()
 if(-not $script:desktopClosed){$script:exitRequested=$false}
}
function Initialize-AgendaExcelWatcher {
 $script:excelWatcher=New-Object QingxuWorkbookWatcher $script:planPath
 $script:excelWatchTimer=New-Object Windows.Threading.DispatcherTimer
 $script:excelWatchTimer.Interval=[TimeSpan]::FromSeconds(1)
 $script:excelWatchTimer.Add_Tick({
  if($script:busy -or $script:refreshing -or $script:syncProcess){return}
  # Excel 没动但日期翻篇了也要检查，否则常驻窗口会一直停在启动当天
  $rolledOver=$script:followToday -and $script:anchorDate -ne [DateTime]::UtcNow.AddHours(8).ToString('yyyy-MM-dd')
  if($script:excelWatcher.ConsumeChange() -or $rolledOver){Refresh-AgendaIfChanged}
 })
 $script:excelWatchTimer.Start()
 $script:window.Add_Closed({$script:excelWatchTimer.Stop();$script:excelWatcher.Dispose()})
}

function Initialize-AgendaDesktop($restoreEvent){
 Set-AgendaApplicationIdentity
 Add-Type -AssemblyName System.Windows.Forms,System.Drawing
 $script:exitRequested=$false
 $script:desktopClosed=$false
 $script:trayIcon=New-Object Windows.Forms.NotifyIcon
 $script:trayIcon.Text='青序 · 点击恢复日程'
 $iconPath=Join-Path $script:baseDir 'icon.ico'
 $script:trayImage=if(Test-Path -LiteralPath $iconPath){New-Object Drawing.Icon $iconPath}else{[Drawing.SystemIcons]::Application.Clone()}
 $script:trayIcon.Icon=$script:trayImage
 if(Test-Path -LiteralPath $iconPath){$script:window.Icon=[Windows.Media.Imaging.BitmapFrame]::Create([Uri]::new($iconPath))}
 $script:window.Add_SourceInitialized({$handle=(New-Object Windows.Interop.WindowInteropHelper $script:window).Handle;[QingxuTaskbar]::SetWindowIcon($handle,$script:trayImage.Handle)})
 $script:trayMenu=New-Object Windows.Forms.ContextMenuStrip
 $show=$script:trayMenu.Items.Add('显示青序');$show.Add_Click({Show-AgendaWindow})
 $exit=$script:trayMenu.Items.Add('退出青序');$exit.Add_Click({Request-AgendaExit})
 $script:trayIcon.ContextMenuStrip=$script:trayMenu
 $script:trayIcon.Add_MouseClick({param($sender,$event)if($event.Button -eq [Windows.Forms.MouseButtons]::Left){Show-AgendaWindow}})
 $script:trayIcon.Visible=$true
 $script:window.FindName('CloseButton').ToolTip='收起到系统托盘；右键托盘图标可退出'
 $script:window.Add_Closing({param($sender,$event)if(-not $event.Cancel -and -not $script:exitRequested){$event.Cancel=$true;$script:window.Hide();$script:window.ShowInTaskbar=$false}})
 $script:window.Add_Closed({
  $script:desktopClosed=$true
  if($script:restoreTimer){$script:restoreTimer.Stop()}
  $script:trayIcon.Visible=$false;$script:trayIcon.Dispose();$script:trayMenu.Dispose();$script:trayImage.Dispose()
 })
 if($restoreEvent){
  $script:restoreEvent=$restoreEvent
  $script:restoreTimer=New-Object Windows.Threading.DispatcherTimer
  $script:restoreTimer.Interval=[TimeSpan]::FromMilliseconds(250)
  $script:restoreTimer.Add_Tick({if($script:restoreEvent.WaitOne(0)){Show-AgendaWindow}})
  $script:restoreTimer.Start()
 }
}
