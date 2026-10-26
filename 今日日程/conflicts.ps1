function Apply-AgendaSyncReview($review,[int[]]$selected){
 $check=Invoke-Backend '--check';if($check.error){throw $check.error}
 $stamp=Invoke-Backend '--stamp'
 if([string]$stamp.version -ne [string]$review.version -or -not $review.digest -or [string]$stamp.digest -ne [string]$review.digest){throw '核对期间 Excel 又有更新，本次选择尚未应用，请重新核对。'}
 $token=[guid]::NewGuid().ToString('N')
 Copy-Item -LiteralPath $script:planPath -Destination ($script:planPath+'.before-merge-'+$token+'.xlsx')
 Copy-Item -LiteralPath $script:cachePath -Destination ($script:cachePath+'.before-merge-'+$token+'.json')
 $stamp=Invoke-Backend '--stamp'
 if([string]$stamp.version -ne [string]$review.version -or [string]$stamp.digest -ne [string]$review.digest){throw '备份期间 Excel 又有更新，本地修改仍保留，请重新核对。'}
 $kept=@()
 foreach($entry in $review.entries){
  if($entry.index -notin $selected){continue}
  if(-not $entry.canApply){throw '不能恢复已删除或位置改变的任务，请保留 Excel 版本。'}
  $operation=$script:pendingOps[[int]$entry.index]
  if($operation.id -ne $entry.id -or $operation.date -ne $entry.date){throw '本地任务已变化，请重新核对。'}
  $kept+=@($operation)
 }
 $script:pendingOps=$kept;$script:baseVersion=[string]$review.version;$script:baseDigest=[string]$stamp.digest
 foreach($operation in $script:pendingOps){Set-LocalValue $operation 'version' $script:baseVersion}
 Write-AgendaCache
 $script:data=$null
 Refresh-Agenda $script:lastDate -ReloadExcel
}
function Show-AgendaSyncReview {
 try{
  $review=Invoke-Backend '--review-sync' (@{ops=@($script:pendingOps)}|ConvertTo-Json -Depth 12 -Compress)
  if($review.error){throw $review.error}
  $dialog=New-Object Windows.Window;$dialog.Title='核对 Excel 与小程序的修改';$dialog.Owner=$script:window;$dialog.WindowStartupLocation='CenterOwner';$dialog.Width=720;$dialog.Height=600;$dialog.Background='#FCFEFC';$dialog.FontFamily='Microsoft YaHei'
  $layout=New-Object Windows.Controls.DockPanel;$layout.Margin='18'
  $hint=New-Object Windows.Controls.TextBlock;$hint.Text="Excel 已被保存或修改，程序没有覆盖它。请逐项选择要采用的小程序修改；未勾选的项目保留 Excel 版本。`n目前没有足够的旧值用于自动判断冲突，确认前会备份两份数据。";$hint.TextWrapping='Wrap';$hint.Foreground='#547562';$hint.Margin='0,0,0,12';[Windows.Controls.DockPanel]::SetDock($hint,'Top');[void]$layout.Children.Add($hint)
  $actions=New-Object Windows.Controls.StackPanel;$actions.Orientation='Horizontal';$actions.HorizontalAlignment='Right';$actions.Margin='0,12,0,0';[Windows.Controls.DockPanel]::SetDock($actions,'Bottom')
  $cancel=New-Object Windows.Controls.Button;$cancel.Content='取消，保留现状';$cancel.Padding='12,7';$cancel.Margin='0,0,10,0';$cancel.IsCancel=$true
  $confirm=New-Object Windows.Controls.Button;$confirm.Content='确认选择并继续同步';$confirm.Padding='12,7';$confirm.Foreground='#43765C';$confirm.Background='#E7F2EB'
  [void]$actions.Children.Add($cancel);[void]$actions.Children.Add($confirm);[void]$layout.Children.Add($actions)
  $scroll=New-Object Windows.Controls.ScrollViewer;$scroll.VerticalScrollBarVisibility='Auto';$list=New-Object Windows.Controls.StackPanel;$choices=New-Object 'System.Collections.Generic.List[object]'
  foreach($entry in $review.entries){
   $card=New-Object Windows.Controls.StackPanel;$card.Margin='0,0,8,16'
   $title=New-Object Windows.Controls.TextBlock;$title.Text=$entry.date+' · '+$entry.title;$title.FontWeight='SemiBold';$title.TextWrapping='Wrap';$title.Foreground='#345347';[void]$card.Children.Add($title)
   foreach($side in @(@('Excel 当前内容',$entry.excel),@('小程序待同步内容',$entry.local))){$text=New-Object Windows.Controls.TextBlock;$text.Text=$side[0]+"`n"+$side[1];$text.TextWrapping='Wrap';$text.Margin='0,6,0,0';[void]$card.Children.Add($text)}
   $choice=New-Object Windows.Controls.CheckBox;$choice.Content='采用这项小程序修改';$choice.Foreground='#43765C';$choice.Margin='0,8,0,0';$choice.Tag=[int]$entry.index;$choice.IsEnabled=[bool]$entry.canApply;$choice.IsChecked=$false;[void]$card.Children.Add($choice);$choices.Add($choice);[void]$list.Children.Add($card)
  }
  $scroll.Content=$list;[void]$layout.Children.Add($scroll);$dialog.Content=$layout
  $confirm.Add_Click({$dialog.DialogResult=$true})
  if($dialog.ShowDialog() -ne $true){$script:statusLabel.Text='尚未合并 · 两边修改均保留';return}
  $selected=@($choices|Where-Object {$_.IsChecked -and $_.IsEnabled}|ForEach-Object {[int]$_.Tag})
  Apply-AgendaSyncReview $review $selected
  if($script:pendingOps.Count){Start-AgendaSync}else{$script:statusLabel.Text='已采用 Excel 版本 · 原本地修改已备份'}
 }catch{$script:statusLabel.Text='合并未完成 · 本地修改已保留';[void][Windows.MessageBox]::Show($_.Exception.Message,'暂未合并','OK','Information')}
}
