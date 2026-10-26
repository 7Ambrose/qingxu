using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json.Nodes;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Forms = System.Windows.Forms;

namespace Qingxu;

public sealed class MainWindow : Window
{
    public AgendaStore Store { get; }
    private readonly StackPanel content = new();
    private readonly StackPanel tabs = new() { Orientation = Orientation.Horizontal };
    private readonly ScrollViewer scroll = new() { VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
    private readonly TextBlock status = new() { FontSize = 10, TextTrimming = TextTrimming.CharacterEllipsis };
    private readonly TextBlock progressText = new() { FontSize = 11 };
    private readonly ProgressBar progressBar = new() { Height = 4, Minimum = 0, Maximum = 1, BorderThickness = new Thickness(0) };
    private AgendaDay? currentDay;
    private readonly TextBlock dateLabel = new() { FontSize = 11 };
    private readonly Forms.NotifyIcon tray;
    private readonly Forms.ContextMenuStrip trayMenu = new();
    private readonly System.Drawing.Icon trayImage;
    private readonly RegisteredWaitHandle? restoreWait;
    private readonly System.Windows.Controls.Button syncButton;
    private readonly Popup schedulePopup = new() { StaysOpen = true, AllowsTransparency = true };
    private DateOnly anchor;
    private DateOnly dropDate;
    private int page;
    private bool exitRequested;
    private bool syncing;
    public bool IsSyncing => syncing;
    private bool followToday = true;
    private readonly DispatcherTimer dayTimer;
    private readonly DispatcherTimer watchTimer;
    private readonly List<(string Id, Action<AgendaItem> Update)> visuals = new();
    private (AgendaItem Item, bool Done, DispatcherTimer Timer)? pendingCheck;

    public MainWindow(string path, EventWaitHandle? restore = null, int startPage = 0)
    {
        Store = new AgendaStore(path);
        anchor = AgendaRules.Today;
        page = Math.Max(0, Math.Min(3, startPage));
        if (anchor < AgendaRules.FirstDate || anchor > AgendaRules.LastDate) anchor = AgendaRules.FirstDate;
        dropDate = anchor;
        Title = "青序"; Width = 398; Height = 700; MinWidth = 365; MinHeight = 320;
        WindowStyle = WindowStyle.None; ResizeMode = ResizeMode.CanResizeWithGrip; AllowsTransparency = true;
        Background = Brushes.Transparent; FontFamily = new FontFamily("Microsoft YaHei"); FontSize = 13;
        Resources.MergedDictionaries.Add(new ResourceDictionary { Source = new Uri("/青序;component/Theme.xaml", UriKind.Relative) });
        using var iconStream = typeof(MainWindow).Assembly.GetManifestResourceStream("Qingxu.icon.ico") ?? throw new IOException("青序图标资源缺失。");
        Icon = BitmapFrame.Create(iconStream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad); Icon.Freeze();
        iconStream.Position = 0; trayImage = new System.Drawing.Icon(iconStream);
        tray = new Forms.NotifyIcon { Text = "青序 · 点击恢复日程", Icon = trayImage, ContextMenuStrip = trayMenu, Visible = true };
        trayMenu.Items.Add("显示青序", null, (_, _) => Restore());
        trayMenu.Items.Add("重新读取 Excel", null, (_, _) => Run(() => { Store.Reload(); Render(); }));
        trayMenu.Items.Add("退出青序", null, (_, _) => Exit());
        tray.MouseClick += (_, eventArgs) => { if (eventArgs.Button == Forms.MouseButtons.Left) Restore(); };
        if (restore is not null) restoreWait = ThreadPool.RegisterWaitForSingleObject(restore, (_, _) => Dispatcher.BeginInvoke(Restore), null, Timeout.Infinite, false);
        var root = new Grid();
        for (var row = 0; row < 5; row++) root.RowDefinitions.Add(new RowDefinition { Height = row == 3 ? new GridLength(1, GridUnitType.Star) : GridLength.Auto });
        var header = new DockPanel { Cursor = Cursors.SizeAll, Margin = new Thickness(0, 0, 0, 12) };
        var actions = new StackPanel { Orientation = Orientation.Horizontal };
        var pin = Button("📌", (_, _) => { Topmost = !Topmost; }); pin.ToolTip = "切换置顶";
        actions.Children.Add(pin);
        var close = Button("×", (_, _) => Close()); close.ToolTip = "收起到系统托盘，右键托盘可退出"; actions.Children.Add(close);
        DockPanel.SetDock(actions, Dock.Right); header.Children.Add(actions);
        var branding = new StackPanel(); branding.Children.Add(new TextBlock { Text = "青序", FontSize = 23, FontWeight = FontWeights.SemiBold, Foreground = Brush("#345347") });
        dateLabel.Foreground = Brush("#8AA394"); branding.Children.Add(dateLabel); header.Children.Add(branding);
        branding.MouseLeftButtonDown += (_, eventArgs) => { if (eventArgs.LeftButton == MouseButtonState.Pressed) DragMove(); };
        root.Children.Add(header);
        Grid.SetRow(tabs, 1); tabs.Margin = new Thickness(0, 0, 0, 6); root.Children.Add(tabs);
        var progressRow = new StackPanel { Margin = new Thickness(2, 0, 2, 10) };
        progressText.Foreground = Brush("#8AA394"); progressRow.Children.Add(progressText);
        progressBar.Foreground = Brush("#7FB794"); progressBar.Background = Brush("#E4EEE7"); progressBar.Margin = new Thickness(0, 5, 0, 0);
        progressRow.Children.Add(progressBar);
        Grid.SetRow(progressRow, 2); root.Children.Add(progressRow);
        scroll.Content = content; Grid.SetRow(scroll, 3); root.Children.Add(scroll);
        var footer = new StackPanel { Margin = new Thickness(0, 10, 0, 0) };
        var quick = new TextBox { FontSize = 12, Padding = new Thickness(8, 5, 8, 5), BorderBrush = Brush("#DCE9E1"), Background = Brush("#FBFDFB"), Foreground = Brush("#345347"),
            ToolTip = "输入任务名称，回车添加到当前日期；不需要先点新建" };
        quick.KeyDown += (_, eventArgs) => { if (eventArgs.Key != Key.Enter) return; eventArgs.Handled = true; QuickAdd(quick.Text); quick.Clear(); };
        var quickRow = new DockPanel();
        var quickHint = new TextBlock { Text = "＋", Foreground = Brush("#8AA394"), FontSize = 14, Margin = new Thickness(2, 0, 6, 0), VerticalAlignment = VerticalAlignment.Center, ToolTip = "快速添加" };
        DockPanel.SetDock(quickHint, Dock.Left); quickRow.Children.Add(quickHint); quickRow.Children.Add(quick);
        footer.Children.Add(quickRow);
        status.Margin = new Thickness(0, 7, 0, 0); status.Foreground = Brush("#8AA394"); footer.Children.Add(status);
        var footerActions = new WrapPanel { HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 4, 0, 0) };
        footerActions.Children.Add(Button("拖拽排期", (_, _) => OpenScheduling()));
        footerActions.Children.Add(Button("＋新建", (_, _) => NewTask()));
        syncButton = Button("同步并打开Excel", async (_, _) => await Synchronize()); syncButton.FontSize = 10; footerActions.Children.Add(syncButton);
        footer.Children.Add(footerActions); Grid.SetRow(footer, 4); root.Children.Add(footer);
        Content = new Border { Background = Brush("#FCFEFC"), BorderBrush = Brush("#DCE9E1"), BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(18), Padding = new Thickness(20, 17, 20, 15), Child = root };
        Loaded += (_, _) => { Render(); var area = SystemParameters.WorkArea; Left = area.Right - ActualWidth - 22; Top = area.Bottom - ActualHeight - 22; };
        // 常驻窗口跨过午夜后锚点不会自己前进，必须定时把视图挪到当天
        dayTimer = new DispatcherTimer { Interval = TimeSpan.FromMinutes(1) };
        dayTimer.Tick += (_, _) => { if (followToday && anchor != AgendaRules.Today) Run(() => { Store.Reload(); ChangeAnchor(AgendaRules.Today, true); }); };
        dayTimer.Start();
        // Excel 被别处改动时自动重新读取，等同旧版的 1 秒文件监控
        watchTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(2) };
        watchTimer.Tick += (_, _) =>
        {
            if (syncing) return;
            try { if (Store.Book.Version != XlsxBook.CurrentVersion(Store.BookPath)) { Store.Reload(); Render(); } } catch { }
        };
        watchTimer.Start();
        PreviewKeyDown += (_, eventArgs) =>
        {
            if (Keyboard.FocusedElement is TextBox && eventArgs.Key != Key.Escape) return;
            switch (eventArgs.Key)
            {
                case Key.Left: Navigate(-1); eventArgs.Handled = true; break;
                case Key.Right: Navigate(1); eventArgs.Handled = true; break;
                case Key.T: ChangeAnchor(AgendaRules.Today); eventArgs.Handled = true; break;
                case Key.N: NewTask(); eventArgs.Handled = true; break;
                case Key.R: Run(() => { Store.Reload(); Render(); }); status.Text = "已重新读取 Excel"; eventArgs.Handled = true; break;
                case Key.D1: SetPage(0); eventArgs.Handled = true; break;
                case Key.D2: SetPage(1); eventArgs.Handled = true; break;
                case Key.D3: SetPage(2); eventArgs.Handled = true; break;
                case Key.D4: SetPage(3); eventArgs.Handled = true; break;
                case Key.Escape: schedulePopup.IsOpen = false; eventArgs.Handled = true; break;
            }
        };
        Closing += OnClosing;
        Closed += (_, _) => { dayTimer.Stop(); watchTimer.Stop(); schedulePopup.IsOpen = false; restoreWait?.Unregister(null); tray.Visible = false; tray.Dispose(); trayImage.Dispose(); trayMenu.Dispose(); };
    }
    public static SolidColorBrush Brush(string value) => new((Color)ColorConverter.ConvertFromString(value));
    public static System.Windows.Controls.Button Button(string title, RoutedEventHandler handler)
    {
        var button = new System.Windows.Controls.Button { Content = title, Foreground = Brush("#658477"), Background = Brushes.Transparent,
            BorderThickness = new Thickness(0), Padding = new Thickness(7, 5, 7, 5), Cursor = Cursors.Hand, FontSize = 12 };
        button.Click += handler; return button;
    }
    private void Run(Action action)
    {
        try { action(); } catch (Exception error) { MessageBox.Show(this, error.Message, "青序", MessageBoxButton.OK, MessageBoxImage.Information); }
    }
    public void Restore() { Show(); ShowInTaskbar = true; WindowState = WindowState.Normal; Activate(); }
    private void OnClosing(object? sender, CancelEventArgs eventArgs)
    {
        if (syncing) { eventArgs.Cancel = true; status.Text = "后台同步尚未结束，请稍候"; return; }
        if (!exitRequested) { eventArgs.Cancel = true; schedulePopup.IsOpen = false; Hide(); ShowInTaskbar = false; }
    }
    public void Exit()
    {
        if (syncing) { status.Text = "后台同步尚未结束，请稍候"; return; }
        exitRequested = true; Close(); Application.Current.Shutdown();
    }
    private DateOnly SelectedDate => anchor.AddDays(Math.Min(page, 2));
    private void SetPage(int value) { page = value; schedulePopup.IsOpen = false; Render(); scroll.ScrollToTop(); }
    private void Navigate(int direction)
    {
        if (page == 3) { ChangeAnchor(anchor.AddDays(direction * 7)); page = 3; Render(); return; }
        var next = page + direction;
        if (next >= 0 && next < 3) { if (anchor.AddDays(next) <= AgendaRules.LastDate) SetPage(next); }
        else { ChangeAnchor(anchor.AddDays(direction * 3)); page = direction > 0 ? 0 : Math.Min(2, AgendaRules.LastDate.DayNumber - anchor.DayNumber); Render(); }
    }
    private void ChangeAnchor(DateOnly date, bool follow = false)
    {
        if (date < AgendaRules.FirstDate || date > AgendaRules.LastDate) return;
        anchor = date; page = 0; dropDate = date; schedulePopup.IsOpen = false;
        // 用户主动翻到别的日期后就不再跟随今天，否则看上周数据会被拽回来
        followToday = follow || date == AgendaRules.Today;
        Render();
    }
    private void PickDate()
    {
        var calendar = new System.Windows.Controls.Calendar { SelectedDate = anchor.ToDateTime(TimeOnly.MinValue), DisplayDate = anchor.ToDateTime(TimeOnly.MinValue), DisplayDateStart = AgendaRules.FirstDate.ToDateTime(TimeOnly.MinValue), DisplayDateEnd = AgendaRules.LastDate.ToDateTime(TimeOnly.MinValue) };
        var popup = new Popup { PlacementTarget = tabs, Placement = PlacementMode.Bottom, StaysOpen = false, AllowsTransparency = true, Child = new Border { Background = Brush("#FCFEFC"), BorderBrush = Brush("#DCE9E1"), BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(10), Padding = new Thickness(6), Child = calendar } };
        calendar.SelectedDatesChanged += (_, _) => { if (calendar.SelectedDate is DateTime selected) { popup.IsOpen = false; ChangeAnchor(DateOnly.FromDateTime(selected)); } };
        popup.IsOpen = true;
    }
    public void Render()
    {
        if (syncing) return;
        var offset = scroll.VerticalOffset;
        tabs.Children.Clear(); tabs.Children.Add(Button("‹", (_, _) => Navigate(-1)));
        for (var index = 0; index < 3; index++)
        {
            var captured = index; var date = anchor.AddDays(index);
            var button = Button(date == AgendaRules.Today ? "今天" : date.ToString("MM/dd"), (_, _) => SetPage(captured));
            button.IsEnabled = date <= AgendaRules.LastDate;
            if (page == index) button.Background = Brush("#E6F2EB"); tabs.Children.Add(button);
        }
        var week = Button("本周", (_, _) => SetPage(3)); if (page == 3) week.Background = Brush("#E6F2EB"); tabs.Children.Add(week);
        tabs.Children.Add(Button("选日期⌄", (_, _) => PickDate())); tabs.Children.Add(Button("›", (_, _) => Navigate(1)));
        var day = Store.Day(page == 3 ? anchor : SelectedDate);
        currentDay = day;
        dateLabel.Text = page == 3 ? day.WeekName + " · " + day.WeekStart : SelectedDate.ToString("M月d日 dddd", CultureInfo.GetCultureInfo("zh-CN"));
        content.Children.Clear(); visuals.Clear();
        if (page == 3)
        {
            Section("本周常规任务");
            foreach (var item in day.Weekly) Card(item, true);
            Section("每周课程");
            foreach (var item in day.WeekCourses)
                content.Children.Add(new TextBlock { Text = AgendaRules.ParseDate(item.Date).ToString("ddd", CultureInfo.GetCultureInfo("zh-CN")) + "  " + item.Time + "\n" + item.Title, TextWrapping = TextWrapping.Wrap, Foreground = Brush("#658477"), Margin = new Thickness(0, 5, 0, 6) });
        }
        else
        {
            Section("日程");
            foreach (var item in day.Schedule)
            {
                if (item.Class || !item.Editable)
                {
                    var course = new DockPanel { Margin = new Thickness(0, 5, 0, 4) };
                    if (item.Class) { var edit = Button("···", (_, _) => Edit(item)); DockPanel.SetDock(edit, Dock.Right); course.Children.Add(edit); }
                    course.Children.Add(new TextBlock { Text = item.Time + "\n" + item.Title, TextWrapping = TextWrapping.Wrap, Foreground = Brush("#345347") }); content.Children.Add(course);
                    if (item.Class && item.TaskTitle.Length > 0) Card(item, false);
                }
                else Card(item, false);
            }
            Section("待办"); foreach (var item in day.Tasks) Card(item, false);
        }
        status.Text = (Store.ImportedLegacy ? "已读取旧版待同步内容 · " : "") + (Store.State.Operations.Count > 0 ? "本地已保存 · 待同步 " + Store.State.Operations.Count + " 项" : "已与 Excel 同步");
        status.ToolTip = status.Text; scroll.ScrollToVerticalOffset(offset); UpdateProgress();
    }
    private void UpdateProgress()
    {
        if (currentDay is null) { progressText.Text = ""; progressBar.Value = 0; return; }
        var items = page == 3 ? currentDay.Weekly.ToList() : currentDay.Tasks.Concat(currentDay.Schedule.Where(item => item.Editable && !item.Class)).ToList();
        var done = items.Count(item => item.Done);
        var label = page == 3 ? "本周" : "今日";
        progressText.Text = items.Count == 0 ? label + "暂无待办" : $"{label} {done}/{items.Count} 已完成 · 还剩 {items.Count - done} 项";
        progressBar.Value = items.Count == 0 ? (page == 3 ? 0 : 1) : (double)done / items.Count;
    }
    private void QuickAdd(string text)
    {
        var title = text.Trim();
        if (title.Length == 0) return;
        if (page == 3) { status.Text = "周清单沿用 Excel 模板，请在日程页添加"; return; }
        Run(() =>
        {
            var day = Store.Day(SelectedDate);
            var item = new AgendaItem { Id = "local:" + Guid.NewGuid().ToString("N"), Date = day.Date,
                Order = Math.Max(0, day.Tasks.Select(task => task.Order).DefaultIfEmpty(0).Max()) + 1 };
            Store.Save(item, new JsonObject { ["title"] = title, ["done"] = false, ["incomplete"] = false });
            Render(); status.Text = "已添加 · 待同步";
        });
    }
    private void Section(string title) => content.Children.Add(new TextBlock { Text = title, Foreground = Brush("#8AA394"), FontWeight = FontWeights.SemiBold, FontSize = 12, Margin = new Thickness(0, 10, 0, 7) });
    private void Card(AgendaItem item, bool weekly)
    {
        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        grid.ColumnDefinitions.Add(new ColumnDefinition());
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        var handle = new TextBlock { Text = "⠿", Foreground = Brush("#9BB3A5"), Cursor = Cursors.SizeAll, Margin = new Thickness(0, 2, 8, 0), ToolTip = "拖动排序，或拖到排期时间点" };
        if (item.Class || item.Id.StartsWith("semester:")) handle.Visibility = Visibility.Collapsed;
        Point? start = null;
        handle.MouseLeftButtonDown += (_, eventArgs) => start = eventArgs.GetPosition(handle);
        handle.MouseMove += (_, eventArgs) =>
        {
            if (eventArgs.LeftButton != MouseButtonState.Pressed || start is not Point point) return;
            var position = eventArgs.GetPosition(handle);
            if (Math.Abs(position.X - point.X) < SystemParameters.MinimumHorizontalDragDistance && Math.Abs(position.Y - point.Y) < SystemParameters.MinimumVerticalDragDistance) return;
            start = null; DragDrop.DoDragDrop(handle, new DataObject("QingxuTask", item), DragDropEffects.Move);
        };
        grid.Children.Add(handle);
        var text = new StackPanel();
        var label = new TextBlock { Text = item.DisplayTitle, TextWrapping = TextWrapping.Wrap, FontSize = 14 };
        var tick = new System.Windows.Shapes.Path { Data = Geometry.Parse("M1,5 L4,8 L9,1"), Stroke = Brush("#547D65"), StrokeThickness = 1.6, Visibility = Visibility.Collapsed };
        var slash = new System.Windows.Shapes.Path { Data = Geometry.Parse("M1,9 L9,1"), Stroke = Brush("#969C98"), StrokeThickness = 1.5, Visibility = Visibility.Collapsed };
        var box = new Border { Width = 15, Height = 15, Margin = new Thickness(0, 2, 9, 0), VerticalAlignment = VerticalAlignment.Top,
            BorderThickness = new Thickness(1), BorderBrush = Brush("#91A298"), Background = Brushes.White,
            Child = new Grid { Children = { tick, slash } } };
        var head = new Grid { Background = Brushes.Transparent, Cursor = Cursors.Hand, ToolTip = "单击完成／取消完成；双击标记未完成（灰色斜线）" };
        head.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        head.ColumnDefinitions.Add(new ColumnDefinition());
        head.Children.Add(box);
        Grid.SetColumn(label, 1); head.Children.Add(label);
        head.ContextMenu = StateMenu(item);
        head.MouseLeftButtonDown += (_, eventArgs) =>
        {
            if (eventArgs.ClickCount == 2) { HandleCheck(item, true); eventArgs.Handled = true; return; }
            HandleCheck(item, false); eventArgs.Handled = true;
        };
        text.Children.Add(head);
        if (item.Detail.Length > 0 || item.Class) text.Children.Add(new TextBlock { Text = item.Class ? "课内安排" : item.Detail, FontSize = 11, Foreground = Brush("#8B9B92"), Margin = new Thickness(18, 4, 0, 0), TextWrapping = TextWrapping.Wrap });
        Grid.SetColumn(text, 1); grid.Children.Add(text);
        var actions = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Top };
        var star = new ToggleButton { FontSize = 16, Padding = new Thickness(3, 0, 3, 0), Margin = new Thickness(3, 0, 3, 0), Background = Brushes.Transparent,
            BorderThickness = new Thickness(0), Cursor = Cursors.Hand, ToolTip = "强调", VerticalAlignment = VerticalAlignment.Top };
        actions.Children.Add(star);
        actions.Children.Add(Button("···", (_, _) => Edit(item)));
        var delete = Button(item.Weekly && !weekly ? "撤" : "删", (_, _) => Run(() =>
        {
            var message = item.Weekly && !weekly ? "只取消这一天的安排，任务仍保留在本周清单中。" : item.Class ? "只删除课内任务，保留课程本身。" : "确定删除这个任务吗？";
            if (MessageBox.Show(this, message, "青序", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;
            Store.Remove(item, weekly); Render();
        })); delete.ToolTip = item.Weekly && !weekly ? "取消安排，保留本周任务" : "删除任务"; actions.Children.Add(delete);
        Grid.SetColumn(actions, 2); grid.Children.Add(actions);
        var border = new Border { CornerRadius = new CornerRadius(8), BorderBrush = Brush("#E3EDE7"), BorderThickness = new Thickness(0, 0, 0, 1),
            Margin = new Thickness(0, 0, 0, 4), Padding = new Thickness(9, 6, 6, 6), Child = grid, AllowDrop = true };
        void Paint(AgendaItem current)
        {
            border.Background = Brush(current.Incomplete ? "#F2F3F2" : current.Done ? "#F0F8F3" : current.Highlight ? "#FFF7DF" : "#FFFFFF");
            label.Foreground = Brush(current.Incomplete ? "#969C98" : current.Done ? "#8AA197" : "#2D463C");
            label.TextDecorations = current.Done ? TextDecorations.Strikethrough : null;
            label.FontWeight = current.Highlight ? FontWeights.Bold : FontWeights.Normal;
            box.Background = current.Incomplete ? Brush("#ECEEEC") : current.Done ? Brush("#EDF5EF") : Brushes.White;
            box.BorderBrush = Brush(current.Incomplete ? "#A7ADA9" : "#91A298");
            tick.Visibility = current.Done ? Visibility.Visible : Visibility.Collapsed;
            slash.Visibility = current.Incomplete ? Visibility.Visible : Visibility.Collapsed;
            star.IsChecked = current.Highlight;
            star.Content = current.Highlight ? "★" : "☆";
            star.Foreground = Brush(current.Highlight ? "#B68B31" : "#A7BBB0");
        }
        star.Click += (_, _) =>
        {
            // 点击时 IsChecked 已是新值；保存成功后只重绘这张卡片，不重建整个列表
            var value = star.IsChecked == true;
            Run(() => { Store.Save(item, new JsonObject { ["highlight"] = value }); item.Highlight = value; Paint(item); });
        };
        Paint(item);
        visuals.Add((item.Id, Paint));
        border.DragOver += (_, eventArgs) => { eventArgs.Effects = eventArgs.Data.GetDataPresent("QingxuTask") ? DragDropEffects.Move : DragDropEffects.None; eventArgs.Handled = true; };
        border.Drop += (_, eventArgs) => Run(() =>
        {
            if (eventArgs.Data.GetData("QingxuTask") is not AgendaItem source || source.Id == item.Id) return;
            if (weekly != source.Weekly && weekly || !weekly && source.Date != item.Date || !weekly && source.Start.Length > 0 || !weekly && item.Start.Length > 0) throw new InvalidOperationException("有时间的任务请拖到排期时间点；排序只在同一天的待办或同周任务间进行。");
            var day = Store.Day(weekly ? anchor : SelectedDate);
            var list = weekly ? day.Weekly.ToList() : day.Tasks.ToList();
            var moving = list.FirstOrDefault(candidate => candidate.Id == source.Id);
            if (moving is null) throw new InvalidOperationException("请在同一个列表内拖动排序。");
            list.Remove(moving); var position = list.FindIndex(candidate => candidate.Id == item.Id);
            if (position < 0) return;
            if (eventArgs.GetPosition(border).Y > border.ActualHeight / 2) position++;
            list.Insert(position, moving);
            for (var index = 0; index < list.Count; index++) Store.Save(list[index], new JsonObject { ["order"] = index + 1 });
            eventArgs.Handled = true; Render();
        });
        content.Children.Add(border);
    }
    /// <summary>点击复选框的完整链路（不等待双击间隔），供快捷键与自动化测试使用。</summary>
    public void ToggleItem(AgendaItem item) { HandleCheck(item, false); CommitPendingCheck(); }
    private void HandleCheck(AgendaItem item, bool doubleClick)    {
        if (doubleClick)
        {
            // 双击：丢弃单击排队的那一次勾选，直接标记未完成
            if (pendingCheck is { } other && other.Item.Id != item.Id) CommitPendingCheck();
            if (pendingCheck is { } queued) { queued.Timer.Stop(); pendingCheck = null; }
            Run(() => { Store.Save(item, new JsonObject { ["incomplete"] = true }); item.Incomplete = true; item.Done = false; PaintItem(item); });
            return;
        }
        CommitPendingCheck();
        var done = !item.Done;
        var preview = item.Copy(); preview.Done = done; preview.Incomplete = false;
        PaintItem(preview);
        var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(Math.Max(200, Forms.SystemInformation.DoubleClickTime)) };
        pendingCheck = (item, done, timer);
        timer.Tick += (sender, _) =>
        {
            if (pendingCheck is { } current && ReferenceEquals(current.Timer, sender)) CommitPendingCheck();
            else (sender as DispatcherTimer)?.Stop();
        };
        timer.Start();
    }
    private void CommitPendingCheck()
    {
        if (pendingCheck is not { } pending) return;
        pendingCheck = null; pending.Timer.Stop();
        try
        {
            Store.Save(pending.Item, new JsonObject { ["done"] = pending.Done });
            pending.Item.Done = pending.Done; if (pending.Done) pending.Item.Incomplete = false;
        }
        catch (Exception error) { MessageBox.Show(this, error.Message, "青序", MessageBoxButton.OK, MessageBoxImage.Information); }
        PaintItem(pending.Item);
    }
    private void PaintItem(AgendaItem item)
    {
        foreach (var visual in visuals) if (visual.Id == item.Id) visual.Update(item);
        UpdateProgress();
    }
    private ContextMenu StateMenu(AgendaItem item)
    {
        var menu = new ContextMenu();
        foreach (var (label, incomplete, done) in new[] { ("标记未完成", true, false), ("标记已完成", false, true), ("清除状态", false, false) })
        {
            var option = new MenuItem { Header = label, Foreground = Brush("#658477") };
            option.Click += (_, _) => { CommitPendingCheck(); Run(() => { Store.Save(item, new JsonObject { ["incomplete"] = incomplete, ["done"] = done }); item.Incomplete = incomplete; item.Done = done; PaintItem(item); }); };
            menu.Items.Add(option);
        }
        if (!item.Class && !item.Weekly && item.Id.Length > 0 && item.Date.Length > 0)
        {
            menu.Items.Add(new Separator());
            var copy = new MenuItem { Header = "复制到明天", Foreground = Brush("#658477"), ToolTip = "明天会出现一条同样的任务，状态为待完成" };
            copy.Click += (_, _) => Run(() =>
            {
                var target = AgendaRules.ParseDate(item.Date).AddDays(1);
                if (target > AgendaRules.LastDate) { status.Text = "已经到学期最后一天"; return; }
                var day = Store.Day(target);
                var created = new AgendaItem { Id = "local:" + Guid.NewGuid().ToString("N"), Date = day.Date,
                    Order = Math.Max(0, day.Tasks.Select(task => task.Order).DefaultIfEmpty(0).Max()) + 1 };
                var fields = new JsonObject { ["title"] = item.DisplayTitle, ["done"] = false, ["incomplete"] = false, ["highlight"] = item.Highlight };
                if (item.Planned is double planned) fields["planned"] = planned;
                Store.Save(created, fields);
                status.Text = "已复制到 " + target.ToString("M月d日", CultureInfo.GetCultureInfo("zh-CN"));
            });
            menu.Items.Add(copy);
        }
        return menu;
    }
    private void NewTask()
    {
        if (page == 3) { status.Text = "周清单沿用 Excel 模板；请在日程页新建个人任务"; return; }
        var day = Store.Day(SelectedDate);
        var item = new AgendaItem { Id = "local:" + Guid.NewGuid().ToString("N"), Date = day.Date, Order = Math.Max(0, day.Tasks.Select(task => task.Order).DefaultIfEmpty(0).Max()) + 1 };
        Edit(item);
    }
    private void Edit(AgendaItem item) => Run(() =>
    {
        var editor = new TaskEditor(this, item);
        if (editor.ShowDialog() != true || editor.Result is null) return;
        if (item.Id.StartsWith("local:") && !Store.State.Operations.Any(operation => operation.Id == item.Id)) editor.Result["order"] = item.Order;
        Store.Save(item, editor.Result); Render();
    });
    private void OpenScheduling()
    {
        dropDate = page == 3 ? anchor : SelectedDate;
        var panel = new StackPanel();
        panel.Children.Add(Button("收起排期 ×", (_, _) => schedulePopup.IsOpen = false));
        if (page == 3)
        {
            var days = new WrapPanel(); var start = AgendaRules.ParseDate(Store.Day(anchor).WeekStart); dropDate = start;
            for (var index = 0; index < 7; index++)
            {
                var date = start.AddDays(index); var target = Button(date.ToString("ddd M/d", CultureInfo.GetCultureInfo("zh-CN")), (_, _) => { dropDate = date; status.Text = "排期目标：" + AgendaRules.DateText(date); });
                target.AllowDrop = true; target.Drop += (_, eventArgs) => Run(() => DropTask(eventArgs, date, null)); days.Children.Add(target);
            }
            panel.Children.Add(days);
        }
        var times = new WrapPanel { Width = 304 };
        for (var minute = 7 * 60; minute <= 22 * 60; minute += 30)
        {
            var time = AgendaRules.MinuteText(minute); var target = Button(time, (_, _) => status.Text = "把任务左侧拖动柄拖到时间点"); target.Width = 74;
            target.Background = Brush("#F1F7F3"); target.AllowDrop = true; target.Drop += (_, eventArgs) => Run(() => DropTask(eventArgs, dropDate, time)); times.Children.Add(target);
        }
        panel.Children.Add(times);
        schedulePopup.Child = new Border { Child = panel, Background = Brush("#FCFEFC"), BorderBrush = Brush("#DCE9E1"), BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(12), Padding = new Thickness(8) };
        schedulePopup.PlacementTarget = content; schedulePopup.Placement = PlacementMode.Right; schedulePopup.IsOpen = true;
    }
    private void DropTask(DragEventArgs eventArgs, DateOnly date, string? time)
    {
        if (eventArgs.Data.GetData("QingxuTask") is not AgendaItem item || item.Class || item.Id.StartsWith("semester:")) return;
        Store.Schedule(item, date, time); eventArgs.Handled = true; Render();
    }
    private async Task Synchronize()
    {
        if (syncing) return;
        if (Store.State.Operations.Count == 0) { Run(() => { Store.Reload(); Render(); OpenExcel(); }); return; }
        var sent = Store.Snapshot(); var versions = sent.Select(operation => operation.Version).Distinct().ToArray();
        syncing = true; syncButton.IsEnabled = false; status.Text = "正在后台同步 Excel…";
        foreach (var operation in sent) Store.SyncingIds.Add(operation.Id);
        var conflict = false; var succeeded = false;
        try
        {
            if (versions.Length != 1 || versions[0].Length == 0) throw new SyncConflictException();
            var result = await Task.Run(() => WorkbookSync.Write(Store.BookPath, versions[0], sent));
            Store.Complete(sent, result); succeeded = true;
        }
        catch (SyncConflictException) { conflict = true; }
        catch (Exception error) { status.Text = "同步未完成 · 本地修改已保留"; MessageBox.Show(this, error.Message, "暂未同步", MessageBoxButton.OK, MessageBoxImage.Information); }
        finally { syncing = false; syncButton.IsEnabled = true; Store.SyncingIds.Clear(); }
        // 同步期间工作簿正被替换，此时重绘会读到中间态；统一在 syncing 复位之后再刷新界面
        Render();
        if (succeeded && Store.State.Operations.Count == 0) OpenExcel();
        if (conflict) Run(() => ReviewConflicts());
    }
    private void OpenExcel() => Process.Start(new ProcessStartInfo(Store.BookPath) { UseShellExecute = true });
    private async void ReviewConflicts()
    {
        try
        {
            var review = new ConflictReview(this, Store);
            if (review.ShowDialog() != true) { status.Text = "尚未合并 · 两边修改均保留"; return; }
            Store.Resolve(review.Version, review.Operations, review.KeepLocal); Render();
            if (Store.State.Operations.Count > 0) await Synchronize();
        }
        catch (Exception error) { MessageBox.Show(this, error.Message, "暂未合并", MessageBoxButton.OK, MessageBoxImage.Information); }
    }
}
