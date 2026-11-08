using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
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
    private readonly Grid tabs = new();
    private readonly ScrollViewer scroll = new() { VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
    private readonly TextBlock status = new() { FontSize = 10, TextTrimming = TextTrimming.CharacterEllipsis };
    private readonly TextBlock progressText = new() { FontSize = 11 };
    private readonly ProgressBar progressBar = new() { Height = 4, Minimum = 0, Maximum = 1, BorderThickness = new Thickness(0) };
    private AgendaDay? currentDay;
    private readonly TextBlock dateLabel = new() { FontSize = 13.5, Margin = new Thickness(1, 3, 0, 0) };
    private readonly TextBox quickInput;
    private readonly TextBlock quickPlaceholder;
    private readonly System.Windows.Controls.Button quickSubmit;
    private readonly Forms.NotifyIcon tray;
    private readonly Forms.ContextMenuStrip trayMenu = new();
    private readonly System.Drawing.Icon trayImage;
    private readonly RegisteredWaitHandle? restoreWait;
    private readonly System.Windows.Controls.Button syncButton;
    private DateOnly anchor;
    private int page;
    private Dictionary<string, AbsorbRecord> absorbed = new();
    private (Point Origin, FrameworkElement Host, AgendaItem Item)? rowDrag;
    private (Point Origin, AgendaItem Item)? clickPending;
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
        absorbed = ReadSettings().Absorbed;
        anchor = AgendaRules.Today;
        page = Math.Max(0, Math.Min(4, startPage));
        if (anchor < AgendaRules.FirstDate || anchor > AgendaRules.LastDate) anchor = AgendaRules.FirstDate;
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
        trayMenu.Items.Add("打开 Excel", null, (_, _) => Run(OpenExcel));
        trayMenu.Items.Add("退出青序", null, (_, _) => Exit());
        tray.MouseClick += (_, eventArgs) => { if (eventArgs.Button == Forms.MouseButtons.Left) Restore(); };
        if (restore is not null) restoreWait = ThreadPool.RegisterWaitForSingleObject(restore, (_, _) => Dispatcher.BeginInvoke(Restore), null, Timeout.Infinite, false);
        var root = new Grid();
        for (var row = 0; row < 5; row++) root.RowDefinitions.Add(new RowDefinition { Height = row == 3 ? new GridLength(1, GridUnitType.Star) : GridLength.Auto });
        var header = new DockPanel { Cursor = Cursors.SizeAll, Margin = new Thickness(0, 0, 0, 14) };
        var actions = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Top };
        var pin = Button("📌", (_, _) => { Topmost = !Topmost; }); pin.ToolTip = "切换置顶";
        actions.Children.Add(pin);
        var close = Button("×", (_, _) => Close()); close.ToolTip = "收起到系统托盘，右键托盘可退出"; actions.Children.Add(close);
        DockPanel.SetDock(actions, Dock.Right); header.Children.Add(actions);
        var branding = new StackPanel(); branding.Children.Add(new TextBlock { Text = "青序", FontSize = 26, FontWeight = FontWeights.SemiBold, Foreground = Brush("#2F4C40") });
        dateLabel.Foreground = Brush("#7E978A"); branding.Children.Add(dateLabel); header.Children.Add(branding);
        branding.MouseLeftButtonDown += (_, eventArgs) => { if (eventArgs.LeftButton == MouseButtonState.Pressed) DragMove(); };
        root.Children.Add(header);
        Grid.SetRow(tabs, 1); tabs.Margin = new Thickness(0, 0, 0, 8); tabs.HorizontalAlignment = HorizontalAlignment.Stretch; root.Children.Add(tabs);
        var progressRow = new StackPanel { Margin = new Thickness(2, 0, 2, 10) };
        progressText.Foreground = Brush("#8AA394"); progressRow.Children.Add(progressText);
        progressBar.Foreground = Brush("#7FB794"); progressBar.Background = Brush("#E4EEE7"); progressBar.Margin = new Thickness(0, 5, 0, 0);
        progressRow.Children.Add(progressBar);
        Grid.SetRow(progressRow, 2); root.Children.Add(progressRow);
        scroll.Content = content; Grid.SetRow(scroll, 3); root.Children.Add(scroll);
        var footer = new StackPanel { Margin = new Thickness(0, 12, 0, 0) };
        var quickShell = new Border { Background = Brush("#FBFDFB"), BorderBrush = Brush("#DCE9E1"), BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(10), Padding = new Thickness(1) };
        var quickGrid = new Grid();
        quickGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        quickGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        quickGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        var plus = new TextBlock { Text = "＋", FontSize = 14, Foreground = Brush("#A7C3B4"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(9, 0, 1, 0), IsHitTestVisible = false };
        quickGrid.Children.Add(plus);
        var inputLayer = new Grid(); Grid.SetColumn(inputLayer, 1);
        quickPlaceholder = new TextBlock { Text = "添加任务，回车确认", FontSize = 12, Foreground = Brush("#B2C5BA"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(5, 0, 0, 0), IsHitTestVisible = false };
        quickInput = new TextBox { FontSize = 12.5, BorderThickness = new Thickness(0), Background = Brushes.Transparent, Foreground = Brush("#345347"),
            Padding = new Thickness(4, 7, 4, 7), VerticalContentAlignment = VerticalAlignment.Center, CaretBrush = Brush("#3F6B57"),
            ToolTip = "输入任务名称后按回车，或点右侧「添加」；记到当前显示的那一天" };
        inputLayer.Children.Add(quickPlaceholder); inputLayer.Children.Add(quickInput);
        quickGrid.Children.Add(inputLayer);
        quickSubmit = Button("添加", (_, _) => SubmitQuickAdd());
        quickSubmit.FontSize = 11.5; quickSubmit.Foreground = Brush("#3F6B57"); quickSubmit.Background = Brush("#EDF6F0"); quickSubmit.Padding = new Thickness(11, 5, 11, 5);
        quickSubmit.Margin = new Thickness(2, 3, 3, 3); quickSubmit.IsEnabled = false; quickSubmit.ToolTip = "把输入框里的内容加到当前日期";
        Grid.SetColumn(quickSubmit, 2); quickGrid.Children.Add(quickSubmit);
        quickShell.Child = quickGrid; footer.Children.Add(quickShell);
        quickShell.MouseEnter += (_, _) => { if (quickInput.IsEnabled) quickShell.BorderBrush = Brush("#BFD8C8"); };
        quickShell.MouseLeave += (_, _) => quickShell.BorderBrush = Brush("#DCE9E1");
        quickShell.MouseLeftButtonUp += (_, _) => { if (quickInput.IsEnabled && !quickInput.IsKeyboardFocusWithin) quickInput.Focus(); };
        quickInput.TextChanged += (_, _) => RefreshQuickAdd();
        quickInput.KeyDown += (_, eventArgs) => { if (eventArgs.Key != Key.Enter) return; eventArgs.Handled = true; SubmitQuickAdd(); };
        status.Margin = new Thickness(0, 7, 0, 0); status.Foreground = Brush("#8AA394"); footer.Children.Add(status);
        var footerActions = new WrapPanel { HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 4, 0, 0) };
        footerActions.Children.Add(Button("＋新建", (_, _) => NewTask()));
        syncButton = Button("同步 Excel", async (_, _) => await Synchronize()); syncButton.FontSize = 10; footerActions.Children.Add(syncButton);
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
                case Key.D5: SetPage(4); eventArgs.Handled = true; break;
            }
        };
        // 整条任务都能抓：按住任意位置（按钮和调整手柄除外）移动超过阈值就开拖，点一下仍然只是勾选
        PreviewMouseLeftButtonDown += (_, eventArgs) =>
        {
            rowDrag = null; clickPending = null;
            var source = eventArgs.OriginalSource as DependencyObject ?? eventArgs.Source as DependencyObject;
            if (source is null || Hit(source) is not { } target) return;
            var origin = eventArgs.GetPosition(this);
            if (target.Head) clickPending = (origin, target.Item);
            if (target.Item.Editable && !target.Item.Class) rowDrag = (origin, target.Host, target.Item);
        };
        PreviewMouseMove += (_, eventArgs) =>
        {
            if (eventArgs.LeftButton != MouseButtonState.Pressed) return;
            var position = eventArgs.GetPosition(this);
            if (rowDrag is { } drag && Far(position, drag.Origin))
            {
                rowDrag = null; clickPending = null;
                DragDrop.DoDragDrop(drag.Host, new DataObject("QingxuTask", drag.Item), DragDropEffects.Move);
                return;
            }
            if (clickPending is { } click && Far(position, click.Origin)) clickPending = null;
        };
        PreviewMouseLeftButtonUp += (_, eventArgs) =>
        {
            rowDrag = null;
            if (clickPending is not { } click) return;
            clickPending = null;
            if (eventArgs.ClickCount >= 2) HandleCheck(click.Item, true);
            else HandleCheck(click.Item, false);
        };
        Closing += OnClosing;
        Closed += (_, _) => { dayTimer.Stop(); watchTimer.Stop(); restoreWait?.Unregister(null); tray.Visible = false; tray.Dispose(); trayImage.Dispose(); trayMenu.Dispose(); };
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
        if (!exitRequested) { eventArgs.Cancel = true; Hide(); ShowInTaskbar = false; }
    }
    public void Exit()
    {
        if (syncing) { status.Text = "后台同步尚未结束，请稍候"; return; }
        exitRequested = true; Close(); Application.Current.Shutdown();
    }
    private DateOnly SelectedDate => anchor.AddDays(page is 3 or 4 ? 0 : page);
    private void SetPage(int value) { page = value; Render(); scroll.ScrollToTop(); }
    private void Navigate(int direction)
    {
        if (page == 4) return;
        if (page == 3) { ChangeAnchor(anchor.AddDays(direction * 7)); page = 3; Render(); return; }
        var next = page + direction;
        if (next >= 0 && next < 3) { if (anchor.AddDays(next) <= AgendaRules.LastDate) SetPage(next); }
        else { ChangeAnchor(anchor.AddDays(direction * 3)); page = direction > 0 ? 0 : Math.Min(2, AgendaRules.LastDate.DayNumber - anchor.DayNumber); Render(); }
    }
    private void ChangeAnchor(DateOnly date, bool follow = false)
    {
        if (date < AgendaRules.FirstDate || date > AgendaRules.LastDate) return;
        anchor = date; page = 0;
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
        // 左右箭头贴住两端、中间自动居中：整行铺满窗口宽度
        tabs.Children.Clear(); tabs.ColumnDefinitions.Clear();
        tabs.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        tabs.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        tabs.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        var previous = NavGlyph("‹", page == 4 ? "学期计划不分日期" : "上一天（键盘 ←）", () => Navigate(-1), page != 4);
        Grid.SetColumn(previous, 0); tabs.Children.Add(previous);
        var chips = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center };
        for (var index = 0; index < 3; index++)
        {
            var captured = index; var date = anchor.AddDays(index);
            var chip = Button(date == AgendaRules.Today ? "今天" : date.ToString("MM/dd"), (_, _) => SetPage(captured));
            chip.IsEnabled = date <= AgendaRules.LastDate; chip.Padding = new Thickness(5, 5, 5, 5);
            if (page == index) { chip.Background = Brush("#E1F0E7"); chip.Foreground = Brush("#2F4C40"); chip.FontWeight = FontWeights.SemiBold; }
            chips.Children.Add(chip);
        }
        var week = Button("本周", (_, _) => SetPage(3)); week.Padding = new Thickness(5, 5, 5, 5);
        if (page == 3) { week.Background = Brush("#E1F0E7"); week.Foreground = Brush("#2F4C40"); week.FontWeight = FontWeights.SemiBold; }
        chips.Children.Add(week);
        var semester = Button("学期", (_, _) => SetPage(4)); semester.Padding = new Thickness(5, 5, 5, 5);
        if (page == 4) { semester.Background = Brush("#E1F0E7"); semester.Foreground = Brush("#2F4C40"); semester.FontWeight = FontWeights.SemiBold; }
        chips.Children.Add(semester);
        var picker = Button("选日期⌄", (_, _) => PickDate()); picker.FontSize = 11.5; picker.Foreground = Brush("#8AA394"); picker.Padding = new Thickness(7, 5, 7, 5);
        chips.Children.Add(picker);
        Grid.SetColumn(chips, 1); tabs.Children.Add(chips);
        var next = NavGlyph("›", page == 4 ? "学期计划不分日期" : "下一天（键盘 →）", () => Navigate(1), page != 4);
        Grid.SetColumn(next, 2); tabs.Children.Add(next);
        var day = Store.Day(page >= 3 ? anchor : SelectedDate);
        currentDay = day;
        if (page == 3)
        {
            var weekStart = AgendaRules.ParseDate(day.WeekStart);
            dateLabel.Text = day.WeekName + " · " + weekStart.ToString("M.d") + " — " + weekStart.AddDays(6).ToString("M.d");
        }
        else if (page == 4) dateLabel.Text = "学期计划 · 至 " + AgendaRules.LastDate.ToString("M月d日");
        else dateLabel.Text = SelectedDate.ToString("M月d日 dddd", CultureInfo.GetCultureInfo("zh-CN"));
        content.Children.Clear(); visuals.Clear();
        if (page == 4)
        {
            Section("学期计划");
            foreach (var item in day.Semester) content.Children.Add(Card(item, true));
        }
        else if (page == 3)
        {
            var weekBanner = ExamBanner(day);
            if (weekBanner is not null) content.Children.Add(weekBanner);
            Section("本周常规任务");
            foreach (var item in day.Weekly) content.Children.Add(Card(item, true));
            Section("每周课程");
            foreach (var item in day.WeekCourses)
                content.Children.Add(new TextBlock { Text = AgendaRules.ParseDate(item.Date).ToString("ddd", CultureInfo.GetCultureInfo("zh-CN")) + "  " + item.Time + "\n" + item.Title, TextWrapping = TextWrapping.Wrap, Foreground = Brush("#658477"), Margin = new Thickness(0, 5, 0, 6) });
        }
        else
        {
            // 日程就是一条时间流：固定课程和自己的安排混排，空档压成一条提示，拖动时展开刻度
            var date = AgendaRules.ParseDate(day.Date);
            Section("日程");
            var banner = ExamBanner(day);
            if (banner is not null) content.Children.Add(banner);
            var lastEnd = RenderFlow(day, date);
            if (lastEnd > 0) content.Children.Add(ScheduleTail(date, lastEnd));
            Section("待办", AutoPlanChip());
            foreach (var item in day.Tasks) content.Children.Add(Card(item, false));
        }
        status.Text = (Store.ImportedLegacy ? "已读取旧版待同步内容 · " : "") + (Store.State.Operations.Count > 0 ? "本地已保存 · 待同步 " + Store.State.Operations.Count + " 项" : "已与 Excel 同步");
        status.ToolTip = status.Text; scroll.ScrollToVerticalOffset(offset); UpdateProgress(); RefreshQuickAdd();
    }
    /// <summary>同步快速添加栏的可用状态：周页不支持新增，输入为空时按钮置灰。</summary>
    private void RefreshQuickAdd()
    {
        var weeklyPage = page >= 3;
        quickInput.IsEnabled = !weeklyPage;
        quickSubmit.IsEnabled = !weeklyPage && quickInput.Text.Trim().Length > 0;
        quickPlaceholder.Text = weeklyPage ? "清单读取 Excel，切回日程页可添加" : "添加任务，回车确认";
        quickPlaceholder.Visibility = quickInput.Text.Length == 0 ? Visibility.Visible : Visibility.Collapsed;
    }
    public bool SubmitQuickAdd()
    {
        if (!QuickAdd(quickInput.Text)) return false;
        quickInput.Clear(); quickInput.Focus(); RefreshQuickAdd(); return true;
    }
    private Border NavGlyph(string glyph, string tip, Action action, bool enabled = true)
    {
        var label = new TextBlock { Text = glyph, FontSize = 20, Foreground = Brush(enabled ? "#7FA493" : "#D8E3DB"), HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, -3, 0, 0) };
        var border = new Border { Width = 32, Height = 30, CornerRadius = new CornerRadius(10), Background = Brushes.Transparent, Child = label, Cursor = enabled ? Cursors.Hand : Cursors.Arrow, ToolTip = tip, VerticalAlignment = VerticalAlignment.Center };
        if (!enabled) return border;
        border.MouseEnter += (_, _) => { border.Background = Brush("#EDF5F0"); label.Foreground = Brush("#3F6B57"); };
        border.MouseLeave += (_, _) => { border.Background = Brushes.Transparent; label.Foreground = Brush("#7FA493"); };
        border.MouseLeftButtonUp += (_, eventArgs) => { action(); eventArgs.Handled = true; };
        return border;
    }
    private void UpdateProgress()
    {
        if (currentDay is null) { progressText.Text = ""; progressBar.Value = 0; return; }
        var items = page == 4 ? currentDay.Semester.ToList() : page == 3 ? currentDay.Weekly.ToList() : currentDay.Tasks.Concat(currentDay.Schedule.Where(item => item.Editable && !item.Class)).ToList();
        var done = items.Count(item => item.Done);
        var label = page == 4 ? "学期" : page == 3 ? "本周" : "今日";
        progressText.Text = items.Count == 0 ? label + "暂无待办" : $"{label} {done}/{items.Count} 已完成 · 还剩 {items.Count - done} 项";
        progressBar.Value = items.Count == 0 ? (page >= 3 ? 0 : 1) : (double)done / items.Count;
    }
    public bool QuickAdd(string text)
    {
        var title = text.Trim();
        if (title.Length == 0) return false;
        if (page >= 3) { status.Text = "清单读取 Excel，请在日程页添加"; return false; }
        var added = false;
        Run(() =>
        {
            var day = Store.Day(SelectedDate);
            var item = new AgendaItem { Id = "local:" + Guid.NewGuid().ToString("N"), Date = day.Date,
                Order = Math.Max(0, day.Tasks.Select(task => task.Order).DefaultIfEmpty(0).Max()) + 1 };
            Store.Save(item, new JsonObject { ["title"] = title, ["done"] = false, ["incomplete"] = false });
            Render(); status.Text = "已添加「" + title + "」· 待同步"; added = true;
        });
        return added;
    }
    private void Section(string title, UIElement? trailing = null)
    {
        var label = new TextBlock { Text = title, Foreground = Brush("#8AA394"), FontWeight = FontWeights.SemiBold, FontSize = 12, Margin = new Thickness(0, 0, 0, 0), VerticalAlignment = VerticalAlignment.Center };
        if (trailing is null) { content.Children.Add(new Border { Child = label, Margin = new Thickness(0, 10, 0, 7) }); return; }
        var row = new DockPanel { Margin = new Thickness(0, 10, 0, 7) };
        DockPanel.SetDock(trailing, Dock.Right); row.Children.Add(trailing); row.Children.Add(label);
        content.Children.Add(row);
    }
    private sealed class LocalSettings
    {
        /// <summary>「拖进课内任务」的撤销记录：课程 id → 并进去的标题 + 原来的课内任务。</summary>
        public Dictionary<string, AbsorbRecord> Absorbed { get; set; } = new();
    }
    /// <summary>一门课被并进来的任务，按先后排队，撤销时从最后一条开始退。</summary>
    private sealed class AbsorbRecord { public string Previous { get; set; } = ""; public List<AbsorbEntry> Entries { get; set; } = new(); }
    private sealed class AbsorbEntry { public string Title { get; set; } = ""; public string Id { get; set; } = ""; }
    private string SettingsPath => Store.BookPath + ".qingxu-settings.json";
    private LocalSettings ReadSettings()
    {
        try { if (File.Exists(SettingsPath)) return JsonSerializer.Deserialize<LocalSettings>(File.ReadAllText(SettingsPath), AgendaRules.JsonOptions) ?? new LocalSettings(); }
        catch (Exception) { }
        return new LocalSettings();
    }
    private void SaveSettings()
    {
        try { File.WriteAllText(SettingsPath + ".pending", JsonSerializer.Serialize(new LocalSettings { Absorbed = absorbed }, AgendaRules.JsonOptions)); File.Move(SettingsPath + ".pending", SettingsPath, true); }
        catch (Exception) { }
    }
    /// <summary>从按下的控件往上找它属于哪一条任务；按钮和调整时长的手柄不算拖动的起点。</summary>
    private static (FrameworkElement Host, AgendaItem Item, bool Head)? Hit(DependencyObject source)
    {
        var head = false;
        for (DependencyObject? node = source; node is not null; node = Parent(node))
        {
            if (node is ButtonBase) return null;
            if (node is not FrameworkElement element) continue;
            if (element.Tag as string == "resize") return null;
            if (element.Tag as string == "head") head = true;
            if (element.Tag is AgendaItem item) return (element, item, head);
        }
        return null;
    }
    private static DependencyObject? Parent(DependencyObject node) => node is Visual || node is System.Windows.Media.Media3D.Visual3D ? VisualTreeHelper.GetParent(node) : LogicalTreeHelper.GetParent(node);
    private static bool Far(Point position, Point origin) => Math.Abs(position.X - origin.X) >= SystemParameters.MinimumHorizontalDragDistance || Math.Abs(position.Y - origin.Y) >= SystemParameters.MinimumVerticalDragDistance;
    private sealed record HeadParts(Grid Head, TextBlock Label, Border Box, System.Windows.Shapes.Path Tick, System.Windows.Shapes.Path Slash);
    /// <summary>复选框 + 标题 + 单击/双击/右键，日程流和待办卡片共用同一套。</summary>
    private HeadParts BuildHead(AgendaItem item, double fontSize)
    {
        var label = new TextBlock { Text = item.DisplayTitle, TextWrapping = TextWrapping.Wrap, FontSize = fontSize };
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
        head.Tag = "head";
        return new HeadParts(head, label, box, tick, slash);
    }
    /// <summary>完成 / 未完成 / 强调 / 考试四套配色，只换颜色不重建控件。</summary>
    private void PaintRow(AgendaItem current, Border shell, TextBlock? label, Border? box, System.Windows.Shapes.Path? tick, System.Windows.Shapes.Path? slash, ToggleButton? star, bool exam)
    {
        shell.Background = Brush(current.Incomplete ? "#F2F3F2" : current.Done ? "#F0F8F3" : exam ? "#FDF2EE" : current.Highlight ? "#FFF7DF" : "#FFFFFF");
        if (label is not null)
        {
            label.Foreground = Brush(current.Incomplete ? "#969C98" : current.Done ? "#8AA197" : exam ? "#8A3C28" : "#2D463C");
            label.TextDecorations = current.Done ? TextDecorations.Strikethrough : null;
            label.FontWeight = current.Highlight || exam ? FontWeights.Bold : FontWeights.Normal;
        }
        if (box is not null) { box.Background = current.Incomplete ? Brush("#ECEEEC") : current.Done ? Brush("#EDF5EF") : Brushes.White; box.BorderBrush = Brush(current.Incomplete ? "#A7ADA9" : "#91A298"); }
        if (tick is not null) tick.Visibility = current.Done ? Visibility.Visible : Visibility.Collapsed;
        if (slash is not null) slash.Visibility = current.Incomplete ? Visibility.Visible : Visibility.Collapsed;
        if (star is not null) { star.IsChecked = current.Highlight; star.Content = current.Highlight ? "★" : "☆"; star.Foreground = Brush(current.Highlight ? "#B68B31" : "#A7BBB0"); }
    }
    private Border Card(AgendaItem item, bool weekly)
    {
        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        grid.ColumnDefinitions.Add(new ColumnDefinition());
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        var handle = new TextBlock { Text = "⠿", Foreground = Brush("#9BB3A5"), Cursor = Cursors.SizeAll, Margin = new Thickness(0, 2, 8, 0), ToolTip = "拖动排序，也可以拖进上面的日程" };
        if (item.Class) handle.Visibility = Visibility.Collapsed;
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
        var parts = BuildHead(item, 14);
        text.Children.Add(parts.Head);
        if (item.Detail.Length > 0 || item.Class) text.Children.Add(new TextBlock { Text = item.Class ? "课内安排" : item.Detail, FontSize = 11, Foreground = Brush("#8B9B92"), Margin = new Thickness(18, 4, 0, 0), TextWrapping = TextWrapping.Wrap });
        Grid.SetColumn(text, 1); grid.Children.Add(text);
        var actions = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Top };
        var star = new ToggleButton { FontSize = 16, Padding = new Thickness(3, 0, 3, 0), Margin = new Thickness(3, 0, 3, 0), Background = Brushes.Transparent,
            BorderThickness = new Thickness(0), Cursor = Cursors.Hand, ToolTip = "强调", VerticalAlignment = VerticalAlignment.Top };
        actions.Children.Add(star);
        actions.Children.Add(Button("···", (_, _) => Edit(item)));
        var delete = Button((item.Weekly || item.Semester) && !weekly ? "撤" : "删", (_, _) => Run(() =>
        {
            var message = (item.Weekly || item.Semester) && !weekly ? "只取消这一天的安排，任务仍保留在清单中。" : item.Class ? "只删除课内任务，保留课程本身。" : "确定删除这个任务吗？";
            if (MessageBox.Show(this, message, "青序", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;
            Store.Remove(item, weekly); Render();
        })); delete.ToolTip = (item.Weekly || item.Semester) && !weekly ? "取消安排，保留清单任务" : "删除任务"; actions.Children.Add(delete);
        Grid.SetColumn(actions, 2); grid.Children.Add(actions);
        var border = new Border { CornerRadius = new CornerRadius(8), BorderBrush = Brush("#E3EDE7"), BorderThickness = new Thickness(0, 0, 0, 1),
            Margin = new Thickness(0, 0, 0, 4), Padding = new Thickness(9, 6, 6, 6), Child = grid, AllowDrop = true, Tag = item };
        void Paint(AgendaItem current) => PaintRow(current, border, parts.Label, parts.Box, parts.Tick, parts.Slash, star, false);
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
            // 从日程拖回待办 = 取消时间，交给用户重新排
            if (!weekly && source.Start.Length > 0 && item.Start.Length == 0)
            {
                Store.Save(source, new JsonObject { ["start"] = "", ["planned"] = "" });
                eventArgs.Handled = true; Render(); status.Text = "已取消「" + source.DisplayTitle + "」的时间，回到待办"; return;
            }
            if (weekly && (source.Weekly != item.Weekly || source.Semester != item.Semester) || !weekly && source.Date != item.Date || !weekly && source.Start.Length > 0 || !weekly && item.Start.Length > 0) throw new InvalidOperationException("请在同一个清单内排序，有时间的任务请直接拖进上面的日程。");
            var day = Store.Day(weekly ? anchor : SelectedDate);
            var list = page == 4 ? day.Semester.ToList() : weekly ? day.Weekly.ToList() : day.Tasks.ToList();
            var moving = list.FirstOrDefault(candidate => candidate.Id == source.Id);
            if (moving is null) throw new InvalidOperationException("请在同一个列表内拖动排序。");
            list.Remove(moving); var position = list.FindIndex(candidate => candidate.Id == item.Id);
            if (position < 0) return;
            if (eventArgs.GetPosition(border).Y > border.ActualHeight / 2) position++;
            list.Insert(position, moving);
            for (var index = 0; index < list.Count; index++) Store.Save(list[index], new JsonObject { ["order"] = index + 1 });
            eventArgs.Handled = true; Render();
        });
        return border;
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
        if (!item.Class && !item.Weekly && !item.Semester && item.Id.Length > 0 && item.Date.Length > 0)
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
        if (item.Editable && !item.Class && item.Start.Length > 0)
        {
            menu.Items.Add(new Separator());
            var clear = new MenuItem { Header = "取消时间，放回待办", Foreground = Brush("#658477") };
            clear.Click += (_, _) => Run(() =>
            {
                if (item.Weekly || item.Semester) Store.Remove(item, false);
                else Store.Save(item, new JsonObject { ["start"] = "", ["planned"] = "" });
                Render(); status.Text = "已取消「" + item.DisplayTitle + "」的时间";
            });
            menu.Items.Add(clear);
        }
        if (item.Editable && !item.Class)
        {
            var unschedule = (item.Weekly || item.Semester) && page != 3;
            menu.Items.Add(new Separator());
            var remove = new MenuItem { Header = unschedule ? "取消这一天的安排" : "删除这条", Foreground = Brush("#A9482D") };
            remove.Click += (_, _) => Run(() =>
            {
                if (MessageBox.Show(this, unschedule ? "只取消这一天的安排，任务仍保留在清单中。" : "确定删除这个任务吗？", "青序", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;
                Store.Remove(item, page >= 3); Render();
            });
            menu.Items.Add(remove);
        }
        return menu;
    }
    private void NewTask()
    {
        if (page >= 3) { status.Text = "清单读取 Excel；请在日程页新建个人任务"; return; }
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
    /// <summary>一条时间流：课程和自己的安排混排，空档压成一行提示；返回当天最后一条的结束时间。</summary>
    private int RenderFlow(AgendaDay day, DateOnly date)
    {
        var entries = day.Schedule.Where(item => Range(item.Time) is not null)
            .Select(item => (Item: item, Span: Span(item)))
            .OrderBy(pair => pair.Span.Start).ThenBy(pair => pair.Item.Class ? 0 : 1).ToList();
        if (entries.Count == 0)
        {
            var blank = GhostDrop("放到这里 · 落点就是开始时间", 10, 42, (source, args) =>
            {
                status.Text = PlaceAt(source, date, 8 * 60, (int)Math.Max(15, source.Planned ?? 60));
                _ = AutoSyncAfterSchedule();
            });
            blank.Margin = new Thickness(52, 2, 0, 2);
            content.Children.Add(blank);
            return 0;
        }
        var cursor = -1;
        foreach (var entry in entries)
        {
            if (cursor >= 0 && entry.Span.Start > cursor) AddGap(cursor, entry.Span.Start, date);
            AddFlowRow(entry.Item, entry.Span, date);
            cursor = Math.Max(cursor, entry.Span.End);
        }
        // 没有时间（或时间格空着）的课程/事项也要留一行，避免直接从界面上消失
        foreach (var item in day.Schedule.Where(item => Range(item.Time) is null && item.Title.Length > 0))
            content.Children.Add(new TextBlock { Text = (item.Class ? item.Title : item.DisplayTitle) + " · 时间待定", FontSize = 11, Foreground = Brush("#A9BDB2"), Margin = new Thickness(54, 2, 0, 2), TextWrapping = TextWrapping.Wrap });
        var summary = DaySummary(entries);
        if (summary is not null) content.Children.Add(summary);
        return cursor;
    }
    private void AddFlowRow(AgendaItem item, (int Start, int End) span, DateOnly date)
    {
        var exam = item.Exam;
        var grid = new Grid { Margin = new Thickness(0, 1, 0, 1) };
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(52) });
        grid.ColumnDefinitions.Add(new ColumnDefinition());
        grid.Children.Add(FlowGutter(item, span, exam));
        var stack = new StackPanel();
        var head = new Grid();
        head.ColumnDefinitions.Add(new ColumnDefinition());
        head.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        head.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        TextBlock? label = null; Border? box = null;
        System.Windows.Shapes.Path? tick = null; System.Windows.Shapes.Path? slash = null;
        if (item.Editable && !item.Class)
        {
            var parts = BuildHead(item, 13.5);
            label = parts.Label; box = parts.Box; tick = parts.Tick; slash = parts.Slash;
            head.Children.Add(parts.Head);
        }
        else head.Children.Add(new TextBlock { Text = item.Class ? item.Title : item.DisplayTitle, FontSize = 13.5, FontWeight = FontWeights.SemiBold, Foreground = Brush("#2D463C"), TextWrapping = TextWrapping.Wrap });
        if (exam) { var badge = ExamBadge(item); Grid.SetColumn(badge, 1); head.Children.Add(badge); }
        var actions = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Top, Margin = new Thickness(2, -1, 0, 0) };
        if (item.Editable && !item.Class && item.Start.Length > 0)
        {
            var undo = SmallButton("↺", "撤销时间排期，放回待办");
            undo.Click += (_, _) => UndoSchedule(item);
            actions.Children.Add(undo);
        }
        ToggleButton? star = null;
        if (item.Editable && !item.Class)
        {
            star = new ToggleButton { FontSize = 13.5, Padding = new Thickness(2, 0, 2, 0), Margin = new Thickness(1, 0, 1, 0), Background = Brushes.Transparent, BorderThickness = new Thickness(0), Cursor = Cursors.Hand, ToolTip = "强调" };
            actions.Children.Add(star);
        }
        if (item.Editable) actions.Children.Add(Button("···", (_, _) => Edit(item)));
        Grid.SetColumn(actions, 2); head.Children.Add(actions);
        stack.Children.Add(head);
        var meta = new List<string> { item.Time.Contains('—') ? AgendaRules.MinuteText(span.Start) + "—" + AgendaRules.MinuteText(span.End) : AgendaRules.MinuteText(span.Start) };
        if (item.Subject.Length > 0) meta.Add(item.Subject);
        if (item.Note.Length > 0) meta.Add(item.Note);
        var metaText = new TextBlock { Text = string.Join(" · ", meta), FontSize = 10.5, Foreground = Brush(exam ? "#B46A52" : "#8B9B92"), Margin = new Thickness(item.Editable && !item.Class ? 24 : 0, 2, 0, 0), TextWrapping = TextWrapping.Wrap };
        stack.Children.Add(metaText);
        if (item.Class) stack.Children.Add(CourseTaskStrip(item));
        var accent = new Border { Width = 3, CornerRadius = new CornerRadius(2), Background = Brush(exam ? "#C0563F" : item.Class ? "#6E9B84" : "#C6DACD"), Margin = new Thickness(0, 1, 7, 1) };
        var dock = new DockPanel(); DockPanel.SetDock(accent, Dock.Left); dock.Children.Add(accent); dock.Children.Add(stack);
        var body = new Grid(); body.Children.Add(dock);
        var shell = new Border { CornerRadius = new CornerRadius(9), BorderThickness = new Thickness(1), BorderBrush = Brush("#E4EEE8"), Padding = new Thickness(7, 6, 4, 6), MinHeight = Math.Clamp((span.End - span.Start) * 0.45, 34, 108), Child = body, AllowDrop = true, Tag = item };
        Grid.SetColumn(shell, 1); grid.Children.Add(shell);
        void Paint(AgendaItem current) => PaintRow(current, shell, label, box, tick, slash, star, exam);
        if (item.Editable && !item.Class)
        {
            var handle = new Border { Height = 7, Background = Brushes.Transparent, Cursor = Cursors.SizeNS, VerticalAlignment = VerticalAlignment.Bottom, ToolTip = "拖动底边调整时长", Tag = "resize" };
            body.Children.Add(handle);
            AttachResize(handle, shell, item, span, metaText);
        }
        if (star is not null) star.Click += (_, _) =>
        {
            var value = star.IsChecked == true;
            Run(() => { Store.Save(item, new JsonObject { ["highlight"] = value }); item.Highlight = value; Paint(item); });
        };
        AttachFlowDrop(shell, item, date);
        Paint(item);
        if (item.Id.Length > 0) visuals.Add((item.Id, Paint));
        content.Children.Add(grid);
    }
    /// <summary>时间轴左栏：竖线、圆点和开始时间，整栏就是拖动这一条的把手。</summary>
    private UIElement FlowGutter(AgendaItem item, (int Start, int End) span, bool exam)
    {
        var gutter = new Grid { Background = Brushes.Transparent, Cursor = item.Editable ? Cursors.SizeAll : Cursors.Arrow, MinHeight = 20, ToolTip = item.Editable ? "整条都能拖：按住任意位置拖动改时间，拖回待办取消时间" : null };
        var line = new System.Windows.Shapes.Rectangle { Width = 1.2, Fill = Brush("#E4EEE8"), HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 0, 6, 0), IsHitTestVisible = false };
        var dot = new System.Windows.Shapes.Ellipse { Width = 7, Height = 7, Fill = Brush(exam ? "#C0563F" : item.Class ? "#6E9B84" : "#C6DACD"), Stroke = Brush("#FCFEFC"), StrokeThickness = 1.5, HorizontalAlignment = HorizontalAlignment.Right, VerticalAlignment = VerticalAlignment.Top, Margin = new Thickness(0, 10, 3.1, 0), IsHitTestVisible = false };
        var time = new TextBlock { Text = AgendaRules.MinuteText(span.Start), FontSize = 10.5, Foreground = Brush(exam ? "#B4482F" : "#8AA394"), HorizontalAlignment = HorizontalAlignment.Right, VerticalAlignment = VerticalAlignment.Top, Margin = new Thickness(0, 6, 13, 0), IsHitTestVisible = false };
        gutter.Children.Add(line); gutter.Children.Add(dot); gutter.Children.Add(time);
        return gutter;
    }
    private Border ExamBadge(AgendaItem item) => new()
    {
        Background = Brush("#FBEDE8"), BorderBrush = Brush("#E7C1B3"), BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(5),
        Padding = new Thickness(5, 1, 5, 1), Margin = new Thickness(4, 1, 0, 0), VerticalAlignment = VerticalAlignment.Top,
        Child = new TextBlock { Text = ExamWord(item), FontSize = 10, Foreground = Brush("#A9482D"), FontWeight = FontWeights.SemiBold }
    };
    /// <summary>角标写字：优先用具体考法（小测/期中/期末/机考），都没有就写「考试」。</summary>
    private static string ExamWord(AgendaItem item)
    {
        foreach (var word in new[] { "小测", "期中", "期末", "机考", "统考", "补考", "测验", "考试" })
            if (item.Category == word || item.Title.Contains(word) || item.TaskTitle.Contains(word) || item.Note.Contains(word)) return word;
        return item.Category.Length > 0 ? item.Category : "考试";
    }
    /// <summary>课程底下的课内任务：把待办拖进来就并进这门课。</summary>
    private Border CourseTaskStrip(AgendaItem course)
    {
        // 有课内任务就显示一行勾选框；没有就留一条几乎看不见的投放带，拖进来才展开
        if (course.TaskTitle.Length == 0)
        {
            var empty = GhostDrop("放到课内任务", 8, 30, (source, _) => status.Text = AbsorbIntoClass(source, course));
            empty.Margin = new Thickness(0, 3, 2, 0);
            return empty;
        }
        var strip = new Border { Background = Brush("#F6FAF7"), BorderBrush = Brush("#DCEBE2"), BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(6), Padding = new Thickness(6, 3, 6, 3), Margin = new Thickness(0, 5, 2, 0), AllowDrop = true, ToolTip = "把待办拖到这里，就成了这门课的课内任务" };
        var parts = BuildHead(course, 11.5);
        var layout = new Grid();
        layout.ColumnDefinitions.Add(new ColumnDefinition());
        layout.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        layout.Children.Add(parts.Head);
        if (absorbed.TryGetValue(course.Id, out var record) && record.Entries.Count > 0 && course.TaskTitle.Contains(record.Entries[^1].Title))
        {
            var undo = SmallButton("↺", "撤销并入，把这件任务放回待办");
            undo.Click += (_, _) => UndoAbsorb(course);
            Grid.SetColumn(undo, 1); layout.Children.Add(undo);
        }
        strip.Child = layout;
        void Paint(AgendaItem current)
        {
            parts.Label.Foreground = Brush(current.Incomplete ? "#969C98" : current.Done ? "#8AA197" : "#3E5A4E");
            parts.Label.TextDecorations = current.Done ? TextDecorations.Strikethrough : null;
            parts.Box.Background = current.Incomplete ? Brush("#ECEEEC") : current.Done ? Brush("#EDF5EF") : Brushes.White;
            parts.Tick.Visibility = current.Done ? Visibility.Visible : Visibility.Collapsed;
            parts.Slash.Visibility = current.Incomplete ? Visibility.Visible : Visibility.Collapsed;
        }
        Paint(course); visuals.Add((course.Id, Paint));
        strip.DragOver += (_, eventArgs) =>
        {
            var ok = eventArgs.Data.GetDataPresent("QingxuTask");
            eventArgs.Effects = ok ? DragDropEffects.Move : DragDropEffects.None; eventArgs.Handled = true;
            strip.BorderBrush = Brush(ok ? "#7FB794" : "#DCEBE2");
        };
        strip.DragLeave += (_, _) => strip.BorderBrush = Brush("#DCEBE2");
        strip.Drop += (_, eventArgs) =>
        {
            eventArgs.Handled = true; strip.BorderBrush = Brush("#DCEBE2");
            try { if (eventArgs.Data.GetData("QingxuTask") is AgendaItem source) status.Text = AbsorbIntoClass(source, course); }
            catch (Exception error) { status.Text = error.Message; }
        };
        return strip;
    }
    /// <summary>平时不占版面的投放区：拖任务进来才展开并显示一行提示。</summary>
    private Border GhostDrop(string hint, double collapsed, double expanded, Action<AgendaItem, DragEventArgs> drop)
    {
        var text = new TextBlock { Text = hint, FontSize = 10.5, Foreground = Brush("#8FB3A0"), HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center, IsHitTestVisible = false };
        var host = new Border { Height = collapsed, Background = Brushes.Transparent, CornerRadius = new CornerRadius(7), AllowDrop = true, Focusable = false };
        void Collapse() { host.Height = collapsed; host.Background = Brushes.Transparent; host.Child = null; }
        host.DragEnter += (_, eventArgs) =>
        {
            if (!eventArgs.Data.GetDataPresent("QingxuTask")) return;
            host.Height = expanded; host.Background = Brush("#F3FAF6"); host.Child = text;
        };
        host.DragLeave += (_, _) => Collapse();
        host.DragOver += (_, eventArgs) => { eventArgs.Effects = eventArgs.Data.GetDataPresent("QingxuTask") ? DragDropEffects.Move : DragDropEffects.None; eventArgs.Handled = true; };
        host.Drop += (_, eventArgs) =>
        {
            eventArgs.Handled = true; Collapse();
            try { if (eventArgs.Data.GetData("QingxuTask") is AgendaItem source) drop(source, eventArgs); }
            catch (Exception error) { status.Text = error.Message; }
        };
        return host;
    }
    /// <summary>小小的圆形操作键，用在行内撤销。</summary>
    private System.Windows.Controls.Button SmallButton(string glyph, string tip)
    {
        var button = new System.Windows.Controls.Button { Content = glyph, FontSize = 12, Padding = new Thickness(3, 0, 3, 0), Background = Brushes.Transparent, BorderThickness = new Thickness(0), Foreground = Brush("#93A99C"), Cursor = Cursors.Hand, ToolTip = tip, VerticalAlignment = VerticalAlignment.Center, Focusable = false };
        button.MouseEnter += (_, _) => button.Foreground = Brush("#3F6B57");
        button.MouseLeave += (_, _) => button.Foreground = Brush("#93A99C");
        return button;
    }
    /// <summary>空档：平时只是一行「空 2 小时」，拖任务进来时才展开成时间刻度和落点线。</summary>
    private void AddGap(int from, int to, DateOnly date)
    {
        var minutes = to - from;
        if (minutes < 5) { content.Children.Add(new Border { Height = 2 }); return; }
        var collapsed = minutes < 20 ? 9d : 16d;
        var hint = new TextBlock { Text = minutes < 20 ? "·" : "空 " + AgendaRules.Human(minutes), FontSize = 10, Foreground = Brush("#C2D2C8"), HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center, IsHitTestVisible = false };
        var host = new Border { Height = collapsed, Child = hint, CornerRadius = new CornerRadius(8), Background = Brushes.Transparent, AllowDrop = true, ToolTip = "把任务拖到这里，按刻度选开始时间" };
        var expanded = Math.Clamp(minutes * 0.55, 52, 280);
        Grid? ruler = null; Border? marker = null; TextBlock? badge = null;
        void Collapse() { host.Height = collapsed; host.Child = hint; host.Background = Brushes.Transparent; }
        host.DragEnter += (_, eventArgs) =>
        {
            if (!eventArgs.Data.GetDataPresent("QingxuTask")) return;
            if (ruler is null) (ruler, marker, badge) = GapRuler(from, to, expanded);
            host.Height = expanded; host.Child = ruler; host.Background = Brush("#F7FBF8");
        };
        host.DragLeave += (_, _) => Collapse();
        host.DragOver += (_, eventArgs) =>
        {
            if (!eventArgs.Data.GetDataPresent("QingxuTask")) { eventArgs.Effects = DragDropEffects.None; eventArgs.Handled = true; return; }
            eventArgs.Effects = DragDropEffects.Move; eventArgs.Handled = true;
            var y = eventArgs.GetPosition(host).Y;
            var minute = MinuteAt(y, expanded, from, to);
            if (marker is not null) marker.Margin = new Thickness(0, Math.Clamp(y - 1, 0, expanded - 2), 0, 0);
            if (badge is not null) { badge.Text = AgendaRules.MinuteText(minute); badge.Margin = new Thickness(7, Math.Clamp(y - 19, 0, expanded - 18), 0, 0); }
        };
        host.Drop += (_, eventArgs) =>
        {
            eventArgs.Handled = true;
            try
            {
                if (eventArgs.Data.GetData("QingxuTask") is not AgendaItem source) return;
                var minute = MinuteAt(eventArgs.GetPosition(host).Y, expanded, from, to);
                var length = (int)Math.Max(15, source.Planned ?? Math.Clamp(minutes, 15, 60));
                status.Text = PlaceAt(source, date, minute, length);
                Collapse();
                _ = AutoSyncAfterSchedule();
            }
            catch (Exception error) { Collapse(); status.Text = error.Message; }
        };
        content.Children.Add(host);
    }
    private (Grid Root, Border Marker, TextBlock Badge) GapRuler(int from, int to, double height)
    {
        var root = new Grid();
        root.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(52) });
        root.ColumnDefinitions.Add(new ColumnDefinition());
        var gutter = new Canvas { Width = 52, Height = height };
        var ticks = new Canvas { Height = height, ClipToBounds = true };
        var total = to - from;
        for (var minute = (from / 30 + 1) * 30; minute < to; minute += 30)
        {
            var y = (minute - from) * height / total;
            var hour = minute % 60 == 0;
            var line = new System.Windows.Shapes.Rectangle { Width = hour ? 300 : 14, Height = 1, Fill = Brush(hour ? "#E4EFE8" : "#F0F6F2"), IsHitTestVisible = false };
            Canvas.SetLeft(line, 0); Canvas.SetTop(line, y); ticks.Children.Add(line);
            if (!hour) continue;
            var label = new TextBlock { Text = AgendaRules.MinuteText(minute), Width = 40, FontSize = 9.5, Foreground = Brush("#B4C6BB"), TextAlignment = TextAlignment.Right, IsHitTestVisible = false };
            Canvas.SetLeft(label, 0); Canvas.SetTop(label, y - 7); gutter.Children.Add(label);
        }
        Grid.SetColumn(gutter, 0); root.Children.Add(gutter);
        Grid.SetColumn(ticks, 1); root.Children.Add(ticks);
        var marker = new Border { Height = 1.6, Background = Brush("#6FA98A"), VerticalAlignment = VerticalAlignment.Top, HorizontalAlignment = HorizontalAlignment.Stretch, CornerRadius = new CornerRadius(1) };
        var badge = new TextBlock { Text = "", FontSize = 10, Foreground = Brush("#2F4C40"), Background = Brush("#FCFEFC"), HorizontalAlignment = HorizontalAlignment.Left, VerticalAlignment = VerticalAlignment.Top, Padding = new Thickness(2, 0, 2, 0) };
        var overlay = new Grid { IsHitTestVisible = false }; overlay.Children.Add(marker); overlay.Children.Add(badge);
        Grid.SetColumn(overlay, 1); root.Children.Add(overlay);
        return (root, marker, badge);
    }
    private static int MinuteAt(double y, double height, int from, int to)
    {
        if (height <= 0) return from;
        var ratio = Math.Clamp(y / height, 0, 1);
        var minute = from + (int)Math.Round(ratio * (to - from) / 5.0) * 5;
        return Math.Clamp(minute, from, to);
    }
    /// <summary>三周内的大考：日程页顶部一条倒计时，流里对应的事也会标红。</summary>
    private UIElement? ExamBanner(AgendaDay day)
    {
        var upcoming = day.Exams.Take(3).ToList();
        if (upcoming.Count == 0) return null;
        var today = AgendaRules.ParseDate(day.Date);
        var wrap = new WrapPanel { Margin = new Thickness(0, 0, 0, 3) };
        foreach (var exam in upcoming)
        {
            var date = AgendaRules.ParseDate(exam.Date);
            var wait = date.DayNumber - today.DayNumber;
            var when = wait <= 0 ? "今天" : wait == 1 ? "明天" : wait + " 天后";
            var prefix = exam.Category.Length > 0 && !exam.Title.Contains(exam.Category) ? exam.Category + " " : "";
            var label = prefix + date.ToString("M.d") + " " + exam.Title + " · " + when;
            wrap.Children.Add(new Border { Background = Brush("#FBEFEA"), BorderBrush = Brush("#EBC7BA"), BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(7), Padding = new Thickness(7, 3, 7, 3), Margin = new Thickness(0, 0, 4, 4), ToolTip = exam.Time.Length > 0 ? exam.Time : null,
                Child = new TextBlock { Text = label, FontSize = 10.5, Foreground = Brush("#A9482D"), MaxWidth = 268, TextWrapping = TextWrapping.Wrap } });
        }
        return wrap;
    }
    /// <summary>容量提示：今天排了多少、还剩多少空档（借鉴 Super Productivity 的 overload 提醒）。</summary>
    private UIElement? DaySummary(List<(AgendaItem Item, (int Start, int End) Span)> entries)
    {
        if (entries.Count == 0) return null;
        var scheduled = entries.Sum(entry => entry.Span.End - entry.Span.Start);
        var gaps = 0;
        for (var index = 1; index < entries.Count; index++) gaps += Math.Max(0, entries[index].Span.Start - entries[index - 1].Span.End);
        var packed = scheduled >= 600;
        return new TextBlock { Text = "已排 " + AgendaRules.Human(scheduled) + " · 空档 " + AgendaRules.Human(gaps) + (packed ? " · 今天排得挺满" : ""), FontSize = 10.5, Foreground = Brush(packed ? "#B4482F" : "#9DB3A6"), Margin = new Thickness(2, 7, 0, 2) };
    }
    private Border ScheduleTail(DateOnly date, int lastEnd)
    {
        var full = lastEnd >= 24 * 60;
        var hint = full ? "放到这里 · 排进今天最早的空档" : "放到最后 · " + AgendaRules.MinuteText(lastEnd) + " 之后";
        var host = GhostDrop(hint, 12, 38, (source, args) =>
        {
            var length = (int)Math.Max(15, source.Planned ?? 60);
            status.Text = PlaceAt(source, date, full ? 8 * 60 : lastEnd, length);
            _ = AutoSyncAfterSchedule();
        });
        host.Margin = new Thickness(52, 2, 0, 2);
        return host;
    }
    /// <summary>日程条目当拖放目标：上半插到它前面，下半插到它后面，具体时间由 PlaceInDay 算。</summary>
    private void AttachFlowDrop(Border shell, AgendaItem target, DateOnly date)
    {
        var after = false;
        shell.DragOver += (_, eventArgs) =>
        {
            if (!eventArgs.Data.GetDataPresent("QingxuTask")) { eventArgs.Effects = DragDropEffects.None; eventArgs.Handled = true; return; }
            eventArgs.Effects = DragDropEffects.Move; eventArgs.Handled = true;
            after = eventArgs.GetPosition(shell).Y > shell.ActualHeight / 2;
            shell.BorderBrush = Brush(after ? "#B7D8C4" : "#9CC8AE");
            shell.BorderThickness = new Thickness(1, after ? 1 : 2, 1, after ? 2 : 1);
        };
        shell.DragLeave += (_, _) => { shell.BorderBrush = Brush("#E4EEE8"); shell.BorderThickness = new Thickness(1); };
        shell.Drop += (_, eventArgs) =>
        {
            eventArgs.Handled = true;
            shell.BorderBrush = Brush("#E4EEE8"); shell.BorderThickness = new Thickness(1);
            try
            {
                if (eventArgs.Data.GetData("QingxuTask") is not AgendaItem source || source.Id == target.Id) return;
                status.Text = PlaceInDay(source, date, target, eventArgs.GetPosition(shell).Y > shell.ActualHeight / 2);
                _ = AutoSyncAfterSchedule();
            }
            catch (Exception error) { status.Text = error.Message; }
        };
    }
    /// <summary>拉底边改时长：按 5 分钟一档，松手才写盘。</summary>
    private void AttachResize(Border handle, Border shell, AgendaItem item, (int Start, int End) span, TextBlock meta)
    {
        var original = (int)Math.Max(15, span.End - span.Start);
        var origin = meta.Text;
        Point? from = null; var preview = original;
        handle.MouseLeftButtonDown += (_, eventArgs) => { from = eventArgs.GetPosition(this); preview = original; handle.CaptureMouse(); eventArgs.Handled = true; };
        handle.MouseMove += (_, eventArgs) =>
        {
            if (from is not Point start) return;
            var steps = (int)Math.Round((eventArgs.GetPosition(this).Y - start.Y) / 2.0);
            preview = Math.Clamp(original + steps * 5, 15, 24 * 60 - span.Start);
            meta.Text = AgendaRules.MinuteText(span.Start) + "—" + AgendaRules.MinuteText(span.Start + preview) + " · " + AgendaRules.Human(preview);
            shell.MinHeight = Math.Clamp(preview * 0.45, 34, 160);
        };
        handle.MouseLeftButtonUp += (_, _) =>
        {
            handle.ReleaseMouseCapture(); if (from is null) return; from = null;
            if (preview == original) { meta.Text = origin; shell.MinHeight = Math.Clamp(original * 0.45, 34, 108); return; }
            try { Store.Save(item, new JsonObject { ["planned"] = preview }); Render(); status.Text = "时长改为 " + AgendaRules.Human(preview); }
            catch (Exception error) { status.Text = error.Message; meta.Text = origin; }
        };
    }
    /// <summary>撤销一条任务的时间排期，放回待办。</summary>
    private void UndoSchedule(AgendaItem item)
    {
        Run(() =>
        {
            if (item.Weekly || item.Semester) Store.Remove(item, false);
            else Store.Save(item, new JsonObject { ["start"] = "", ["planned"] = "" });
            Render();
            status.Text = "已撤销「" + item.DisplayTitle + "」的时间，回到待办";
        });
    }
    /// <summary>撤销「拖进课内任务」：课内任务还原，任务重新出现在待办里。</summary>
    private void UndoAbsorb(AgendaItem course)
    {
        Run(() =>
        {
            if (!absorbed.TryGetValue(course.Id, out var record) || record.Entries.Count == 0) { absorbed.Remove(course.Id); SaveSettings(); Render(); status.Text = "没有可以撤销的并入记录"; return; }
            var last = record.Entries[^1];
            if (!course.TaskTitle.Contains(last.Title)) { absorbed.Remove(course.Id); SaveSettings(); Render(); status.Text = "课内任务已经改过，这条并入记录作废"; return; }
            record.Entries.RemoveAt(record.Entries.Count - 1);
            Store.Save(course, new JsonObject { ["title"] = MergedTitle(record) });
            // 每日/本地任务被并进去时原条目已删除，这里补一条回来；周任务只是取消安排，留在清单里
            if (last.Id.StartsWith("local:") || last.Id.StartsWith("daily:"))
            {
                var task = new AgendaItem { Id = "local:" + Guid.NewGuid().ToString("N")[..12], Date = course.Date, Title = last.Title };
                Store.Save(task, new JsonObject { ["title"] = last.Title, ["order"] = 0 });
            }
            if (record.Entries.Count == 0) absorbed.Remove(course.Id);
            SaveSettings(); Render();
            status.Text = "已把「" + last.Title + "」放回待办，课内任务还原";
        });
    }
    /// <summary>课内任务 = 原来的内容 + 依次并进来的任务。</summary>
    private static string MergedTitle(AbsorbRecord record)
    {
        var parts = record.Entries.Select(entry => entry.Title).ToList();
        if (record.Previous.Length > 0) parts.Insert(0, record.Previous);
        return string.Join("；", parts);
    }
    /// <summary>把待办并进课程的课内任务：写进「课内安排」的课内任务列，原待办不再单独占一条。</summary>
    public string AbsorbIntoClass(AgendaItem source, AgendaItem course)
    {
        if (source.Class || source.Id.Length == 0) throw new InvalidOperationException("这一条不能放进课内任务。");
        if (source.Id == course.Id) return status.Text;
        if (!absorbed.TryGetValue(course.Id, out var record)) { record = new AbsorbRecord { Previous = course.TaskTitle }; absorbed[course.Id] = record; }
        record.Entries.Add(new AbsorbEntry { Title = source.DisplayTitle, Id = source.Id });
        Store.Save(course, new JsonObject { ["title"] = MergedTitle(record) });
        SaveSettings();
        var moved = source.Id.StartsWith("local:") || source.Id.StartsWith("daily:");
        if (moved) Store.Remove(source, false);
        Render();
        return moved ? "已并进「" + course.Title + "」的课内任务" : "已加进「" + course.Title + "」的课内任务，原任务还在清单里";
    }
    /// <summary>「自动排」：把待办按顺序填进今天的空档（借鉴 Motion / Sunsama 的 Plan my day）。</summary>
    private void AutoPlan()
    {
        if (page >= 3) { status.Text = "清单读取 Excel，请在日程页使用"; return; }
        var date = SelectedDate;
        var placed = 0;
        try
        {
            var day = Store.Day(date);
            var blocks = day.Schedule.Where(item => Range(item.Time) is not null).Select(item => (Item: item, Range: Span(item))).OrderBy(pair => pair.Range.Start).ToList();
            var plan = new List<(AgendaItem Item, int Start, int Length)>();
            foreach (var task in day.Tasks.Where(item => item.Editable).ToList())
            {
                var length = (int)Math.Max(15, task.Planned ?? 30);
                var start = FreeStart(blocks, 8 * 60, length);
                if (start is null || start.Value + length > 23 * 60) continue;
                blocks.Add((task, (start.Value, start.Value + length)));
                blocks = blocks.OrderBy(pair => pair.Range.Start).ToList();
                plan.Add((task, start.Value, length));
            }
            foreach (var entry in plan) Store.Schedule(entry.Item, date, AgendaRules.MinuteText(entry.Start), entry.Length);
            placed = plan.Count;
        }
        catch (Exception error) { status.Text = error.Message; return; }
        Render();
        status.Text = placed == 0 ? "今天排不下更多任务了" : "已把 " + placed + " 项待办排进空档";
    }
    private UIElement AutoPlanChip()
    {
        var chip = Button("自动排", (_, _) => AutoPlan());
        chip.FontSize = 10.5; chip.Padding = new Thickness(8, 2, 8, 2);
        chip.Background = Brush("#F1F6F3"); chip.Foreground = Brush("#5C8171");
        chip.ToolTip = "把待办按顺序填进今天的空档（从 08:00 开始，15 分钟起）";
        return chip;
    }
    /// <summary>把任务排到某天：贴着落点放，放不下就顺延到后面第一个装得下的空档。</summary>
    public string PlaceInDay(AgendaItem source, DateOnly date, AgendaItem? target, bool after)
    {
        if (source.Class || source.Semester) throw new InvalidOperationException("课程和学期任务不用排时间。");
        if (source.Id.Length == 0) throw new InvalidOperationException("固定安排不能拖动。");
        var day = currentDay ?? Store.Day(date);
        var blocks = day.Schedule.Where(item => item.Id != source.Id && Range(item.Time) is not null)
            .Select(item => (Item: item, Range: Span(item))).OrderBy(pair => pair.Range.Start).ToList();
        var index = blocks.Count;
        if (target is not null)
        {
            var found = target.Id.Length > 0
                ? blocks.FindIndex(pair => pair.Item.Id == target.Id)
                : blocks.FindIndex(pair => pair.Item.Time == target.Time && pair.Item.Title == target.Title);
            if (found >= 0) index = after ? found + 1 : found;
        }
        var duration = (int)Math.Max(15, source.Planned ?? 60);
        var low = index == 0 ? 7 * 60 : blocks[index - 1].Range.End;
        var high = index >= blocks.Count ? 24 * 60 : blocks[index].Range.Start;
        // 装得下就贴着落点放，装不下就占住落点、把后面的往后推
        var candidate = high - low >= duration ? (after ? low : high - duration) : low;
        return PlaceAt(source, date, candidate, duration);
    }
    /// <summary>按精确时间落位：占住这个时刻，后面的任务依次后推；推不下 24:00 就退回最近的空档。</summary>
    public string PlaceAt(AgendaItem source, DateOnly date, int start, int? duration = null)
    {
        if (source.Class || source.Semester) throw new InvalidOperationException("课程和学期任务不用排时间。");
        if (!source.Editable) throw new InvalidOperationException("固定安排不能拖动。");
        var day = currentDay is not null && currentDay.Date == AgendaRules.DateText(date) ? currentDay : Store.Day(date);
        var blocks = day.Schedule.Where(item => item.Id != source.Id && Range(item.Time) is not null)
            .Select(item => (Item: item, Range: Span(item))).OrderBy(pair => pair.Range.Start).ToList();
        var length = (int)Math.Max(15, duration ?? (source.Planned ?? 60));
        start = Math.Clamp(start, 0, Math.Max(0, 24 * 60 - length));
        var head = "已排 " + AgendaRules.MinuteText(start) + "—" + AgendaRules.MinuteText(start + length);
        var ordered = blocks.Select(pair => pair.Item).ToList();
        var index = ordered.FindIndex(item => Span(item).End > start);
        if (index < 0) index = ordered.Count;
        ordered.Insert(index, source);
        var plan = new List<(AgendaItem Item, int Start, int Length)>();
        var cursor = 0;
        foreach (var item in ordered)
        {
            var isSource = ReferenceEquals(item, source);
            var own = Span(item);
            var ownLength = isSource ? length : own.End - own.Start;
            var pinned = item.Class || !item.Editable;
            var begin = isSource ? start : pinned ? own.Start : Math.Max(own.Start, cursor);
            if (!pinned)
            {
                var pushed = PushPastFixed(blocks, begin, ownLength);
                if (pushed is null || pushed.Value + ownLength > 24 * 60) { plan.Clear(); break; }
                begin = pushed.Value;
            }
            plan.Add((item, begin, ownLength));
            cursor = Math.Max(cursor, begin + ownLength);
        }
        if (plan.Count == 0)
        {
            var free = FreeStart(blocks, start, length) ?? throw new InvalidOperationException("这一天排不下它了，先缩短时长或换一天。");
            Store.Schedule(source, date, AgendaRules.MinuteText(free), length);
            Render();
            return "已排 " + AgendaRules.MinuteText(free) + "—" + AgendaRules.MinuteText(free + length) + "（后面推不动了，改放最近空档）";
        }
        var moved = 0;
        for (var position = plan.Count - 1; position >= 0; position--)
        {
            var entry = plan[position];
            var isSource = ReferenceEquals(entry.Item, source);
            if (!isSource && entry.Start == Span(entry.Item).Start) continue;
            Store.Schedule(entry.Item, date, AgendaRules.MinuteText(entry.Start), isSource ? length : null);
            if (!isSource) moved++;
        }
        Render();
        return moved > 0 ? head + " · 顺延 " + moved + " 项" : head;
    }
    /// <summary>从 wanted 开始找第一个装得下 length 的空位。</summary>
    private static int? FreeStart(List<(AgendaItem Item, (int Start, int End) Range)> blocks, int wanted, int length)
    {
        var start = Math.Max(0, wanted);
        for (var guard = 0; guard <= blocks.Count + 1; guard++)
        {
            var hit = blocks.FirstOrDefault(pair => start < pair.Range.End && start + length > pair.Range.Start);
            if (hit.Item is null) return start + length > 24 * 60 ? null : start;
            start = hit.Range.End;
        }
        return null;
    }
    /// <summary>把起点推到所有固定块（课程、日程事项、习惯）之后。</summary>
    private static int? PushPastFixed(List<(AgendaItem Item, (int Start, int End) Range)> blocks, int begin, int length)
    {
        var start = Math.Max(0, begin);
        for (var guard = 0; guard <= blocks.Count; guard++)
        {
            var hit = blocks.FirstOrDefault(pair => (pair.Item.Class || !pair.Item.Editable) && start < pair.Range.End && start + length > pair.Range.Start);
            if (hit.Item is null) return start;
            start = hit.Range.End;
        }
        return null;
    }
    /// <summary>条目占用的时间：没有结束时间的（起床、单点事项）按 15 分钟算。</summary>
    private static (int Start, int End) Span(AgendaItem item)
    {
        var range = Range(item.Time);
        if (range is not { } value) return (0, 0);
        return value.End > value.Start ? value : (value.Start, value.Start + (int)Math.Max(15, item.Planned ?? 15));
    }
    private static (int Start, int End)? Range(string time)
    {
        var parts = time.Split('—');
        try
        {
            if (parts.Length == 1 && parts[0].Length > 0) { var single = AgendaRules.Minutes(parts[0]); return (single, single); }
            if (parts.Length != 2 || parts[0].Length == 0 || parts[1].Length == 0) return null;
            return (AgendaRules.Minutes(parts[0]), AgendaRules.Minutes(parts[1], true));
        }
        catch (InvalidOperationException) { return null; }
    }
    /// <summary>只有这一条改动时顺手写进 Excel；其它情况留给手动同步，避免把无关修改一起带走。</summary>
    private async Task AutoSyncAfterSchedule()
    {
        if (syncing || Store.State.Operations.Count != 1) return;
        var sent = Store.Snapshot(); var versions = sent.Select(operation => operation.Version).Distinct().ToArray();
        if (versions.Length != 1 || versions[0].Length == 0) return;
        syncing = true; syncButton.IsEnabled = false;
        foreach (var operation in sent) Store.SyncingIds.Add(operation.Id);
        var written = false;
        try
        {
            var result = await Task.Run(() => WorkbookSync.Write(Store.BookPath, versions[0], sent));
            Store.Complete(sent, result); written = true;
        }
        catch (Exception) { written = false; }
        finally { syncing = false; syncButton.IsEnabled = true; Store.SyncingIds.Clear(); }
        Render();
        status.Text = written ? "已排期并写入 Excel" : "已排期 · 待同步（关闭 Excel 后点「同步 Excel」）";
    }
    private async Task Synchronize()
    {
        if (syncing) return;
        if (Store.State.Operations.Count == 0) { Run(() => { Store.Reload(); Render(); }); return; }
        var sent = Store.Snapshot(); var versions = sent.Select(operation => operation.Version).Distinct().ToArray();
        syncing = true; syncButton.IsEnabled = false; status.Text = "正在后台同步 Excel…";
        foreach (var operation in sent) Store.SyncingIds.Add(operation.Id);
        var conflict = false;
        try
        {
            if (versions.Length != 1 || versions[0].Length == 0) throw new SyncConflictException();
            var result = await Task.Run(() => WorkbookSync.Write(Store.BookPath, versions[0], sent));
            Store.Complete(sent, result);
        }
        catch (SyncConflictException) { conflict = true; }
        catch (Exception error) { status.Text = "同步未完成 · 本地修改已保留"; MessageBox.Show(this, error.Message, "暂未同步", MessageBoxButton.OK, MessageBoxImage.Information); }
        finally { syncing = false; syncButton.IsEnabled = true; Store.SyncingIds.Clear(); }
        // 同步期间工作簿正被替换，此时重绘会读到中间态；统一在 syncing 复位之后再刷新界面
        Render();
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
