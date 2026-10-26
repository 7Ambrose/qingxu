using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Nodes;
using System.Windows;
using System.Windows.Controls;

namespace Qingxu;

public sealed class ConflictReview : Window
{
    public string Version { get; }
    public List<PendingOperation> Operations { get; }
    public HashSet<string> KeepLocal { get; } = new();
    public ConflictReview(Window owner, AgendaStore store)
    {
        Owner = owner; Title = "核对 Excel 与小程序的修改"; Width = 700; Height = 600; WindowStartupLocation = WindowStartupLocation.CenterOwner;
        FontFamily = owner.FontFamily; Background = MainWindow.Brush("#FCFEFC");
        WorkbookSync.CheckAvailable(store.BookPath);
        var current = new XlsxBook(store.BookPath); Version = current.Version; Operations = store.Snapshot();
        var root = new DockPanel { Margin = new Thickness(18) };
        var hint = new TextBlock { Text = "Excel 已有更新，程序没有覆盖它。逐项勾选采用小程序修改，未勾选保留 Excel。\n无法自动判断哪边更新；确认前会备份两边，取消不改数据。", TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 0, 0, 12), Foreground = MainWindow.Brush("#547562") };
        DockPanel.SetDock(hint, Dock.Top); root.Children.Add(hint);
        var actions = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        actions.Children.Add(MainWindow.Button("取消，保留现状", (_, _) => DialogResult = false));
        var choices = new List<(CheckBox Choice, PendingOperation Operation)>();
        actions.Children.Add(MainWindow.Button("确认选择并继续同步", (_, _) => { foreach (var pair in choices.Where(pair => pair.Choice.IsChecked == true && pair.Choice.IsEnabled)) KeepLocal.Add(pair.Operation.Signature); DialogResult = true; }));
        DockPanel.SetDock(actions, Dock.Bottom); root.Children.Add(actions);
        var list = new StackPanel();
        foreach (var operation in Operations)
        {
            var day = AgendaReader.Read(current, AgendaRules.ParseDate(operation.Date));
            var item = day.Schedule.Concat(day.Tasks).Concat(day.Weekly).FirstOrDefault(candidate => candidate.Id == operation.Id);
            var newItem = operation.Id.StartsWith("local:");
            var title = item?.DisplayTitle;
            if (string.IsNullOrEmpty(title)) title = AgendaRules.Text(operation.Fields, "title");
            if (title.Length == 0) title = item?.Title ?? operation.Id;
            var card = new StackPanel { Margin = new Thickness(0, 4, 8, 16) };
            card.Children.Add(new TextBlock { Text = operation.Date + " · " + title, FontWeight = FontWeights.SemiBold, TextWrapping = TextWrapping.Wrap });
            var excel = item is null ? newItem ? "尚未创建" : "任务已删除或位置已改变，请保留 Excel 版本。" : Describe(Current(item, operation.Fields));
            card.Children.Add(new TextBlock { Text = "Excel 当前内容\n" + excel + "\n\n小程序待同步内容\n" + Describe(operation.Fields), TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 6, 0, 8) });
            var choice = new CheckBox { Content = "采用这项小程序修改", IsChecked = false, IsEnabled = item is not null || newItem, Foreground = MainWindow.Brush("#43765C") };
            choices.Add((choice, operation)); card.Children.Add(choice); list.Children.Add(card);
        }
        root.Children.Add(new ScrollViewer { Content = list, VerticalScrollBarVisibility = ScrollBarVisibility.Auto }); Content = root;
        if (XlsxBook.CurrentVersion(store.BookPath) != Version) throw new SyncConflictException();
    }
    private static JsonObject Current(AgendaItem item, JsonObject fields)
    {
        var values = new JsonObject { ["title"] = item.DisplayTitle, ["done"] = item.Done, ["incomplete"] = item.Incomplete, ["highlight"] = item.Highlight, ["start"] = item.Start, ["planned"] = item.Planned, ["actual"] = item.Actual, ["subject"] = item.Subject, ["scheduleDate"] = item.PlannedDate, ["order"] = item.Order };
        if (AgendaRules.Flag(fields, "delete")) return values;
        return new JsonObject(fields.Where(pair => values.ContainsKey(pair.Key)).Select(pair => new KeyValuePair<string, JsonNode?>(pair.Key, values[pair.Key]?.DeepClone())));
    }
    private static string Describe(JsonObject fields)
    {
        var labels = new Dictionary<string, string> { ["title"] = "名称", ["done"] = "完成", ["incomplete"] = "未完成", ["highlight"] = "强调", ["start"] = "开始时间", ["planned"] = "预计分钟", ["actual"] = "实际分钟", ["subject"] = "科目", ["scheduleDate"] = "安排日期", ["order"] = "顺序", ["delete"] = "删除" };
        return string.Join("\n", fields.Where(pair => labels.ContainsKey(pair.Key)).Select(pair => labels[pair.Key] + "：" + (pair.Value is null || pair.Value.ToString().Length == 0 ? "未填写" : pair.Value.ToString() == "true" ? "是" : pair.Value.ToString() == "false" ? "否" : pair.Value.ToString())));
    }
}
