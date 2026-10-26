using System;
using System.Globalization;
using System.Text.Json.Nodes;
using System.Windows;
using System.Windows.Controls;

namespace Qingxu;

public sealed class TaskEditor : Window
{
    public JsonObject? Result { get; private set; }
    public TaskEditor(Window owner, AgendaItem item)
    {
        Owner = owner; Title = item.Id.StartsWith("local:") && item.Title.Length == 0 ? "新建任务" : "编辑任务";
        Width = 360; SizeToContent = SizeToContent.Height; ResizeMode = ResizeMode.NoResize; WindowStartupLocation = WindowStartupLocation.CenterOwner;
        Background = MainWindow.Brush("#FCFEFC"); FontFamily = owner.FontFamily;
        var panel = new StackPanel { Margin = new Thickness(20) };
        var title = Field(panel, "任务名称", item.DisplayTitle);
        TextBox? start = null, end = null, minutes = null, subject = null, actual = null;
        DatePicker? date = null;
        if (!item.Class && !item.Id.StartsWith("semester:"))
        {
            start = Field(panel, "开始时间（可直接输入，如 13:30；留空取消时间）", item.Start);
            end = Field(panel, "结束时间（可留空，或填写24:00）", item.Start.Length > 0 ? AgendaRules.EndTime(item.Start, item.Planned ?? 60) : "");
            minutes = Field(panel, "预计时长（分钟）", item.Planned?.ToString(CultureInfo.InvariantCulture) ?? "");
        }
        if (item.Weekly)
        {
            panel.Children.Add(new TextBlock { Text = "安排日期（清空可取消安排）", Margin = new Thickness(0, 10, 0, 4) });
            date = new DatePicker { DisplayDateStart = AgendaRules.ParseDate(item.WeekStart).ToDateTime(TimeOnly.MinValue), DisplayDateEnd = AgendaRules.ParseDate(item.WeekStart).AddDays(6).ToDateTime(TimeOnly.MinValue) };
            if (item.PlannedDate.Length > 0) date.SelectedDate = AgendaRules.ParseDate(item.PlannedDate).ToDateTime(TimeOnly.MinValue);
            panel.Children.Add(date);
        }
        if (item.Class || item.Id.StartsWith("daily:") || item.Id.StartsWith("local:"))
        {
            if (item.Class) minutes = Field(panel, "预计时长（分钟，选填）", item.Planned?.ToString(CultureInfo.InvariantCulture) ?? "");
            subject = Field(panel, "科目（选填）", item.Subject);
            actual = Field(panel, "实际用时（分钟，选填）", item.Actual?.ToString(CultureInfo.InvariantCulture) ?? "");
        }
        var done = new CheckBox { Content = "已完成", IsChecked = item.Done, Margin = new Thickness(0, 12, 0, 6) };
        var highlight = new CheckBox { Content = "强调", IsChecked = item.Highlight, Margin = new Thickness(0, 0, 0, 12) };
        panel.Children.Add(done); panel.Children.Add(highlight);
        var save = MainWindow.Button("保存", (_, _) =>
        {
            try
            {
                var fields = new JsonObject { ["title"] = title.Text.Trim(), ["done"] = done.IsChecked == true, ["highlight"] = highlight.IsChecked == true };
                if (start is not null)
                {
                    fields["start"] = start.Text.Trim(); fields["planned"] = minutes!.Text.Trim();
                    if (start.Text.Trim().Length > 0 && end!.Text.Trim().Length > 0)
                    {
                        var duration = AgendaRules.Minutes(end.Text.Trim(), true) - AgendaRules.Minutes(start.Text.Trim());
                        if (duration <= 0) throw new InvalidOperationException("结束时间必须晚于开始时间。");
                        fields["planned"] = duration;
                    }
                }
                if (date is not null) fields["scheduleDate"] = date.SelectedDate?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) ?? "";
                if (item.Class && minutes is not null) fields["planned"] = minutes.Text.Trim();
                if (subject is not null) fields["subject"] = subject.Text.Trim();
                if (actual is not null) fields["actual"] = actual.Text.Trim();
                AgendaRules.Validate(item, fields); Result = fields; DialogResult = true;
            }
            catch (Exception error) { MessageBox.Show(this, error.Message, "请检查输入", MessageBoxButton.OK, MessageBoxImage.Information); }
        });
        panel.Children.Add(save); Content = panel;
    }
    private static TextBox Field(Panel panel, string label, string value)
    {
        panel.Children.Add(new TextBlock { Text = label, Margin = new Thickness(0, 8, 0, 4), Foreground = MainWindow.Brush("#658477"), TextWrapping = TextWrapping.Wrap });
        var input = new TextBox { Text = value, Padding = new Thickness(7), BorderBrush = MainWindow.Brush("#DCE9E1") };
        panel.Children.Add(input); return input;
    }
}
