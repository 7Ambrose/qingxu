using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Qingxu;

public sealed class AgendaItem
{
    public string Id { get; set; } = "";
    public string Date { get; set; } = "";
    public string Title { get; set; } = "";
    public string TaskTitle { get; set; } = "";
    public string Kind { get; set; } = "";
    public string Time { get; set; } = "";
    public string Start { get; set; } = "";
    public string PlannedDate { get; set; } = "";
    public string WeekStart { get; set; } = "";
    public string Subject { get; set; } = "";
    public double? Planned { get; set; }
    public double? Actual { get; set; }
    public double Order { get; set; }
    public bool Done { get; set; }
    public bool Incomplete { get; set; }
    public bool Highlight { get; set; }
    public bool Weekly => Id.StartsWith("weekly:", StringComparison.Ordinal);
    public bool Class => Id.StartsWith("class:", StringComparison.Ordinal);
    public bool Editable => Id.Length > 0;
    public string DisplayTitle => Class ? TaskTitle : Title;
    public AgendaItem Copy() => (AgendaItem)MemberwiseClone();
    public string Detail => Start.Length > 0 ? Start + "—" + AgendaRules.EndTime(Start, Planned ?? 60)
        : Weekly ? PlannedDate.Length > 0 ? "安排：" + PlannedDate : "本周未安排" : "";
}

public sealed class AgendaDay
{
    public string Date { get; set; } = "";
    public string WeekName { get; set; } = "";
    public string WeekStart { get; set; } = "";
    public List<AgendaItem> Schedule { get; set; } = new();
    public List<AgendaItem> Tasks { get; set; } = new();
    public List<AgendaItem> Weekly { get; set; } = new();
    public List<AgendaItem> WeekCourses { get; set; } = new();
}

public sealed class PendingOperation
{
    public string Id { get; set; } = "";
    public string Date { get; set; } = "";
    public string Version { get; set; } = "";
    public JsonObject Fields { get; set; } = new();
    public string Signature => Id + "|" + Date + "|" + string.Join("|", Fields.OrderBy(pair => pair.Key).Select(pair => pair.Key + ":" + pair.Value?.ToJsonString()));
    public PendingOperation Copy() => new() { Id = Id, Date = Date, Version = Version, Fields = (JsonObject)Fields.DeepClone() };
}

public sealed class Journal
{
    public int Schema { get; set; } = 1;
    public List<PendingOperation> Operations { get; set; } = new();
}

public static class AgendaRules
{
    public static readonly DateOnly FirstDate = new(2026, 10, 18);
    public static readonly DateOnly LastDate = new(2027, 1, 17);
    public static DateOnly Today => DateOnly.FromDateTime(TimeZoneInfo.ConvertTimeBySystemTimeZoneId(DateTime.UtcNow, "China Standard Time"));
    public static DateOnly ParseDate(string text) => DateOnly.ParseExact(text, "yyyy-MM-dd", CultureInfo.InvariantCulture);
    public static string DateText(DateOnly date) => date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
    public static void CheckDate(string date)
    {
        var parsed = ParseDate(date);
        if (parsed < FirstDate || parsed > LastDate) throw new InvalidOperationException("日期须在2026年10月18日至2027年1月17日内。");
    }
    public static int Minutes(string time, bool midnight = false)
    {
        if (midnight && time == "24:00") return 1440;
        if (!TimeOnly.TryParseExact(time, new[] { "H:mm", "HH:mm" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed))
            throw new InvalidOperationException("时间请填写 HH:mm，例如 13:30。");
        return parsed.Hour * 60 + parsed.Minute;
    }
    public static string MinuteText(int minutes) => $"{minutes / 60:00}:{minutes % 60:00}";
    public static string EndTime(string start, double duration) => MinuteText(Minutes(start) + (int)Math.Max(1, duration));
    public static string Text(JsonObject fields, string name) => fields[name]?.ToString() ?? "";
    public static double? Number(JsonObject fields, string name) => double.TryParse(Text(fields, name), NumberStyles.Float, CultureInfo.InvariantCulture, out var number) ? number : null;
    public static bool Flag(JsonObject fields, string name) => bool.TryParse(Text(fields, name), out var value) && value;
    public static void Validate(AgendaItem item, JsonObject fields)
    {
        CheckDate(item.Date);
        if (fields.ContainsKey("title") && string.IsNullOrWhiteSpace(Text(fields, "title"))) throw new InvalidOperationException("请填写任务名称。");
        foreach (var name in new[] { "planned", "actual" })
            if (Text(fields, name).Length > 0 && (Number(fields, name) is not double value || !double.IsFinite(value) || value < 0 || value > 1440))
                throw new InvalidOperationException("分钟须在0—1440之间。");
        var start = fields.ContainsKey("start") ? Text(fields, "start") : item.Start;
        var duration = fields.ContainsKey("planned") ? Number(fields, "planned") : item.Planned;
        if (start.Length > 0)
        {
            var minute = Minutes(start);
            if (minute + Math.Max(1, duration ?? 60) > 1440) throw new InvalidOperationException("结束时间不能超过24:00。");
            if (fields.ContainsKey("start")) fields["start"] = MinuteText(minute);
        }
        if (item.Weekly && Text(fields, "scheduleDate").Length > 0)
        {
            var assigned = ParseDate(Text(fields, "scheduleDate"));
            var week = ParseDate(item.WeekStart);
            if (assigned < week || assigned > week.AddDays(6)) throw new InvalidOperationException("周任务只能安排在所属周内。");
        }
    }
    public static void Apply(AgendaItem item, JsonObject fields)
    {
        foreach (var field in fields)
        {
            switch (field.Key)
            {
                case "title": if (item.Class) item.TaskTitle = Text(fields, "title"); else item.Title = Text(fields, "title"); break;
                case "start": item.Start = Text(fields, "start"); break;
                case "scheduleDate": item.PlannedDate = Text(fields, "scheduleDate"); break;
                case "planned": item.Planned = Number(fields, "planned"); break;
                case "actual": item.Actual = Number(fields, "actual"); break;
                case "order": item.Order = Number(fields, "order") ?? item.Order; break;
                case "done": item.Done = Flag(fields, "done"); if (item.Done) item.Incomplete = false; break;
                case "incomplete": item.Incomplete = Flag(fields, "incomplete"); if (item.Incomplete) item.Done = false; break;
                case "highlight": item.Highlight = Flag(fields, "highlight"); break;
                case "subject": item.Subject = Text(fields, "subject"); break;
            }
        }
        if (item.Kind == "task" || !item.Class && item.Editable) item.Time = item.Start.Length > 0 ? item.Detail : "";
    }
    public static readonly JsonSerializerOptions JsonOptions = new() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase, PropertyNameCaseInsensitive = true, WriteIndented = true, MaxDepth = 32 };
}
