using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Qingxu;

public static class AgendaReader
{
    private static readonly DateOnly Epoch = new(1899, 12, 30);
    public static int Serial(DateOnly date) => date.DayNumber - Epoch.DayNumber;
    public static string SerialDate(double? serial) => serial is double number ? AgendaRules.DateText(Epoch.AddDays((int)number)) : "";
    public static string Time(double? fraction) => fraction is double number ? AgendaRules.MinuteText((int)Math.Round(number * 1440)) : "";
    public static AgendaDay Read(XlsxBook book, DateOnly date)
    {
        var day = AgendaRules.DateText(date);
        var serial = Serial(date);
        var index = date.DayNumber - new DateOnly(2026, 10, 19).DayNumber;
        var week = index < 0 ? "秋6周" : index >= 91 ? "冬10周" : index / 7 < 3 ? $"秋{6 + index / 7}周" : $"冬{index / 7 - 2}周";
        // 一次开包读完全部需要的表，避免每张表重新打开文件与 ZIP 目录
        book.Prepare(week, "课内安排", "日程事项", "学习记录", "学期总清单");
        var hasWeek = index >= 0 && index < 91 && book.HasSheet(week);
        var start = SerialDate(book.Number(week, "B5") ?? 46314);
        var weekStartSerial = Serial(AgendaRules.ParseDate(start));
        var weekEndSerial = weekStartSerial + 6;
        var result = new AgendaDay { Date = day, WeekName = week, WeekStart = start };
        var column = ((char)('B' + ((int)date.DayOfWeek + 6) % 7)).ToString();
        if (hasWeek)
        {
            var wake = book.Text(week, column + "7");
            result.Schedule.Add(new AgendaItem { Date = day, Title = "起床", Time = wake[..Math.Min(5, wake.Length)], Kind = "fixed" });
            foreach (var row in new[] { 67, 68 })
            {
                if (book.Text(week, column + row).Contains("已删除")) continue;
                var item = new AgendaItem { Id = $"habit:{week}:{column}:{row}", Date = day,
                    Title = book.Text(week, column + (row + 10)), Start = Time(book.Number(week, column + (row == 67 ? 69 : 72))),
                    Planned = book.Number(week, column + (row == 67 ? 73 : 74)) ?? 30,
                    Done = book.Text(week, column + row).Contains("已完成"), Incomplete = book.Text(week, column + row).Contains("未完成"), Highlight = book.Text(week, column + (row + 3)) == "强调",
                    Order = book.Number(week, column + (row + 8)) ?? row - 69 };
                if (item.Title.Length == 0) item.Title = book.Text(week, "A" + row);
                if (item.Title.Length == 0) item.Title = row == 67 ? "背单词" : "跑步打卡";
                Add(result, item);
            }
        }
        for (var row = 6; row < 400; row++)
        {
            var courseSerial = book.Number("课内安排", "A" + row);
            if (courseSerial == serial)
            {
                result.Schedule.Add(new AgendaItem { Id = "class:" + row, Date = day, Kind = "class", Title = book.Text("课内安排", "C" + row),
                    TaskTitle = book.Text("课内安排", "F" + row), Time = Time(book.Number("课内安排", "D" + row)) + "—" + Time(book.Number("课内安排", "E" + row)),
                    Subject = book.Text("课内安排", "G" + row), Planned = book.Number("课内安排", "H" + row), Actual = book.Number("课内安排", "I" + row),
                    Done = book.Text("课内安排", "J" + row).Contains("已完成"), Incomplete = book.Text("课内安排", "J" + row).Contains("未完成"), Highlight = book.Text("课内安排", "M" + row) == "强调" });
            }
            if (courseSerial >= weekStartSerial && courseSerial <= weekEndSerial && book.Text("课内安排", "C" + row).Length > 0)
                result.WeekCourses.Add(new AgendaItem { Id = "class:" + row, Date = SerialDate(courseSerial), Title = book.Text("课内安排", "C" + row), Kind = "class",
                    Time = Time(book.Number("课内安排", "D" + row)) + "—" + Time(book.Number("课内安排", "E" + row)) });
        }
        if (hasWeek && date.DayOfWeek == DayOfWeek.Thursday) result.Schedule.Add(new AgendaItem { Date = day, Title = "部门例会", Time = "21:30—22:30" });
        for (var row = 5; row < 200; row++)
        {
            if (book.Number("日程事项", "A" + row) != serial) continue;
            var span = Time(book.Number("日程事项", "C" + row)) + "—" + Time(book.Number("日程事项", "D" + row));
            var existing = result.Schedule.FirstOrDefault(item => item.Time == span);
            if (existing is not null) existing.Title = book.Text("日程事项", "E" + row);
            else result.Schedule.Add(new AgendaItem { Date = day, Title = book.Text("日程事项", "E" + row), Time = span, Kind = "event" });
        }
        for (var row = 15; row < 600; row++)
        {
            if (book.Number("学习记录", "A" + row) != serial || book.Text("学习记录", "D" + row).Length == 0 || book.Text("学习记录", "G" + row) == "— 取消") continue;
            Add(result, new AgendaItem { Id = "daily:" + row, Date = day, Title = book.Text("学习记录", "D" + row), Subject = book.Text("学习记录", "C" + row),
                Planned = book.Number("学习记录", "E" + row), Actual = book.Number("学习记录", "F" + row), Start = Time(book.Number("学习记录", "L" + row)),
                Done = book.Text("学习记录", "G" + row).Contains("已完成"), Incomplete = book.Text("学习记录", "G" + row).Contains("未完成"), Highlight = book.Text("学习记录", "K" + row) == "强调", Order = book.Number("学习记录", "B" + row) ?? row });
        }
        for (var row = 50; row < 63; row++)
        {
            var title = book.Text(week, "B" + row);
            if (title.Length == 0) continue;
            var item = new AgendaItem { Id = $"weekly:{week}:{row}", Date = start, WeekStart = start, Title = title,
                PlannedDate = SerialDate(book.Number(week, "J" + row)), Start = Time(book.Number(week, "K" + row)), Planned = book.Number(week, "L" + row) ?? 60,
                Done = book.Text(week, "E" + row).Contains("已完成"), Incomplete = book.Text(week, "E" + row).Contains("未完成"), Highlight = book.Text(week, "I" + row) == "强调", Order = book.Number(week, "M" + row) ?? row };
            if (item.PlannedDate.Length == 0 && hasWeek && book.Number(week, "F" + row) is double legacyDate && book.Text(week, "E" + row) != "— 本周无任务") item.PlannedDate = SerialDate(legacyDate);
            result.Weekly.Add(item);
            if (item.PlannedDate == day) Add(result, item.Copy());
        }
        for (var row = 5; row < 8; row++)
            if (book.Number("学期总清单", "D" + row) == serial && book.Text("学期总清单", "A" + row).Length > 0)
                Add(result, new AgendaItem { Id = "semester:" + row, Date = day, Title = book.Text("学期总清单", "A" + row),
                    Done = book.Text("学期总清单", "B" + row).Contains("已完成"), Incomplete = book.Text("学期总清单", "B" + row).Contains("未完成"), Highlight = book.Text("学期总清单", "G" + row) == "强调" });
        Sort(result);
        return result;
    }
    public static void Add(AgendaDay day, AgendaItem item)
    {
        item.Date = day.Date;
        if (item.Start.Length > 0) { item.Time = item.Detail; item.Kind = "task"; day.Schedule.Add(item); }
        else { item.Time = ""; day.Tasks.Add(item); }
    }
    public static void Sort(AgendaDay day)
    {
        day.Schedule = day.Schedule.OrderBy(item => item.Time.Length == 0 ? "99:99" : item.Time, StringComparer.Ordinal).ToList();
        day.Tasks = day.Tasks.OrderBy(item => item.Order).ToList();
        day.Weekly = day.Weekly.OrderBy(item => item.Order).ToList();
        day.WeekCourses = day.WeekCourses.OrderBy(item => item.Date).ThenBy(item => item.Time).ToList();
    }
}
