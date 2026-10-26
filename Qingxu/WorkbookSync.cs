using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text.Json.Nodes;
using System.Text.Json;

namespace Qingxu;

public sealed class SyncConflictException : IOException
{
    public SyncConflictException() : base("Excel 已有更新，尚未写入。请核对两边修改后再合并。") { }
}

public sealed record SyncResult(string Version, Dictionary<string, string> Mappings);
public sealed record SyncReceipt(List<PendingOperation> Sent, SyncResult Result);

public static class WorkbookSync
{
    public static void CheckAvailable(string path)
    {
        if (File.Exists(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(path)!, "~$" + System.IO.Path.GetFileName(path))))
            throw new IOException("请先保存并关闭 Excel／WPS，再同步。");
        using var check = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
    }
    public static SyncResult Write(string path, string version, IReadOnlyList<PendingOperation> operations)
    {
        CheckAvailable(path);
        using var guard = new FileStream(path + ".qingxu.lock", FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None, 1, FileOptions.DeleteOnClose);
        if (XlsxBook.CurrentVersion(path) != version || operations.Any(operation => operation.Version != version)) throw new SyncConflictException();
        var originalHash = Hash(path);
        var book = new XlsxBook(path, true);
        if (book.Version != version) throw new SyncConflictException();
        var mappings = new Dictionary<string, string>();
        foreach (var operation in operations) Apply(book, operation, mappings);
        var temporary = path + ".qingxu-pending-" + Guid.NewGuid().ToString("N") + ".xlsx";
        try
        {
            book.Export(temporary);
            CheckAvailable(path);
            var backup = path + ".qingxu-backup-" + DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) + "-" + Guid.NewGuid().ToString("N") + ".xlsx";
            using (var original = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read | FileShare.Delete))
            {
                var finalHash = Convert.ToHexString(SHA256.HashData(original));
                if (XlsxBook.CurrentVersion(path) != version || finalHash != originalHash) throw new SyncConflictException();
                File.Replace(temporary, path, backup);
            }
            var result = new SyncResult(XlsxBook.CurrentVersion(path), mappings);
            var receiptPath = path + ".qingxu-receipt.json";
            File.WriteAllText(receiptPath + ".pending", JsonSerializer.Serialize(new SyncReceipt(operations.Select(operation => operation.Copy()).ToList(), result), AgendaRules.JsonOptions));
            File.Move(receiptPath + ".pending", receiptPath, true);
            return result;
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }
    private static string Hash(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream));
    }
    private static void Apply(XlsxBook book, PendingOperation operation, Dictionary<string, string> mappings)
    {
        AgendaRules.CheckDate(operation.Date);
        var parts = operation.Id.Split(':');
        var type = parts[0];
        var fields = operation.Fields;
        string sheet;
        int row;
        var columns = new Dictionary<string, string>();
        switch (type)
        {
            case "daily": case "local": case "new":
                sheet = "学习记录";
                row = type == "daily" ? int.Parse(parts[1], CultureInfo.InvariantCulture) : 0;
                columns = new() { ["title"] = "D", ["subject"] = "C", ["planned"] = "E", ["actual"] = "F", ["done"] = "G", ["highlight"] = "K", ["start"] = "L", ["order"] = "B" };
                if (row == 0)
                {
                    var serial = AgendaReader.Serial(AgendaRules.ParseDate(operation.Date));
                    row = Enumerable.Range(15, 585).FirstOrDefault(candidate => book.Number(sheet, "A" + candidate) == serial && book.Text(sheet, "D" + candidate).Length == 0);
                    if (row == 0) row = Enumerable.Range(288, 312).FirstOrDefault(candidate => book.Text(sheet, "D" + candidate).Length == 0);
                    if (row == 0) throw new InvalidOperationException("学习记录已满，请先整理。");
                    book.Set(sheet, "A" + row, serial);
                    book.Set(sheet, "B" + row, Enumerable.Range(15, 585).Count(candidate => book.Number(sheet, "A" + candidate) == serial && book.Text(sheet, "D" + candidate).Length > 0) + 1);
                    mappings[operation.Id] = "daily:" + row;
                }
                break;
            case "class":
                sheet = "课内安排"; row = int.Parse(parts[1], CultureInfo.InvariantCulture);
                columns = new() { ["title"] = "F", ["subject"] = "G", ["planned"] = "H", ["actual"] = "I", ["done"] = "J", ["highlight"] = "M" }; break;
            case "weekly":
                sheet = parts[1]; row = int.Parse(parts[2], CultureInfo.InvariantCulture);
                columns = new() { ["title"] = "B", ["done"] = "E", ["highlight"] = "I", ["scheduleDate"] = "J", ["start"] = "K", ["planned"] = "L", ["order"] = "M" }; break;
            case "semester":
                sheet = "学期总清单"; row = int.Parse(parts[1], CultureInfo.InvariantCulture);
                columns = new() { ["title"] = "A", ["done"] = "B", ["highlight"] = "G" }; break;
            case "habit":
                sheet = parts[1]; row = int.Parse(parts[3], CultureInfo.InvariantCulture);
                if (parts[2].Length != 1 || parts[2][0] < 'B' || parts[2][0] > 'H' || row != 67 && row != 68) throw new InvalidOperationException("每日项目位置无效。");
                columns = new[] { "title", "done", "highlight", "start", "planned", "order" }.ToDictionary(key => key, _ => parts[2]); break;
            default: throw new InvalidOperationException("不支持的任务类型。");
        }
        if (row < 5 || row > 599 || type == "weekly" && (row < 50 || row > 62)) throw new InvalidOperationException("任务位置无效。");
        if (AgendaRules.Flag(fields, "delete"))
        {
            if (type == "habit")
            {
                book.Set(sheet, parts[2] + row, "已删除");
                foreach (var clearRow in new[] { row == 67 ? 69 : 72, row == 67 ? 73 : 74, row + 3, row + 8, row + 10 }) book.Set(sheet, parts[2] + clearRow, null);
            }
            else
            {
                var clearColumns = type switch { "class" => new[] { "F", "G", "H", "I", "J", "M" }, "weekly" => "BCDEFGHIJKLM".Select(character => character.ToString()).ToArray(), "semester" => new[] { "A", "B", "D", "G" }, _ => "ABCDEFGHIJKL".Select(character => character.ToString()).ToArray() };
                foreach (var column in clearColumns) book.Set(sheet, column + row, null);
            }
            return;
        }
        if (fields.ContainsKey("title") && string.IsNullOrWhiteSpace(AgendaRules.Text(fields, "title"))) throw new InvalidOperationException("请填写任务名称。");
        foreach (var pair in fields)
        {
            // incomplete 与 done 落在同一个状态列，最终值由两者共同决定，不受字段先后顺序影响
            var key = pair.Key == "incomplete" ? "done" : pair.Key;
            if (!columns.TryGetValue(key, out var column)) continue;
            var targetRow = row;
            object? value = pair.Value?.ToString();
            if (type == "habit") targetRow = key switch { "title" => row + 10, "order" => row + 8, "start" => row == 67 ? 69 : 72, "planned" => row == 67 ? 73 : 74, "highlight" => row + 3, _ => row };
            switch (key)
            {
                case "done":
                    value = AgendaRules.Flag(fields, "incomplete") ? "☒ 未完成" : AgendaRules.Flag(fields, "done") ? "☑ 已完成" : "☐ 待完成"; break;
                case "highlight": value = AgendaRules.Flag(fields, key) ? "强调" : "普通"; break;
                case "start": value = string.IsNullOrEmpty(value as string) ? null : AgendaRules.Minutes((string)value!) / 1440.0; break;
                case "order":
                    value = AgendaRules.Number(fields, key);
                    if (value is not double order || !double.IsFinite(order) || order < 1 || order % 1 != 0) throw new InvalidOperationException("任务顺序无效。"); break;
                case "planned": case "actual":
                    value = string.IsNullOrEmpty(value as string) ? null : AgendaRules.Number(fields, key) ?? throw new InvalidOperationException("分钟数无效。");
                    if (value is double duration && (!double.IsFinite(duration) || duration < 0 || duration > 1440)) throw new InvalidOperationException("分钟须在0—1440之间。"); break;
                case "scheduleDate":
                    value = string.IsNullOrEmpty(value as string) ? null : AgendaReader.Serial(AgendaRules.ParseDate((string)value!));
                    if (value is int assigned && (book.Number(sheet, "B5") is not double weekStart || assigned < weekStart || assigned > weekStart + 6)) throw new InvalidOperationException("周任务只能安排在所属周内。");
                    book.Set(sheet, "F" + row, null);
                    break;
            }
            book.Set(sheet, column + targetRow, value);
        }
    }
}
