using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Qingxu;

public sealed class AgendaStore
{
    private const long MaxJournalBytes = 16L * 1024 * 1024;
    public string BookPath { get; }
    public string JournalPath => BookPath + ".qingxu.json";
    public XlsxBook Book { get; private set; }
    public Journal State { get; private set; } = new();
    public bool ImportedLegacy { get; }
    public HashSet<string> SyncingIds { get; } = new();
    public AgendaStore(string path)
    {
        BookPath = Path.GetFullPath(path); Book = new XlsxBook(BookPath);
        if (File.Exists(JournalPath))
        {
            if (new FileInfo(JournalPath).Length > MaxJournalBytes) throw new IOException("本地日程文件异常（超过 16MB），请保留文件并检查。");
            State = JsonSerializer.Deserialize<Journal>(File.ReadAllText(JournalPath), AgendaRules.JsonOptions) ?? throw new IOException("本地日程文件无法读取，请保留文件并检查。");
        }
        else if (File.Exists(BookPath + ".agenda.json"))
        {
            var legacy = JsonNode.Parse(File.ReadAllText(BookPath + ".agenda.json"))!.AsObject();
            foreach (var node in legacy["pendingOps"]?.AsArray() ?? new JsonArray())
            {
                if (node is not JsonObject operation) continue;
                var fields = (JsonObject)(operation["fields"]?.DeepClone() ?? new JsonObject());
                var id = operation["id"]?.ToString() ?? "";
                if (id.StartsWith("weekly:") && AgendaRules.Flag(fields, "delete") && !fields.ContainsKey("deleteScope") && operation["snapshot"]?["plannedDate"]?.ToString().Length > 0)
                {
                    fields.Remove("delete"); fields["scheduleDate"] = ""; fields["start"] = "";
                    fields["title"] = operation["snapshot"]?["title"]?.ToString();
                }
                State.Operations.Add(new PendingOperation { Id = id, Date = operation["date"]?.ToString() ?? "", Fields = fields,
                    Version = operation["version"]?.ToString() ?? legacy["version"]?.ToString() ?? "" });
            }
            ImportedLegacy = true; Persist();
        }
        if (State.Schema != 1) throw new IOException("本地日程格式不受当前版本支持，请保留文件并使用对应版本打开。");
        var receiptPath = BookPath + ".qingxu-receipt.json";
        if (File.Exists(receiptPath))
        {
            var receipt = JsonSerializer.Deserialize<SyncReceipt>(File.ReadAllText(receiptPath), AgendaRules.JsonOptions) ?? throw new IOException("同步回执无法读取，请保留文件并检查。");
            Complete(receipt.Sent, receipt.Result);
        }
    }
    public void Reload() => Book = new XlsxBook(BookPath);
    public void Persist()
    {
        var temporary = JournalPath + ".pending";
        File.WriteAllText(temporary, JsonSerializer.Serialize(State, AgendaRules.JsonOptions));
        File.Move(temporary, JournalPath, true);
    }
    public AgendaDay Day(DateOnly date, bool pending = true)
    {
        var result = AgendaReader.Read(Book, date);
        if (!pending) return result;
        foreach (var operation in State.Operations)
        {
            var isWeekly = operation.Id.StartsWith("weekly:");
            if (isWeekly && operation.Id.Split(':')[1] != result.WeekName || !isWeekly && operation.Date != result.Date) continue;
            var item = result.Schedule.Concat(result.Tasks).Concat(result.Weekly).FirstOrDefault(candidate => candidate.Id == operation.Id)?.Copy();
            if (item is null && operation.Id.StartsWith("local:")) item = new AgendaItem { Id = operation.Id, Date = operation.Date };
            if (item is null) continue;
            result.Schedule.RemoveAll(candidate => candidate.Id == operation.Id); result.Tasks.RemoveAll(candidate => candidate.Id == operation.Id);
            if (AgendaRules.Flag(operation.Fields, "delete"))
            {
                if (item.Class) { item.TaskTitle = ""; item.Done = false; item.Highlight = false; item.Planned = null; item.Actual = null; result.Schedule.Add(item); }
                else result.Weekly.RemoveAll(candidate => candidate.Id == operation.Id);
                continue;
            }
            AgendaRules.Apply(item, operation.Fields);
            if (isWeekly)
            {
                result.Weekly.RemoveAll(candidate => candidate.Id == operation.Id); result.Weekly.Add(item.Copy());
                if (item.PlannedDate == result.Date) AgendaReader.Add(result, item);
            }
            else if (item.Class) result.Schedule.Add(item);
            else AgendaReader.Add(result, item);
        }
        AgendaReader.Sort(result); return result;
    }
    public void Save(AgendaItem item, JsonObject fields)
    {
        AgendaRules.Validate(item, fields);
        var date = item.Weekly ? item.WeekStart : item.Date;
        var copy = State.Operations.Select(operation => operation.Copy()).ToList();
        var existing = State.Operations.FirstOrDefault(operation => operation.Id == item.Id && operation.Date == date);
        if (existing is null) { existing = new PendingOperation { Id = item.Id, Date = date, Version = Book.Version }; State.Operations.Add(existing); }
        foreach (var field in fields) existing.Fields[field.Key] = field.Value?.DeepClone();
        try { Persist(); } catch { State.Operations = copy; throw; }
    }
    public void Remove(AgendaItem item, bool weeklyPage)
    {
        if (item.Weekly && !weeklyPage) { Save(item, new JsonObject { ["scheduleDate"] = "", ["start"] = "" }); return; }
        if (item.Id.StartsWith("local:") && !SyncingIds.Contains(item.Id))
        {
            var original = Snapshot(); State.Operations.RemoveAll(operation => operation.Id == item.Id);
            try { Persist(); } catch { State.Operations = original; throw; }
            return;
        }
        var fields = new JsonObject { ["delete"] = true };
        if (item.Weekly) fields["deleteScope"] = "weekly";
        Save(item, fields);
    }
    public void Schedule(AgendaItem item, DateOnly date, string? time)
    {
        if (item.Class || item.Id.StartsWith("semester:")) throw new InvalidOperationException("课程内任务沿用课程时间。");
        if (!item.Weekly && item.Date != AgendaRules.DateText(date)) throw new InvalidOperationException("每日任务只能在当天排期，周任务可安排到所属周的其他日期。");
        var fields = new JsonObject();
        if (item.Weekly) fields["scheduleDate"] = AgendaRules.DateText(date);
        var start = time ?? item.Start;
        if (start.Length > 0)
        {
            var duration = item.Planned ?? 60; var begin = AgendaRules.Minutes(start); var end = begin + Math.Max(1, duration);
            var conflict = Day(date).Schedule.Where(candidate => candidate.Id != item.Id).Any(candidate =>
            {
                var range = candidate.Time.Split('—');
                return range.Length == 2 && range[0].Length > 0 && range[1].Length > 0 && begin < AgendaRules.Minutes(range[1], true) && end > AgendaRules.Minutes(range[0]);
            });
            if (conflict) throw new InvalidOperationException("该时间段已有课程或任务，请换一个时间。");
            fields["start"] = start; fields["planned"] = duration;
        }
        Save(item, fields);
    }
    public List<PendingOperation> Snapshot() => State.Operations.Select(operation => operation.Copy()).ToList();
    public void Complete(IReadOnlyList<PendingOperation> sent, SyncResult result)
    {
        var original = Snapshot();
        var signatures = sent.Select(operation => operation.Signature).ToHashSet();
        State.Operations.RemoveAll(operation => signatures.Contains(operation.Signature));
        foreach (var operation in State.Operations)
        {
            if (result.Mappings.TryGetValue(operation.Id, out var mapped)) operation.Id = mapped;
            operation.Version = result.Version;
        }
        try { Persist(); } catch { State.Operations = original; throw; }
        var receiptPath = BookPath + ".qingxu-receipt.json";
        if (File.Exists(receiptPath)) File.Delete(receiptPath);
        Reload();
    }
    public void Resolve(string reviewedVersion, IReadOnlyList<PendingOperation> reviewed, IReadOnlySet<string> keepLocal)
    {
        WorkbookSync.CheckAvailable(BookPath);
        if (XlsxBook.CurrentVersion(BookPath) != reviewedVersion) throw new SyncConflictException();
        if (!State.Operations.Select(operation => operation.Signature).SequenceEqual(reviewed.Select(operation => operation.Signature))) throw new IOException("本地任务已变化，请重新核对。");
        var backup = JournalPath + ".before-merge-" + Guid.NewGuid().ToString("N") + ".json";
        File.Copy(JournalPath, backup);
        File.Copy(BookPath, BookPath + ".before-merge-" + Guid.NewGuid().ToString("N") + ".xlsx");
        if (XlsxBook.CurrentVersion(BookPath) != reviewedVersion) throw new SyncConflictException();
        var original = Snapshot();
        State.Operations = State.Operations.Where(operation => keepLocal.Contains(operation.Signature)).ToList();
        foreach (var operation in State.Operations) operation.Version = reviewedVersion;
        try { Persist(); } catch { State.Operations = original; throw; }
        Reload();
    }
}
