using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Qingxu;

internal static class NativeTests
{
    private static int assertions;
    private static void Assert(bool condition, string message) { assertions++; if (!condition) throw new Exception(message); }
    private static PendingOperation Op(XlsxBook book, string id, string date, JsonObject fields) => new() { Id = id, Date = date, Version = book.Version, Fields = fields };
    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));
    [STAThread]
    private static int Main(string[] arguments)
    {
        if (arguments[0] == "--dump")
        {
            Console.WriteLine(JsonSerializer.Serialize(AgendaReader.Read(new XlsxBook(arguments[1]), AgendaRules.ParseDate(arguments[2])), AgendaRules.JsonOptions)); return 0;
        }
        if (arguments[0] == "--dump-range")
        {
            var book = new XlsxBook(arguments[1]);
            Console.WriteLine(JsonSerializer.Serialize(Enumerable.Range(0, AgendaRules.LastDate.DayNumber - AgendaRules.FirstDate.DayNumber + 1).Select(offset => AgendaReader.Read(book, AgendaRules.FirstDate.AddDays(offset))), AgendaRules.JsonOptions)); return 0;
        }
        var source = Path.GetFullPath(arguments[0]); var sourceHash = Hash(source);
        var fixture = Path.Combine(Path.GetDirectoryName(source)!, "qa", "native-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(fixture);
        var path = Path.Combine(fixture, "plan.xlsx"); File.Copy(source, path);
        try
        {
            var book = new XlsxBook(path); var original = new XlsxBook(source);
            Assert(book.SheetNames.SequenceEqual(original.SheetNames), "Sheet list changed");
            foreach (var meetingDate in new[] { new DateOnly(2026, 11, 5), new DateOnly(2026, 11, 12) })
                Assert(!AgendaReader.Read(book, meetingDate).Schedule.Any(item => item.Title == "部门例会"), "Cancelled meeting remains in the agenda");
            Assert(AgendaReader.Read(book, new DateOnly(2026, 11, 19)).Schedule.Any(item => item.Title == "部门例会"), "Cancelling two meetings removed later meetings");
            var beforeEntries = Entries(source);
            var created = WorkbookSync.Write(path, book.Version, new[] { Op(book, "local:test", "2026-10-20", new JsonObject { ["title"] = "Native task", ["start"] = "13:30", ["planned"] = 45, ["done"] = false, ["highlight"] = true, ["order"] = 1 }) });
            var identity = created.Mappings["local:test"]; book = new XlsxBook(path);
            var task = AgendaReader.Read(book, new DateOnly(2026, 10, 20)).Schedule.Single(item => item.Id == identity);
            Assert(task.Title == "Native task" && task.Detail == "13:30—14:15" && task.Highlight, "Native create/time/highlight failed");
            WorkbookSync.Write(path, book.Version, new[] { Op(book, identity, "2026-10-20", new JsonObject { ["start"] = "", ["done"] = true }) });
            book = new XlsxBook(path); task = AgendaReader.Read(book, new DateOnly(2026, 10, 20)).Tasks.Single(item => item.Id == identity);
            Assert(task.Start.Length == 0 && task.Detail.Length == 0 && task.Done, "Clearing time left a stale subtitle");
            WorkbookSync.Write(path, book.Version, new[] { Op(book, identity, "2026-10-20", new JsonObject { ["incomplete"] = true }) });
            book = new XlsxBook(path); task = AgendaReader.Read(book, new DateOnly(2026, 10, 20)).Tasks.Single(item => item.Id == identity);
            Assert(task.Incomplete && !task.Done, "Incomplete state did not round-trip through the workbook");
            WorkbookSync.Write(path, book.Version, new[] { Op(book, identity, "2026-10-20", new JsonObject { ["done"] = true, ["incomplete"] = false }) });
            book = new XlsxBook(path); task = AgendaReader.Read(book, new DateOnly(2026, 10, 20)).Tasks.Single(item => item.Id == identity);
            Assert(task.Done && !task.Incomplete, "Completing a task did not clear the incomplete mark");
            WorkbookSync.Write(path, book.Version, new[] {
                Op(book, "weekly:秋6周:50", "2026-10-19", new JsonObject { ["title"] = "Weekly test", ["scheduleDate"] = "2026-10-20", ["start"] = "15:30", ["planned"] = 30, ["done"] = true, ["order"] = 1 }),
                Op(book, "habit:秋6周:B:67", "2026-10-19", new JsonObject { ["done"] = false, ["title"] = "Monday words", ["start"] = "08:30", ["planned"] = 25 }),
                Op(book, "class:10", "2026-10-20", new JsonObject { ["title"] = "Assignment", ["done"] = true }) });
            book = new XlsxBook(path);
            Assert(AgendaReader.Read(book, new DateOnly(2026, 10, 20)).Schedule.Count(item => item.Id == "weekly:秋6周:50") == 1, "Weekly task duplicated");
            Assert(AgendaReader.Read(book, new DateOnly(2026, 10, 20)).Schedule.Single(item => item.Id == "class:10").Title == AgendaReader.Read(original, new DateOnly(2026, 10, 20)).Schedule.Single(item => item.Id == "class:10").Title, "Assignment edit changed the course title");
            // 老的课内任务可以清空：写入空标题应当把 F 列清掉
            WorkbookSync.Write(path, book.Version, new[] { Op(book, "class:10", "2026-10-20", new JsonObject { ["title"] = " " }) });
            book = new XlsxBook(path);
            Assert(string.IsNullOrEmpty(book.Text("课内安排", "F10")), "Clearing the course task left text in the cell");
            Assert(AgendaReader.Read(book, new DateOnly(2026, 10, 19)).Schedule.Single(item => item.Id == "habit:秋6周:B:67").Detail == "08:30—08:55", "Routine time failed");
            foreach (var overflowRow in new[] { 63, 66 })
            {
                WorkbookSync.Write(path, book.Version, new[] { Op(book, $"weekly:秋6周:{overflowRow}", "2026-10-19", new JsonObject { ["title"] = $"Overflow {overflowRow}", ["scheduleDate"] = "2026-10-20", ["start"] = "22:00", ["done"] = true }) });
                book = new XlsxBook(path);
                Assert(AgendaReader.Read(book, new DateOnly(2026, 10, 20)).Schedule.Any(item => item.Id == $"weekly:秋6周:{overflowRow}" && item.Done), "Overflow weekly task did not round-trip");
            }
            Assert(AgendaReader.Read(book, new DateOnly(2026, 10, 20)).Weekly.All(item => !item.Id.EndsWith(":67")), "Routine row entered the weekly list");
            var store = new AgendaStore(path);
            var dragTask = new AgendaItem { Id = "local:drag-conflict", Date = "2026-10-20", Planned = 30 };
            var dragRejected = false;
            try { store.Schedule(dragTask, new DateOnly(2026, 10, 20), "13:30"); } catch (InvalidOperationException) { dragRejected = true; }
            Assert(dragRejected && !store.State.Operations.Any(operation => operation.Id == dragTask.Id), "13:30 course conflict reached the local queue");
            Assert(AgendaRules.Minutes("13:30") == 810 && AgendaRules.EndTime("23:30", 30) == "24:00", "Time grid or midnight boundary failed");
            Assert(AgendaRules.Minutes("１３：３０") == 810 && AgendaRules.Minutes("9：05") == 545 && AgendaRules.Minutes(" 09 : 05 ") == 545, "Full-width colon, digits or spaces were rejected");
            var weekly = store.Day(new DateOnly(2026, 10, 20)).Weekly.Single(item => item.Id == "weekly:秋6周:50");
            store.Remove(weekly, false);
            Assert(store.Day(new DateOnly(2026, 10, 20)).Weekly.Single(item => item.Id == weekly.Id).PlannedDate.Length == 0, "Daily removal deleted weekly task");
            Assert(!store.Day(new DateOnly(2026, 10, 20)).Schedule.Concat(store.Day(new DateOnly(2026, 10, 20)).Tasks).Any(item => item.Id == weekly.Id), "Unscheduling left a daily copy");
            var pending = store.Snapshot(); var saved = WorkbookSync.Write(path, store.Book.Version, pending); store.Complete(pending, saved);
            Assert(store.Day(new DateOnly(2026, 10, 19)).Weekly.Single(item => item.Id == weekly.Id).Done, "Unscheduling lost completion");
            Assert(store.Day(new DateOnly(2026, 10, 26)).Weekly.All(item => item.Id != weekly.Id), "Weekly state leaked into the next week");
            var local = new AgendaItem { Id = "local:in-flight", Date = "2026-10-21", Title = "In flight" };
            store.Save(local, new JsonObject { ["title"] = local.Title }); var sent = store.Snapshot(); store.SyncingIds.Add(local.Id);
            store.Remove(local, false); Assert(store.State.Operations.Any(operation => AgendaRules.Flag(operation.Fields, "delete")), "In-flight deletion lost its tombstone");
            var synced = WorkbookSync.Write(path, store.Book.Version, sent); store.Complete(sent, synced);
            Assert(store.State.Operations.Count == 1 && store.State.Operations[0].Id.StartsWith("daily:") && AgendaRules.Flag(store.State.Operations[0].Fields, "delete"), "In-flight deletion was not remapped");
            sent = store.Snapshot(); synced = WorkbookSync.Write(path, store.Book.Version, sent); store.Complete(sent, synced);
            Assert(!store.Day(new DateOnly(2026, 10, 21)).Tasks.Any(item => item.Title == local.Title), "In-flight deleted task was resurrected");
            var other = new AgendaItem { Id = "local:other-date", Date = "2026-10-22" };
            store.Save(other, new JsonObject { ["title"] = "Other date" });
            store.Save(new AgendaItem { Id = identity, Date = "2026-10-20" }, new JsonObject { ["order"] = 2 });
            var restarted = new AgendaStore(path);
            Assert(restarted.Day(new DateOnly(2026, 10, 22)).Tasks.Any(item => item.Id == other.Id), "Date-specific reorder lost another day's task");
            var crashSent = restarted.Snapshot(); WorkbookSync.Write(path, restarted.Book.Version, crashSent);
            restarted = new AgendaStore(path);
            Assert(restarted.State.Operations.Count == 0 && restarted.Day(new DateOnly(2026, 10, 22)).Tasks.Count(item => item.Title == "Other date") == 1, "Recovery after Excel write duplicated an already-synced local task");
            restarted.Save(new AgendaItem { Id = identity, Date = "2026-10-20" }, new JsonObject { ["done"] = false });
            var review = restarted.Snapshot(); var currentVersion = restarted.Book.Version;
            File.SetLastWriteTimeUtc(path, DateTime.UtcNow.AddMinutes(1)); var unchanged = Hash(path);
            var rejected = false; try { WorkbookSync.Write(path, currentVersion, review); } catch (SyncConflictException) { rejected = true; }
            Assert(rejected && Hash(path) == unchanged, "Stale version overwrote Excel");
            rejected = false; try { restarted.Resolve(currentVersion, review, new HashSet<string>()); } catch (SyncConflictException) { rejected = true; }
            Assert(rejected && restarted.State.Operations.Count == review.Count, "Stale conflict review discarded operations");
            using (var locked = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                rejected = false; try { WorkbookSync.CheckAvailable(path); } catch (IOException) { rejected = true; }
                Assert(rejected, "Locked file was writable");
            }
            var afterEntries = Entries(path);
            foreach (var entry in beforeEntries.Where(pair => !pair.Key.StartsWith("xl/worksheets/") && pair.Key != "xl/workbook.xml"))
                Assert(afterEntries.TryGetValue(entry.Key, out var hash) && hash == entry.Value, "Unrelated workbook part changed: " + entry.Key);
            Assert(beforeEntries.Keys.ToHashSet().SetEquals(afterEntries.Keys), "Workbook package parts were added/lost");
            TestLegacyImport(source, fixture);
            var semesterPath = Path.Combine(fixture, "semester.xlsx"); File.Copy(source, semesterPath);
            var semesterStore = new AgendaStore(semesterPath);
            var termItem = semesterStore.Day(new DateOnly(2026, 11, 9)).Semester.First();
            semesterStore.Save(termItem, new JsonObject { ["done"] = true, ["order"] = 1 });
            Assert(semesterStore.Day(new DateOnly(2026, 12, 7)).Semester.Single(item => item.Id == termItem.Id).Done, "Semester completion was restricted to one date");
            semesterStore.Save(semesterStore.Day(new DateOnly(2026, 12, 7)).Semester.Single(item => item.Id == termItem.Id), new JsonObject { ["done"] = false, ["incomplete"] = true });
            Assert(semesterStore.State.Operations.Count == 1, "Semester edits on different dates created duplicate operations");
            Assert(semesterStore.Day(AgendaRules.LastDate).Semester.Single(item => item.Id == termItem.Id).Incomplete, "Semester incomplete state did not persist to the end of term");
            var termSent = semesterStore.Snapshot(); semesterStore.Complete(termSent, WorkbookSync.Write(semesterPath, semesterStore.Book.Version, termSent));
            semesterStore = new AgendaStore(semesterPath);
            Assert(semesterStore.Day(new DateOnly(2026, 11, 23)).Semester.First().Id == termItem.Id && semesterStore.Day(new DateOnly(2026, 11, 23)).Semester.First().Incomplete, "Semester status/order did not round-trip through Excel");
            var planTarget = semesterStore.Day(new DateOnly(2026, 11, 9)).Semester.First();
            semesterStore.Save(planTarget, new JsonObject { ["scheduleDate"] = "2026-11-09" });
            Assert(semesterStore.Day(new DateOnly(2026, 11, 9)).Tasks.Any(entry => entry.Id == planTarget.Id), "Semester planned date did not show up on that day");
            termSent = semesterStore.Snapshot(); semesterStore.Complete(termSent, WorkbookSync.Write(semesterPath, semesterStore.Book.Version, termSent));
            semesterStore = new AgendaStore(semesterPath);
            Assert(semesterStore.Day(new DateOnly(2026, 12, 1)).Semester.Single(entry => entry.Id == planTarget.Id).PlannedDate == "2026-11-09", "Semester planned date did not reach Excel");
            var offTermRejected = false;
            try { semesterStore.Save(planTarget, new JsonObject { ["scheduleDate"] = "2027-03-01" }); } catch (InvalidOperationException) { offTermRejected = true; }
            Assert(offTermRejected, "Out-of-term semester date was accepted");
            semesterStore.Remove(semesterStore.Day(AgendaRules.LastDate).Semester.First(), true);
            Assert(semesterStore.Day(new DateOnly(2026, 11, 9)).Semester.All(item => item.Id != termItem.Id), "Semester deletion left a copy on another date");
            termSent = semesterStore.Snapshot(); semesterStore.Complete(termSent, WorkbookSync.Write(semesterPath, semesterStore.Book.Version, termSent));
            Assert(new AgendaStore(semesterPath).Day(AgendaRules.LastDate).Semester.All(item => item.Id != termItem.Id), "Semester deletion did not reach Excel");
            TestUi(path, fixture);
            Assert(Hash(source) == sourceHash, "Tests changed the personal workbook");
            Console.WriteLine($"PASS: {assertions} native assertions; workbook fidelity, journal recovery, weekly unschedule, dates, concurrency, locks, conflicts, legacy import and WPF UI.");
            Console.WriteLine("QA directory: " + fixture);
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); Console.Error.WriteLine("QA directory: " + fixture); return 1; }
    }
    private static Dictionary<string, string> Entries(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        return archive.Entries.ToDictionary(entry => entry.FullName, entry => { using var stream = entry.Open(); return Convert.ToHexString(SHA256.HashData(stream)); });
    }
    private static void TestLegacyImport(string source, string fixture)
    {
        var path = Path.Combine(fixture, "legacy.xlsx"); File.Copy(source, path);
        var legacy = new JsonObject { ["version"] = "old-version", ["pendingOps"] = new JsonArray(new JsonObject { ["id"] = "local:legacy", ["date"] = "2026-10-20", ["fields"] = new JsonObject { ["title"] = "Imported" } }) };
        var original = legacy.ToJsonString(); File.WriteAllText(path + ".agenda.json", original);
        var store = new AgendaStore(path);
        Assert(store.ImportedLegacy && store.State.Operations.Single().Version == "old-version", "Legacy import bypassed baseline protection");
        Assert(store.Day(new DateOnly(2026, 10, 20)).Tasks.Any(item => item.Title == "Imported"), "Legacy local task disappeared");
        Assert(File.ReadAllText(path + ".agenda.json") == original, "Legacy JSON was modified");
        Assert(new AgendaStore(path).State.Operations.Count == 1, "Legacy import duplicated operations on restart");
    }
    private static void TestUi(string path, string fixture)
    {
        var app = new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };
        var window = new MainWindow(path); app.MainWindow = window; window.Show(); window.UpdateLayout();
        Assert(window.IsVisible && window.ShowInTaskbar && !window.Topmost && window.Icon is not null, "Window/taskbar/icon defaults failed");
        Descendants<System.Windows.Controls.Button>(window).Single(button => button.Content?.ToString() == "学期").RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent)); window.UpdateLayout();
        Assert(Descendants<TextBlock>(window).Any(text => text.Text == "学期计划"), "Semester page navigation failed");
        Assert(Descendants<Border>(window).Any(border => border.ToolTip?.ToString() == "学期计划不分日期"), "Day arrows were not disabled on the semester page");
        var semesterItem = window.Store.Day(AgendaRules.Today).Semester.First();
        var semesterDone = semesterItem.Done;
        window.ToggleItem(semesterItem);
        Assert(window.Store.Day(AgendaRules.LastDate).Semester.Single(item => item.Id == semesterItem.Id).Done != semesterDone, "Semester page checkbox did not update global state");
        var termBitmap = new RenderTargetBitmap((int)window.ActualWidth, (int)window.ActualHeight, 96, 96, PixelFormats.Pbgra32); termBitmap.Render(window);
        var termEncoder = new PngBitmapEncoder(); termEncoder.Frames.Add(BitmapFrame.Create(termBitmap)); using (var output = File.Create(Path.Combine(fixture, "semester-ui.png"))) termEncoder.Save(output);
        Descendants<System.Windows.Controls.Button>(window).First(button => button.Content?.ToString() == "今天").RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent)); window.UpdateLayout();
        // 拖进日程：落点决定位置，时间由程序算，不能和已有安排重叠
        var today = AgendaRules.Today;
        var slotDay = AgendaRules.DateText(today);
        var slotA = new AgendaItem { Id = "local:slot-a", Date = slotDay, Planned = 60 };
        var slotB = new AgendaItem { Id = "local:slot-b", Date = slotDay, Planned = 60 };
        window.Store.Save(slotA, new JsonObject { ["title"] = "排期甲", ["planned"] = 60 });
        window.Store.Save(slotB, new JsonObject { ["title"] = "排期乙", ["planned"] = 60 });
        window.Render();
        var scheduleBefore = window.Store.Day(today).Schedule.ToArray();
        Assert(scheduleBefore.Length > 0, "Fixture has no schedule block to drop after");
        var firstBlock = scheduleBefore[0];
        var firstTimes = firstBlock.Time.Split('—');
        var firstEnd = firstTimes.Length == 2 ? AgendaRules.Minutes(firstTimes[1], true) : AgendaRules.Minutes(firstTimes[0]);
        var placedA = window.PlaceInDay(slotA, today, firstBlock, true);
        Assert(placedA.StartsWith("已排 ", StringComparison.Ordinal), "Drop did not report an assigned time: " + placedA);
        var placedAItem = window.Store.Day(today).Schedule.Single(item => item.Id == "local:slot-a");
        var placedAStart = AgendaRules.Minutes(placedAItem.Start);
        Assert(placedAStart >= firstEnd, "Dropped task landed before the block it was dropped under");
        Assert(window.Store.Day(today).Schedule.Where(item => item.Id != placedAItem.Id).All(item =>
        {
            var times = item.Time.Split('—'); if (times.Length != 2) return true;
            return placedAStart >= AgendaRules.Minutes(times[1], true) || placedAStart + 60 <= AgendaRules.Minutes(times[0]);
        }), "Dropped task overlaps an existing block");
        var placedB = window.PlaceInDay(slotB, today, placedAItem, true);
        var placedBItem = window.Store.Day(today).Schedule.Single(item => item.Id == "local:slot-b");
        var placedBStart = AgendaRules.Minutes(placedBItem.Start);
        Assert(placedBStart >= placedAStart + 60, "Second drop overlapped the first one: " + placedB);
        // 连锁顺延：插到 B 前面时，B 及其后面的任务自动后移
        var slotC = new AgendaItem { Id = "local:slot-c", Date = slotDay, Planned = 60 };
        window.Store.Save(slotC, new JsonObject { ["title"] = "排期丙", ["planned"] = 60 });
        window.Render();
        var cascadeResult = window.PlaceInDay(slotC, today, placedBItem, false);
        var afterCascade = window.Store.Day(today).Schedule;
        var slotCStart = AgendaRules.Minutes(afterCascade.Single(item => item.Id == "local:slot-c").Start);
        var shiftedB = AgendaRules.Minutes(afterCascade.Single(item => item.Id == "local:slot-b").Start);
        var keptA = AgendaRules.Minutes(afterCascade.Single(item => item.Id == "local:slot-a").Start);
        Assert(slotCStart == placedBStart && shiftedB == placedBStart + 60, "Cascade did not push the following task: " + cascadeResult);
        Assert(keptA == placedAStart, "Cascade moved a task that came before the drop");
        Assert(cascadeResult.Contains("顺延 1 项", StringComparison.Ordinal), "Cascade result did not report the shift: " + cascadeResult);
        // 时间流：空档里拖到哪一刻就是哪一刻，被占了才往后找空位
        var slotD = new AgendaItem { Id = "local:slot-d", Date = slotDay, Planned = 45 };
        window.Store.Save(slotD, new JsonObject { ["title"] = "精确定位", ["planned"] = 45 });
        window.Render();
        var freeSlot = FirstFreeMinute(window.Store.Day(today).Schedule, 60);
        var exact = window.PlaceAt(slotD, today, freeSlot, 45);
        var exactItem = window.Store.Day(today).Schedule.Single(item => item.Id == "local:slot-d");
        Assert(exactItem.Start == AgendaRules.MinuteText(freeSlot) && exact.StartsWith("已排 " + AgendaRules.MinuteText(freeSlot), StringComparison.Ordinal), "Exact-time drop did not land on " + AgendaRules.MinuteText(freeSlot) + ": " + exact + " / " + string.Join(" | ", window.Store.Day(today).Schedule.Select(entry => entry.Title + " " + entry.Time)));
        var slotE = new AgendaItem { Id = "local:slot-e", Date = slotDay, Planned = 30 };
        window.Store.Save(slotE, new JsonObject { ["title"] = "让位", ["planned"] = 30 });
        window.Render();
        var shifted = window.PlaceAt(slotE, today, freeSlot + 10, 30);
        var shiftedItem = window.Store.Day(today).Schedule.Single(item => item.Id == "local:slot-e");
        var pushedDown = window.Store.Day(today).Schedule.Single(item => item.Id == "local:slot-d");
        Assert(AgendaRules.Minutes(shiftedItem.Start) == freeSlot + 10 && AgendaRules.Minutes(pushedDown.Start) == freeSlot + 40 && shifted.Contains("顺延 1 项", StringComparison.Ordinal), "Drop inside a block did not push it: " + shifted);
        // 空档拖放：拖进去才展开刻度，落在哪一刻就是哪一刻
        var gapTask = new AgendaItem { Id = "local:slot-gap", Date = slotDay, Title = "空档落点" };
        window.Store.Save(gapTask, new JsonObject { ["title"] = "空档落点" });
        window.Render(); window.UpdateLayout();
        var gapHost = Descendants<Border>(window).FirstOrDefault(border => border.ToolTip?.ToString() == "把任务拖到这里，按刻度选开始时间");
        Assert(gapHost is not null, "Gap drop target is missing");
        var payload = new DataObject("QingxuTask", window.Store.Day(today).Tasks.Single(item => item.Id == gapTask.Id));
        var enter = DragEvent(payload, gapHost!, new Point(90, 0)); enter.RoutedEvent = DragDrop.DragEnterEvent; gapHost!.RaiseEvent(enter);
        window.UpdateLayout();
        var gapHeight = gapHost.ActualHeight;
        Assert(gapHeight > 40, "Gap did not expand while dragging: " + gapHeight);
        var over = DragEvent(payload, gapHost, new Point(90, gapHeight - 4)); over.RoutedEvent = DragDrop.DragOverEvent; gapHost.RaiseEvent(over);
        gapHost.BringIntoView(); window.UpdateLayout();
        var gapBitmap = new RenderTargetBitmap((int)window.ActualWidth, (int)window.ActualHeight, 96, 96, PixelFormats.Pbgra32); gapBitmap.Render(window);
        var gapEncoder = new PngBitmapEncoder(); gapEncoder.Frames.Add(BitmapFrame.Create(gapBitmap)); using (var gapImage = File.Create(Path.Combine(fixture, "gap-drop.png"))) gapEncoder.Save(gapImage);
        var drop = DragEvent(payload, gapHost, new Point(90, gapHeight - 4)); drop.RoutedEvent = DragDrop.DropEvent; gapHost.RaiseEvent(drop);
        var gapPlaced = window.Store.Day(today).Schedule.SingleOrDefault(item => item.Id == gapTask.Id);
        Assert(gapPlaced is not null && gapPlaced.Start.Length == 5, "Gap drop did not assign a start time");
        // 考试判定与提前提醒
        Assert(AgendaRules.IsExam("期中", "随便") && AgendaRules.IsExam("", "工训机考+报告") && !AgendaRules.IsExam("", "概统小测复习") && !AgendaRules.IsExam("志愿", "队长拿衣服"), "Exam detection misfired");
        Assert(AgendaReader.Read(new XlsxBook(path), new DateOnly(2026, 11, 14)).Schedule.Any(item => item.Exam && item.Category == "期中"), "Exam block is not flagged in the agenda");
        Assert(AgendaReader.Read(new XlsxBook(path), new DateOnly(2026, 11, 8)).Exams.Any(item => item.Title.Contains("大学物理")), "Upcoming exam list is empty");
        // 拖到课程上：并进这门课的课内任务
        var classDay = new DateOnly(2026, 11, 9);
        var classItem = AgendaReader.Read(new XlsxBook(path), classDay).Schedule.FirstOrDefault(item => item.Class);
        Assert(classItem is not null, "Fixture has no course block");
        var absorb = new AgendaItem { Id = "local:absorb", Date = AgendaRules.DateText(classDay), Title = "并进课内" };
        window.Store.Save(absorb, new JsonObject { ["title"] = "并进课内" });
        window.AbsorbIntoClass(absorb, classItem!);
        Assert(window.Store.Day(classDay).Schedule.Single(item => item.Id == classItem!.Id).TaskTitle.Contains("并进课内"), "Course task absorb did not write the class title");
        Assert(window.Store.Day(classDay).Tasks.All(item => item.Id != absorb.Id), "Absorbed task stayed in the todo list");
        // 行内撤销：课内任务还原、任务回到待办、日程条目取消时间
        Descendants<System.Windows.Controls.Button>(window).First(button => button.Content?.ToString() == "11/09").RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent)); window.UpdateLayout();
        var undoAbsorb = Descendants<System.Windows.Controls.Button>(window).FirstOrDefault(button => button.ToolTip?.ToString() == "撤销并入，把这件任务放回待办");
        Assert(undoAbsorb is not null, "Absorb undo button is missing");
        undoAbsorb!.RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent)); window.UpdateLayout();
        Assert(window.Store.Day(classDay).Schedule.Single(item => item.Id == classItem!.Id).TaskTitle == classItem!.TaskTitle, "Undo did not restore the course task");
        Assert(window.Store.Day(classDay).Tasks.Any(item => item.Title == "并进课内"), "Undo did not return the task to the todo list");
        // 老的课内任务可以清空（编辑框留空保存）
        var clearing = window.Store.Day(classDay).Schedule.Single(item => item.Id == classItem!.Id);
        window.Store.Save(clearing, new JsonObject { ["title"] = "" });
        var cleared = window.Store.Day(classDay).Schedule.Single(item => item.Id == classItem!.Id);
        Assert(cleared.TaskTitle.Length == 0, "Clearing the course task was rejected: " + cleared.TaskTitle);
        window.Store.Save(clearing, new JsonObject { ["title"] = "并进课内" });
        Descendants<System.Windows.Controls.Button>(window).First(button => button.Content?.ToString() == "今天").RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent)); window.UpdateLayout();
        var inlineUndo = Descendants<System.Windows.Controls.Button>(window).FirstOrDefault(button => button.ToolTip?.ToString() == "撤销时间排期，放回待办");
        Assert(inlineUndo is not null, "Inline unschedule button is missing");
        var beforeUndo = window.Store.Day(today).Schedule.Count;
        inlineUndo!.RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent)); window.UpdateLayout();
        Assert(window.Store.Day(today).Schedule.Count == beforeUndo - 1, "Inline unschedule did not clear the time");
        // 整条任务都能抓：日程条和待办卡片都挂上任务，作为拖动的抓取目标
        var headTarget = window.Store.Day(today).Tasks.First(item => item.Editable);
        Assert(Descendants<Border>(window).Any(border => border.Tag is AgendaItem tagged && tagged.Id == headTarget.Id && border.AllowDrop), "Todo card is not marked as a drag source");
        Assert(Descendants<Grid>(window).Any(grid => grid.Tag as string == "head"), "Row head is not marked for click detection");
        var listBitmap = new RenderTargetBitmap((int)window.ActualWidth, (int)window.ActualHeight, 96, 96, PixelFormats.Pbgra32); listBitmap.Render(window);
        var listEncoder = new PngBitmapEncoder(); listEncoder.Frames.Add(BitmapFrame.Create(listBitmap)); using (var listOutput = File.Create(Path.Combine(fixture, "schedule-drop.png"))) listEncoder.Save(listOutput);
        window.Close(); window.Restore(); window.UpdateLayout();
        var bitmap = new RenderTargetBitmap((int)window.ActualWidth, (int)window.ActualHeight, 96, 96, PixelFormats.Pbgra32); bitmap.Render(window);
        var encoder = new PngBitmapEncoder(); encoder.Frames.Add(BitmapFrame.Create(bitmap)); using (var image = File.Create(Path.Combine(fixture, "native-preview.png"))) encoder.Save(image);
        // 有课的一天：课程 + 课内任务 + 时间流的样子
        Descendants<System.Windows.Controls.Button>(window).First(button => button.Content?.ToString() == "11/09").RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent)); window.UpdateLayout();
        var classBitmap = new RenderTargetBitmap((int)window.ActualWidth, (int)window.ActualHeight, 96, 96, PixelFormats.Pbgra32); classBitmap.Render(window);
        var classEncoder = new PngBitmapEncoder(); classEncoder.Frames.Add(BitmapFrame.Create(classBitmap)); using (var classImage = File.Create(Path.Combine(fixture, "class-flow.png"))) classEncoder.Save(classImage);
        Descendants<System.Windows.Controls.Button>(window).First(button => button.Content?.ToString() == "今天").RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent)); window.UpdateLayout();
        var buttons = Descendants<System.Windows.Controls.Button>(window).ToArray();
        buttons.First(button => button.Content?.ToString() == "本周").RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent)); window.UpdateLayout();
        Assert(Descendants<TextBlock>(window).Any(text => text.Text == "本周常规任务"), "Weekly navigation failed");
        buttons = Descendants<System.Windows.Controls.Button>(window).ToArray(); buttons.First(button => button.Content?.ToString() == "今天").RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent));
        var toggleDay = window.Store.Day(new DateOnly(2026, 10, 20));
        var toggleTarget = toggleDay.Tasks.FirstOrDefault() ?? toggleDay.Schedule.FirstOrDefault(item => item.Editable);
        Assert(toggleTarget is not null, "Fixture has no toggleable task");
        var toggleCount = window.Store.State.Operations.Count;
        window.ToggleItem(toggleTarget!);
        Assert(window.Store.State.Operations.Count > toggleCount, "Checkbox did not save");
        var edited = new AgendaItem { Id = "local:editor", Date = "2026-10-20", Planned = 30 };
        var editor = new TaskEditor(window, edited);
        var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(100) };
        timer.Tick += (_, _) =>
        {
            timer.Stop(); var inputs = Descendants<TextBox>(editor).ToArray();
            inputs[0].Text = "Manual time"; inputs[1].Text = "21:30"; inputs[2].Text = "22:00";
            Descendants<System.Windows.Controls.Button>(editor).Single(button => button.Content?.ToString() == "保存").RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent));
        };
        timer.Start(); Assert(editor.ShowDialog() == true && AgendaRules.Text(editor.Result!, "start") == "21:30" && AgendaRules.Number(editor.Result!, "planned") == 30, "Native editor did not accept explicit start/end times");
        var wideEditor = new TaskEditor(window, new AgendaItem { Id = "local:wide-time", Date = "2026-10-20" });
        var wideTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(100) };
        wideTimer.Tick += (_, _) =>
        {
            wideTimer.Stop(); var wideInputs = Descendants<TextBox>(wideEditor).ToArray();
            wideInputs[0].Text = "全角时间"; wideInputs[1].Text = "９：０５"; wideInputs[2].Text = "１０：１５";
            Descendants<System.Windows.Controls.Button>(wideEditor).Single(button => button.Content?.ToString() == "保存").RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent));
        };
        wideTimer.Start();
        Assert(wideEditor.ShowDialog() == true && AgendaRules.Text(wideEditor.Result!, "start") == "09:05" && AgendaRules.Number(wideEditor.Result!, "planned") == 70, "Full-width time was not normalised by the editor");
        window.Store.Save(edited, editor.Result!);
        window.Store.Save(edited, new JsonObject { ["start"] = "", ["planned"] = "" });
        Assert(window.Store.Day(new DateOnly(2026, 10, 20)).Tasks.Single(item => item.Id == edited.Id).Detail.Length == 0, "Clearing the editor time left a subtitle");
        // 快速添加：回车必须真的落盘，并把输入框清空
        var quickBox = Descendants<TextBox>(window).FirstOrDefault(box => box.ToolTip?.ToString()?.Contains("回车") == true);
        Assert(quickBox is not null, "Quick add input is missing");
        var quickBefore = window.Store.State.Operations.Count;
        quickBox!.Text = "快速添加";
        var quickButton = Descendants<System.Windows.Controls.Button>(window).FirstOrDefault(button => button.Content?.ToString() == "添加");
        Assert(quickButton is not null && quickButton.IsEnabled, "Quick add button stayed disabled after typing");
        var presentation = PresentationSource.FromVisual(window);
        quickBox.RaiseEvent(new KeyEventArgs(Keyboard.PrimaryDevice, presentation!, 0, Key.Enter) { RoutedEvent = Keyboard.KeyDownEvent });
        Assert(window.Store.State.Operations.Count == quickBefore + 1, "Quick add via Enter did not save");
        Assert(quickBox.Text.Length == 0, "Quick add did not clear the input");
        Assert(window.Store.Day(AgendaRules.Today).Tasks.Any(item => item.Title == "快速添加"), "Quick add item is not on the visible day");
        Descendants<System.Windows.Controls.Button>(window).First(button => button.Content?.ToString() == "本周").RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent));
        window.UpdateLayout();
        Assert(!quickBox.IsEnabled, "Quick add should be disabled on the weekly page");
        Descendants<System.Windows.Controls.Button>(window).First(button => button.Content?.ToString() == "今天").RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent));
        window.UpdateLayout();
        window.Width = 365; window.UpdateLayout();
        var sync = Descendants<System.Windows.Controls.Button>(window).Single(button => button.Content?.ToString() == "同步 Excel");
        var origin = sync.TranslatePoint(new Point(0, 0), window);
        Assert(origin.X >= 0 && origin.X + sync.ActualWidth <= window.ActualWidth, "Sync button is clipped at minimum width");
        var reviewPath = Path.Combine(fixture, "review-ui.xlsx"); File.Copy(path, reviewPath);
        var reviewStore = new AgendaStore(reviewPath);
        reviewStore.Save(new AgendaItem { Id = "local:review-one", Date = "2026-10-20" }, new JsonObject { ["title"] = "Review one" });
        reviewStore.Save(new AgendaItem { Id = "local:review-two", Date = "2026-10-20" }, new JsonObject { ["title"] = "Review two" });
        reviewStore.Save(new AgendaItem { Id = "daily:599", Date = "2026-10-20" }, new JsonObject { ["title"] = "Deleted task" });
        var review = new ConflictReview(window, reviewStore); review.Show(); review.UpdateLayout();
        var choices = Descendants<CheckBox>(review).ToArray();
        Assert(choices.Count(choice => choice.IsEnabled) == 2 && choices.Any(choice => !choice.IsEnabled), "Review fixture lost unavailable-task protection");
        var selectAll = Descendants<System.Windows.Controls.Button>(review).Single(button => button.Content?.ToString() == "全选");
        var clearAll = Descendants<System.Windows.Controls.Button>(review).Single(button => button.Content?.ToString() == "取消全选");
        selectAll.RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent));
        Assert(choices.Where(choice => choice.IsEnabled).All(choice => choice.IsChecked == true) && choices.Where(choice => !choice.IsEnabled).All(choice => choice.IsChecked != true), "Select all selected unavailable tasks or missed valid tasks");
        clearAll.RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent));
        Assert(choices.All(choice => choice.IsChecked == false), "Clear all left selected changes");
        review.Close();
        var syncPath = Path.Combine(fixture, "sync-ui.xlsx"); File.Copy(path, syncPath);
        var syncWindow = new MainWindow(syncPath); syncWindow.Show(); syncWindow.UpdateLayout();
        syncWindow.Store.Save(new AgendaItem { Id = "local:ui-sync", Date = "2026-10-20" }, new JsonObject { ["title"] = "Background sync test" });
        var syncOnly = Descendants<System.Windows.Controls.Button>(syncWindow).Single(button => button.Content?.ToString() == "同步 Excel");
        syncOnly.RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent));
        var deadline = DateTime.UtcNow.AddSeconds(15);
        while (!syncOnly.IsEnabled && DateTime.UtcNow < deadline)
        {
            var frame = new DispatcherFrame();
            Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background, new Action(() => frame.Continue = false));
            Dispatcher.PushFrame(frame); Thread.Sleep(10);
        }
        Assert(syncOnly.IsEnabled && syncWindow.Store.State.Operations.Count == 0, "Sync-only button did not finish or clear saved operations");
        Assert(AgendaReader.Read(new XlsxBook(syncPath), new DateOnly(2026, 10, 20)).Tasks.Any(item => item.Title == "Background sync test"), "Sync-only button did not write the workbook");
        syncOnly.RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent));
        Assert(syncWindow.Store.State.Operations.Count == 0, "Empty sync produced new operations");
        syncWindow.Exit();
        window.Close(); Assert(!window.IsVisible, "Close did not hide to tray"); window.Restore(); Assert(window.IsVisible, "Tray restore failed");
        var process = Process.GetCurrentProcess();
        Console.WriteLine($"Native UI working set after initial rendering: {process.WorkingSet64 / 1048576.0:F1} MiB (test process, not idle production measurement).");
        window.Exit();
    }
    /// <summary>当天从 08:00 起第一段够长的空档（测试用，不假设工作簿的具体安排）。</summary>
    private static int FirstFreeMinute(IEnumerable<AgendaItem> schedule, int length)
    {
        var blocks = new List<(int Start, int End)>();
        foreach (var item in schedule)
        {
            var parts = item.Time.Split('—');
            if (parts.Length == 0 || parts[0].Length == 0) continue;
            var start = AgendaRules.Minutes(parts[0]);
            var end = parts.Length > 1 && parts[1].Length > 0 ? AgendaRules.Minutes(parts[1], true) : start + 15;
            blocks.Add((start, Math.Max(end, start + 5)));
        }
        var cursor = 8 * 60;
        foreach (var block in blocks.OrderBy(range => range.Start))
        {
            if (block.Start - cursor >= length) return cursor;
            cursor = Math.Max(cursor, block.End);
        }
        return cursor;
    }
    private static DragEventArgs DragEvent(IDataObject data, DependencyObject target, Point point)
    {
        // DragEventArgs 的构造函数是 internal，测试里用反射造一个同等的拖放事件
        var constructors = typeof(DragEventArgs).GetConstructors(System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Public);
        var five = constructors.FirstOrDefault(candidate => candidate.GetParameters().Length == 5);
        if (five is not null) return (DragEventArgs)five.Invoke(new object[] { data, DragDropKeyStates.LeftMouseButton, DragDropEffects.Move, target, point });
        var six = constructors.First(candidate => candidate.GetParameters().Length == 6);
        return (DragEventArgs)six.Invoke(new object[] { data, DragDropKeyStates.LeftMouseButton, DragDropEffects.Move, DragDropEffects.Move, target, point });
    }
    private static IEnumerable<T> Descendants<T>(DependencyObject root) where T : DependencyObject
    {
        for (var index = 0; index < VisualTreeHelper.GetChildrenCount(root); index++)
        {
            var child = VisualTreeHelper.GetChild(root, index);
            if (child is T match) yield return match;
            foreach (var descendant in Descendants<T>(child)) yield return descendant;
        }
    }
}
