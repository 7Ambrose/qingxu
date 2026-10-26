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
            Assert(AgendaReader.Read(book, new DateOnly(2026, 10, 19)).Schedule.Single(item => item.Id == "habit:秋6周:B:67").Detail == "08:30—08:55", "Routine time failed");
            var store = new AgendaStore(path);
            var dragTask = new AgendaItem { Id = "local:drag-conflict", Date = "2026-10-20", Planned = 30 };
            var dragRejected = false;
            try { store.Schedule(dragTask, new DateOnly(2026, 10, 20), "13:30"); } catch (InvalidOperationException) { dragRejected = true; }
            Assert(dragRejected && !store.State.Operations.Any(operation => operation.Id == dragTask.Id), "13:30 course conflict reached the local queue");
            Assert(AgendaRules.Minutes("13:30") == 810 && AgendaRules.EndTime("23:30", 30) == "24:00", "Time grid or midnight boundary failed");
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
        var bitmap = new RenderTargetBitmap((int)window.ActualWidth, (int)window.ActualHeight, 96, 96, PixelFormats.Pbgra32); bitmap.Render(window);
        var encoder = new PngBitmapEncoder(); encoder.Frames.Add(BitmapFrame.Create(bitmap)); using (var image = File.Create(Path.Combine(fixture, "native-preview.png"))) encoder.Save(image);
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
        window.Store.Save(edited, editor.Result!);
        window.Store.Save(edited, new JsonObject { ["start"] = "", ["planned"] = "" });
        Assert(window.Store.Day(new DateOnly(2026, 10, 20)).Tasks.Single(item => item.Id == edited.Id).Detail.Length == 0, "Clearing the editor time left a subtitle");
        window.Width = 365; window.UpdateLayout();
        var sync = Descendants<System.Windows.Controls.Button>(window).Single(button => button.Content?.ToString() == "同步并打开Excel");
        var origin = sync.TranslatePoint(new Point(0, 0), window);
        Assert(origin.X >= 0 && origin.X + sync.ActualWidth <= window.ActualWidth, "Sync button is clipped at minimum width");
        window.Close(); Assert(!window.IsVisible, "Close did not hide to tray"); window.Restore(); Assert(window.IsVisible, "Tray restore failed");
        var process = Process.GetCurrentProcess();
        Console.WriteLine($"Native UI working set after initial rendering: {process.WorkingSet64 / 1048576.0:F1} MiB (test process, not idle production measurement).");
        window.Exit();
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
