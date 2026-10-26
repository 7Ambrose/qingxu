using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Qingxu;

public sealed class XlsxBook
{
    private static readonly XNamespace Spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private readonly Dictionary<string, XDocument> documents = new();
    private readonly Dictionary<string, string> paths = new();
    private readonly Dictionary<string, Dictionary<string, object?>> cells = new();
    private readonly HashSet<string> changed = new();
    private readonly string[] shared;
    private readonly bool editable;
    public string Version { get; }
    public string Path { get; }
    public IEnumerable<string> SheetNames => paths.Keys;

    public XlsxBook(string path, bool editable = false)
    {
        Path = System.IO.Path.GetFullPath(path);
        this.editable = editable;
        Version = CurrentVersion(Path);
        using var stream = new FileStream(Path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var sharedEntry = archive.GetEntry("xl/sharedStrings.xml");
        shared = sharedEntry is null ? Array.Empty<string>() : Read(sharedEntry).Root!.Elements(Spreadsheet + "si")
            .Select(element => string.Concat(element.Descendants(Spreadsheet + "t").Select(text => text.Value))).ToArray();
        var workbook = Read(archive.GetEntry("xl/workbook.xml") ?? throw new InvalidDataException("不是有效的 Excel 工作簿。"));
        var relationships = Read(archive.GetEntry("xl/_rels/workbook.xml.rels")!);
        var targets = relationships.Root!.Elements().ToDictionary(element => (string)element.Attribute("Id")!, element => (string)element.Attribute("Target")!);
        foreach (var sheet in workbook.Root!.Element(Spreadsheet + "sheets")!.Elements())
        {
            var name = (string)sheet.Attribute("name")!;
            var relationId = (string)sheet.Attribute(XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"))!;
            var target = targets[relationId];
            var resolved = new Uri(new Uri("https://xlsx.local/xl/workbook.xml"), target).AbsolutePath.TrimStart('/');
            paths[name] = resolved;
        }
        if (Version != CurrentVersion(Path)) throw new IOException("读取期间 Excel 有更新，请保存并关闭后重试。");
    }
    public void Prepare(params string[] names)
    {
        var pending = names.Where(name => !cells.ContainsKey(name) && paths.ContainsKey(name)).Distinct().ToArray();
        if (pending.Length == 0) return;
        if (Version != CurrentVersion(Path)) throw new SyncConflictException();
        using var stream = new FileStream(Path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        foreach (var name in pending) Load(archive, name);
        if (Version != CurrentVersion(Path)) throw new IOException("读取期间 Excel 有更新，请保存并关闭后重试。");
    }
    private void LoadSheet(string name)
    {
        if (cells.ContainsKey(name) || !paths.ContainsKey(name)) return;
        if (Version != CurrentVersion(Path)) throw new SyncConflictException();
        using var stream = new FileStream(Path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        Load(archive, name);
        if (Version != CurrentVersion(Path)) throw new IOException("读取期间 Excel 有更新，请保存并关闭后重试。");
    }
    private void Load(ZipArchive archive, string name)
    {
        var path = paths[name];
        var document = Read(archive.GetEntry(path) ?? throw new InvalidDataException("工作表不存在：" + name));
        if (editable) documents[name] = document;
        var values = new Dictionary<string, object?>();
        foreach (var cell in document.Descendants(Spreadsheet + "c"))
        {
                var address = (string?)cell.Attribute("r");
                if (address is null) continue;
                var type = (string?)cell.Attribute("t");
                var raw = cell.Element(Spreadsheet + "v")?.Value;
                object? value = type switch
                {
                    "s" => int.TryParse(raw, out var index) && index >= 0 && index < shared.Length ? shared[index] : "",
                    "inlineStr" => string.Concat(cell.Element(Spreadsheet + "is")?.Descendants(Spreadsheet + "t").Select(text => text.Value) ?? Array.Empty<string>()),
                    "b" => raw == "1",
                    "str" or "e" => raw ?? "",
                    _ => double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) ? number : raw
                };
                if (value is not null && value is not "") values[address] = value;
        }
        cells[name] = values;
    }
    private static XDocument Read(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
        return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
    }
    public static string CurrentVersion(string path) => ((File.GetLastWriteTimeUtc(path).Ticks - DateTime.UnixEpoch.Ticks) * 100).ToString(CultureInfo.InvariantCulture);
    public bool HasSheet(string sheet) => paths.ContainsKey(sheet);
    public object? Get(string sheet, string address)
    {
        LoadSheet(sheet);
        return cells.TryGetValue(sheet, out var values) && values.TryGetValue(address, out var value) ? value : null;
    }
    public string Text(string sheet, string address) => Convert.ToString(Get(sheet, address), CultureInfo.InvariantCulture)?.Trim() ?? "";
    public double? Number(string sheet, string address) => double.TryParse(Text(sheet, address), NumberStyles.Float, CultureInfo.InvariantCulture, out var number) ? number : null;
    public void Set(string sheet, string address, object? value)
    {
        LoadSheet(sheet);
        if (!documents.TryGetValue(sheet, out var document)) throw new InvalidOperationException("缺少工作表：" + sheet);
        var sheetData = document.Root!.Element(Spreadsheet + "sheetData") ?? throw new InvalidDataException("工作表缺少数据区域。");
        var rowNumber = int.Parse(new string(address.Where(char.IsDigit).ToArray()), CultureInfo.InvariantCulture);
        var row = sheetData.Elements(Spreadsheet + "row").FirstOrDefault(element => (int?)element.Attribute("r") == rowNumber);
        if (row is null)
        {
            row = new XElement(Spreadsheet + "row", new XAttribute("r", rowNumber));
            var nextRow = sheetData.Elements(Spreadsheet + "row").FirstOrDefault(element => (int?)element.Attribute("r") > rowNumber);
            if (nextRow is null) sheetData.Add(row); else nextRow.AddBeforeSelf(row);
        }
        var cell = row.Elements(Spreadsheet + "c").FirstOrDefault(element => (string?)element.Attribute("r") == address);
        if (cell is null)
        {
            cell = new XElement(Spreadsheet + "c", new XAttribute("r", address));
            var nextCell = row.Elements(Spreadsheet + "c").FirstOrDefault(element => Column((string)element.Attribute("r")!) > Column(address));
            if (nextCell is null) row.Add(cell); else nextCell.AddBeforeSelf(cell);
        }
        cell.Elements().Where(element => element.Name == Spreadsheet + "v" || element.Name == Spreadsheet + "f" || element.Name == Spreadsheet + "is").Remove();
        cell.Attribute("t")?.Remove();
        if (value is string text)
        {
            cell.SetAttributeValue("t", "inlineStr");
            cell.Add(new XElement(Spreadsheet + "is", new XElement(Spreadsheet + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text)));
        }
        else if (value is not null)
        {
            if (value is bool flag) { cell.SetAttributeValue("t", "b"); cell.Add(new XElement(Spreadsheet + "v", flag ? "1" : "0")); }
            else cell.Add(new XElement(Spreadsheet + "v", Convert.ToString(value, CultureInfo.InvariantCulture)));
        }
        cells[sheet][address] = value; changed.Add(sheet);
    }
    private static int Column(string address)
    {
        var result = 0;
        foreach (var character in address.TakeWhile(char.IsLetter)) result = result * 26 + character - 'A' + 1;
        return result;
    }
    public void Export(string destination)
    {
        File.Copy(Path, destination, false);
        using var stream = new FileStream(destination, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Update);
        foreach (var sheet in changed)
        {
            archive.GetEntry(paths[sheet])!.Delete();
            var entry = archive.CreateEntry(paths[sheet], CompressionLevel.Optimal);
            using var output = entry.Open();
            using var writer = XmlWriter.Create(output, new XmlWriterSettings { Encoding = new UTF8Encoding(false), OmitXmlDeclaration = false });
            documents[sheet].Save(writer);
        }
        var workbookEntry = archive.GetEntry("xl/workbook.xml")!;
        var workbook = Read(workbookEntry);
        var calculation = workbook.Root!.Element(Spreadsheet + "calcPr");
        if (calculation is null) { calculation = new XElement(Spreadsheet + "calcPr"); workbook.Root.Add(calculation); }
        calculation.SetAttributeValue("fullCalcOnLoad", "1"); calculation.SetAttributeValue("forceFullCalc", "1");
        workbookEntry.Delete();
        using var workbookOutput = archive.CreateEntry("xl/workbook.xml").Open();
        workbook.Save(workbookOutput);
    }
}
