using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;

namespace TradeXmlStudio.Core;

/// <summary>
/// Reads the small subset of the Office Open XML workbook format needed by the
/// batch workflow. This keeps the application dependency-free and supports
/// .xlsx and .xlsm files without requiring Excel to be installed.
/// </summary>
public sealed class OpenXmlWorkbookReader
{
    private static readonly XNamespace Spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace OfficeRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";

    public IReadOnlyList<string> ListSheets(string workbookPath)
    {
        using var archive = OpenWorkbook(workbookPath);
        var workbook = LoadXml(archive, "xl/workbook.xml");
        return workbook.Root?
            .Element(Spreadsheet + "sheets")?
            .Elements(Spreadsheet + "sheet")
            .Select(sheet => (string?)sheet.Attribute("name"))
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Cast<string>()
            .ToList() ?? [];
    }

    public IReadOnlyList<ExcelBatchEntry> ReadEntries(string workbookPath, string sheetName)
    {
        using var archive = OpenWorkbook(workbookPath);
        var sheetEntryPath = ResolveSheetPath(archive, sheetName);
        var sharedStrings = ReadSharedStrings(archive);
        var worksheet = LoadXml(archive, sheetEntryPath);
        var rows = worksheet.Descendants(Spreadsheet + "row");
        var result = new List<ExcelBatchEntry>();

        foreach (var row in rows)
        {
            var rowNumber = ParsePositiveInt((string?)row.Attribute("r")) ?? result.Count + 1;
            var values = row.Elements(Spreadsheet + "c")
                .Select(cell => new
                {
                    Column = GetColumnName((string?)cell.Attribute("r")),
                    Value = ReadCellValue(cell, sharedStrings)
                })
                .Where(cell => cell.Column is "B" or "C")
                .ToDictionary(cell => cell.Column!, cell => cell.Value, StringComparer.OrdinalIgnoreCase);

            values.TryGetValue("B", out var serialText);
            values.TryGetValue("C", out var lotId);
            serialText = serialText?.Trim() ?? "";
            lotId = lotId?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(lotId) || lotId == "箱号" || serialText == "序号")
            {
                continue;
            }

            var serial = int.TryParse(serialText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed)
                ? parsed
                : rowNumber;
            result.Add(new ExcelBatchEntry(serial, lotId));
        }

        return result;
    }

    private static ZipArchive OpenWorkbook(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("Excel 路径不能为空。", nameof(path));
        }
        if (!File.Exists(path))
        {
            throw new FileNotFoundException("Excel 文件不存在。", path);
        }
        var extension = Path.GetExtension(path);
        if (!extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
            && !extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
        {
            throw new NotSupportedException("仅支持 .xlsx 和 .xlsm 工作簿。");
        }

        try
        {
            return ZipFile.OpenRead(path);
        }
        catch (InvalidDataException ex)
        {
            throw new InvalidDataException("文件不是有效的 Excel Open XML 工作簿。", ex);
        }
    }

    private static string ResolveSheetPath(ZipArchive archive, string sheetName)
    {
        var workbook = LoadXml(archive, "xl/workbook.xml");
        var sheet = workbook.Root?
            .Element(Spreadsheet + "sheets")?
            .Elements(Spreadsheet + "sheet")
            .FirstOrDefault(node => string.Equals((string?)node.Attribute("name"), sheetName, StringComparison.Ordinal));
        if (sheet is null)
        {
            throw new InvalidOperationException($"工作表不存在：{sheetName}");
        }

        var relationshipId = (string?)sheet.Attribute(OfficeRelationships + "id")
            ?? throw new InvalidDataException("工作表缺少关系标识。");
        var relationships = LoadXml(archive, "xl/_rels/workbook.xml.rels");
        var relationship = relationships.Root?
            .Elements(PackageRelationships + "Relationship")
            .FirstOrDefault(node => string.Equals((string?)node.Attribute("Id"), relationshipId, StringComparison.Ordinal));
        var target = (string?)relationship?.Attribute("Target")
            ?? throw new InvalidDataException("无法定位工作表文件。");

        return NormalizeZipPath("xl", target);
    }

    private static IReadOnlyList<string> ReadSharedStrings(ZipArchive archive)
    {
        var entry = FindEntry(archive, "xl/sharedStrings.xml");
        if (entry is null)
        {
            return [];
        }

        using var stream = entry.Open();
        var document = XDocument.Load(stream);
        return document.Root?
            .Elements(Spreadsheet + "si")
            .Select(item => string.Concat(item.Descendants(Spreadsheet + "t").Select(text => text.Value)))
            .ToList() ?? [];
    }

    private static string ReadCellValue(XElement cell, IReadOnlyList<string> sharedStrings)
    {
        var type = (string?)cell.Attribute("t");
        if (type == "inlineStr")
        {
            return string.Concat(cell.Descendants(Spreadsheet + "t").Select(text => text.Value));
        }

        var raw = cell.Element(Spreadsheet + "v")?.Value ?? "";
        if (type == "s" && int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index))
        {
            return index >= 0 && index < sharedStrings.Count ? sharedStrings[index] : "";
        }
        if (type == "b")
        {
            return raw == "1" ? "True" : "False";
        }
        return raw;
    }

    private static XDocument LoadXml(ZipArchive archive, string entryPath)
    {
        var entry = FindEntry(archive, entryPath)
            ?? throw new InvalidDataException($"Excel 工作簿缺少必要文件：{entryPath}");
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    private static ZipArchiveEntry? FindEntry(ZipArchive archive, string path)
    {
        var normalized = path.Replace('\\', '/').TrimStart('/');
        return archive.Entries.FirstOrDefault(entry =>
            string.Equals(entry.FullName.Replace('\\', '/'), normalized, StringComparison.OrdinalIgnoreCase));
    }

    private static string NormalizeZipPath(string basePath, string target)
    {
        var combined = target.StartsWith('/') ? target.TrimStart('/') : $"{basePath}/{target}";
        var parts = new Stack<string>();
        foreach (var part in combined.Replace('\\', '/').Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (part == ".")
            {
                continue;
            }
            if (part == "..")
            {
                if (parts.Count == 0)
                {
                    throw new InvalidDataException("工作表路径越界。");
                }
                parts.Pop();
                continue;
            }
            parts.Push(part);
        }

        return string.Join('/', parts.Reverse());
    }

    private static string? GetColumnName(string? cellReference)
    {
        if (string.IsNullOrWhiteSpace(cellReference))
        {
            return null;
        }
        return new string(cellReference.TakeWhile(char.IsLetter).Select(char.ToUpperInvariant).ToArray());
    }

    private static int? ParsePositiveInt(string? value) =>
        int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) && parsed > 0
            ? parsed
            : null;
}
