using System.Globalization;

namespace TradeXmlStudio.Core;

public sealed class ExcelBatchGenerator(
    TradeXmlGenerator? xmlGenerator = null,
    OpenXmlWorkbookReader? workbookReader = null)
{
    private readonly TradeXmlGenerator _xmlGenerator = xmlGenerator ?? new TradeXmlGenerator();
    private readonly OpenXmlWorkbookReader _workbookReader = workbookReader ?? new OpenXmlWorkbookReader();

    public IReadOnlyList<string> ListSheets(string excelPath) => _workbookReader.ListSheets(excelPath);

    public IReadOnlyList<ExcelBatchEntry> ReadEntries(string excelPath, string sheetName) =>
        _workbookReader.ReadEntries(excelPath, sheetName);

    public IReadOnlyList<ExcelBatchItemResult> Preview(
        IReadOnlyList<ExcelBatchEntry> entries,
        string bigFolderPath,
        BatchFolderMode mode)
    {
        return entries.Select(entry => PreviewOne(entry, bigFolderPath, mode)).ToList();
    }

    public IReadOnlyList<ExcelBatchItemResult> RunBatch(
        IReadOnlyList<ExcelBatchEntry> entries,
        string bigFolderPath,
        string outputFolderPath,
        BatchFolderMode mode,
        string seqNo,
        string proBatchNumber,
        DateTimeOffset generatedAt,
        TradeXmlOptions options,
        bool overwrite)
    {
        var results = new List<ExcelBatchItemResult>(entries.Count);

        var processedLotIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var entry in entries)
        {
            if (!processedLotIds.Add(entry.LotId))
            {
                results.Add(new ExcelBatchItemResult(
                    entry.Serial,
                    entry.LotId,
                    ResolvePhotoFolder(bigFolderPath, entry, mode),
                    0,
                    "失败",
                    "箱号重复，已跳过。",
                    false));
                continue;
            }

            results.Add(GenerateEntry(
                entry, bigFolderPath, outputFolderPath, mode, seqNo, proBatchNumber, generatedAt, options, overwrite));
        }

        return results;
    }

    public static IReadOnlyList<string> GetPhotoPaths(
        string bigFolder,
        ExcelBatchEntry entry,
        BatchFolderMode mode)
    {
        var folder = ResolvePhotoFolder(bigFolder, entry, mode);
        if (!Directory.Exists(folder))
        {
            return [];
        }

        return Directory.EnumerateFiles(folder)
            .Where(path => TradeXmlOptions.PhotoExtensions.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase))
            .Where(path => IsSmallFolderMode(mode) || MatchesLotId(Path.GetFileNameWithoutExtension(path), entry.LotId))
            .OrderBy(Path.GetFileName, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public static bool MatchesLotId(string fileNameWithoutExtension, string lotId)
    {
        if (string.IsNullOrWhiteSpace(fileNameWithoutExtension) || string.IsNullOrWhiteSpace(lotId))
        {
            return false;
        }

        var dash = fileNameWithoutExtension.IndexOf('-');
        var numericPrefix = dash < 0 ? fileNameWithoutExtension : fileNameWithoutExtension[..dash];
        var numericSuffix = TrailingDigits(lotId);
        return long.TryParse(numericPrefix, NumberStyles.Integer, CultureInfo.InvariantCulture, out var prefix)
            && long.TryParse(numericSuffix, NumberStyles.Integer, CultureInfo.InvariantCulture, out var suffix)
            && prefix == suffix;
    }

    private ExcelBatchItemResult GenerateEntry(
        ExcelBatchEntry entry,
        string bigFolderPath,
        string outputFolderPath,
        BatchFolderMode mode,
        string seqNo,
        string proBatchNumber,
        DateTimeOffset generatedAt,
        TradeXmlOptions options,
        bool overwrite)
    {
        var photoFolder = ResolvePhotoFolder(bigFolderPath, entry, mode);
        var photos = GetPhotoPaths(bigFolderPath, entry, mode);
        if (IsSmallFolderMode(mode) && !Directory.Exists(photoFolder))
        {
            return new ExcelBatchItemResult(entry.Serial, entry.LotId, photoFolder, 0, "失败", "小文件夹不存在。", false);
        }

        try
        {
            var request = new XmlGenerationRequest(
                photoFolder,
                outputFolderPath,
                seqNo,
                proBatchNumber,
                entry.Serial.ToString(CultureInfo.InvariantCulture),
                entry.LotId,
                generatedAt);
            var generated = _xmlGenerator.GenerateToFilesFromPhotos(photos, request, options, overwrite);
            var detail = $"{generated.Count} 个文件：{string.Join(", ", generated.Select(item => Path.GetFileName(item.OutputPath)))}";
            return new ExcelBatchItemResult(entry.Serial, entry.LotId, photoFolder, photos.Count, "成功", detail, true);
        }
        catch (Exception ex)
        {
            return new ExcelBatchItemResult(entry.Serial, entry.LotId, photoFolder, photos.Count, "失败", ex.Message.Replace(Environment.NewLine, "；"), false);
        }
    }

    private static ExcelBatchItemResult PreviewOne(ExcelBatchEntry entry, string bigFolderPath, BatchFolderMode mode)
    {
        var folder = ResolvePhotoFolder(bigFolderPath, entry, mode);
        var photos = GetPhotoPaths(bigFolderPath, entry, mode);
        var status = IsSmallFolderMode(mode) && !Directory.Exists(folder)
            ? "小文件夹不存在"
            : photos.Count == 4
                ? "就绪"
                : photos.Count == 0
                    ? IsSmallFolderMode(mode) ? "无照片" : "无匹配照片"
                    : $"照片数量异常（{photos.Count}）";
        return new ExcelBatchItemResult(entry.Serial, entry.LotId, folder, photos.Count, status, folder, status == "就绪");
    }

    private static string ResolvePhotoFolder(string bigFolder, ExcelBatchEntry entry, BatchFolderMode mode) =>
        mode switch
        {
            BatchFolderMode.SmallFolders => Path.Combine(bigFolder, entry.LotId),
            BatchFolderMode.SerialSmallFolders => Path.Combine(
                bigFolder,
                entry.Serial.ToString(CultureInfo.InvariantCulture)),
            _ => bigFolder
        };

    private static bool IsSmallFolderMode(BatchFolderMode mode) =>
        mode is BatchFolderMode.SmallFolders or BatchFolderMode.SerialSmallFolders;

    private static string TrailingDigits(string value)
    {
        var index = value.Length;
        while (index > 0 && char.IsDigit(value[index - 1]))
        {
            index--;
        }
        return index == value.Length ? "" : value[index..];
    }
}
