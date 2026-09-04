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
        var results = new ExcelBatchItemResult?[entries.Count];
        var prepared = new List<(int Index, ExcelBatchEntry Entry, string PhotoFolder, IReadOnlyList<string> Photos)>();
        var processedLotIds = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 0; index < entries.Count; index++)
        {
            var entry = entries[index];
            var photoFolder = ResolvePhotoFolder(bigFolderPath, entry, mode);
            if (!processedLotIds.Add(entry.LotId.Trim()))
            {
                results[index] = new ExcelBatchItemResult(
                    entry.Serial,
                    entry.LotId,
                    photoFolder,
                    0,
                    "失败",
                    "箱号重复，已跳过。",
                    false);
                continue;
            }

            var photos = GetPhotoPaths(bigFolderPath, entry, mode);
            if (IsSmallFolderMode(mode) && !Directory.Exists(photoFolder))
            {
                results[index] = new ExcelBatchItemResult(
                    entry.Serial, entry.LotId, photoFolder, 0, "失败", "小文件夹不存在。", false);
                continue;
            }
            if (photos.Count != 4)
            {
                results[index] = new ExcelBatchItemResult(
                    entry.Serial,
                    entry.LotId,
                    photoFolder,
                    photos.Count,
                    "失败",
                    $"照片数量必须为 4 张（A1-A4），当前 {photos.Count} 张。",
                    false);
                continue;
            }

            var oversized = options.MaxImageBytes > 0
                ? photos.FirstOrDefault(path => new FileInfo(path).Length > options.MaxImageBytes)
                : null;
            if (oversized is not null)
            {
                results[index] = new ExcelBatchItemResult(
                    entry.Serial,
                    entry.LotId,
                    photoFolder,
                    photos.Count,
                    "失败",
                    $"附件超过大小限制：{Path.GetFileName(oversized)}。",
                    false);
                continue;
            }

            prepared.Add((index, entry, photoFolder, photos));
        }

        if (prepared.Count > 0)
        {
            try
            {
                var batchItems = prepared.Select(item => new BatchXmlGenerationItem(
                    item.Entry.Serial.ToString(CultureInfo.InvariantCulture),
                    item.Entry.LotId,
                    item.PhotoFolder,
                    item.Photos)).ToList();
                var generated = _xmlGenerator.GenerateBatchToFiles(
                    batchItems,
                    outputFolderPath,
                    seqNo,
                    proBatchNumber,
                    generatedAt,
                    options,
                    overwrite);

                foreach (var item in prepared)
                {
                    var lotFiles = generated.LotResults[item.Entry.LotId];
                    var detail = $"4 个 ELBP005：{string.Join(", ", lotFiles.Select(file => Path.GetFileName(file.OutputPath)))}";
                    results[item.Index] = new ExcelBatchItemResult(
                        item.Entry.Serial,
                        item.Entry.LotId,
                        item.PhotoFolder,
                        item.Photos.Count,
                        "成功",
                        detail,
                        true);
                }
            }
            catch (Exception ex)
            {
                var detail = ex.Message.Replace(Environment.NewLine, "；");
                foreach (var item in prepared)
                {
                    results[item.Index] = new ExcelBatchItemResult(
                        item.Entry.Serial,
                        item.Entry.LotId,
                        item.PhotoFolder,
                        item.Photos.Count,
                        "失败",
                        detail,
                        false);
                }
            }
        }

        return results.Select(result => result!).ToList();
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
        var numericSuffix = TrailingDigits(lotId.Trim());
        return long.TryParse(numericPrefix, NumberStyles.Integer, CultureInfo.InvariantCulture, out var prefix)
            && long.TryParse(numericSuffix, NumberStyles.Integer, CultureInfo.InvariantCulture, out var suffix)
            && prefix == suffix;
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
            BatchFolderMode.SmallFolders => Path.Combine(bigFolder, entry.LotId.Trim()),
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
