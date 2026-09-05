using System.Globalization;
using System.Text;
using System.Xml.Linq;

namespace TradeXmlStudio.Core;

public sealed class TradeXmlGenerator
{
    // The accepted reference messages use unqualified element names. The old
    // value was the XML Signature namespace, which does not belong on ELBP004/5.
    public const string ContractNamespace = "";

    private static readonly XNamespace Ns = ContractNamespace;
    private static readonly UTF8Encoding Utf8WithoutBom = new(false);
    private static readonly string[] PhotoBizTypeCodes = ["A1", "A2", "A3", "A4"];
    private static readonly string[] BatchPhotoBizTypeCodes = ["A1", "A2", "A3", "A4", "B1"];
    private static int _clientSequenceCounter = Environment.TickCount & int.MaxValue;
    private static long _edocSequenceCounter = DateTime.UtcNow.Ticks & long.MaxValue;
    private static long _fileTimestampTicks = DateTime.Now.Ticks;

    public IReadOnlyList<XmlGenerationResult> GenerateToFiles(
        XmlGenerationRequest request,
        TradeXmlOptions options,
        bool overwrite)
    {
        var errors = ValidateFolderRequest(request, options);
        if (errors.Count > 0)
        {
            throw new XmlGenerationException(errors);
        }

        var sources = ScanEdocs(request.SourceFolderPath, options.P0FilePath, options.IncludeP0);
        return GeneratePackage(sources, request, options, overwrite);
    }

    public IReadOnlyList<XmlGenerationResult> GenerateToFilesFromPhotos(
        IReadOnlyList<string> photoPaths,
        XmlGenerationRequest request,
        TradeXmlOptions options,
        bool overwrite)
    {
        var errors = ValidatePhotoRequest(request, options, photoPaths);
        if (errors.Count > 0)
        {
            throw new XmlGenerationException(errors);
        }

        var sources = photoPaths
            .OrderBy(Path.GetFileName, StringComparer.OrdinalIgnoreCase)
            .Select((path, index) => new EdocSource(
                Path.GetFileName(path),
                Path.GetFullPath(path),
                PhotoBizTypeCodes[index],
                GetAttachmentType(path),
                false))
            .ToList();

        if (options.IncludeP0)
        {
            var p0Path = Path.GetFullPath(options.P0FilePath);
            sources.Insert(0, new EdocSource(
                Path.GetFileName(p0Path),
                p0Path,
                "P0",
                GetAttachmentType(p0Path),
                true));
        }

        return GeneratePackage(sources, request, options, overwrite);
    }

    public BatchXmlGenerationResult GenerateBatchToFiles(
        IReadOnlyList<BatchXmlGenerationItem> items,
        string outputFolderPath,
        string seqNo,
        string proBatchNumber,
        DateTimeOffset generatedAt,
        TradeXmlOptions options,
        bool overwrite,
        int expectedPhotoCount = 4)
    {
        var errors = ValidateBatchRequest(
            items,
            outputFolderPath,
            seqNo,
            proBatchNumber,
            options,
            expectedPhotoCount);
        if (errors.Count > 0)
        {
            throw new XmlGenerationException(errors);
        }

        var photoBizTypeCodes = GetBatchPhotoBizTypeCodes(expectedPhotoCount);
        var batchEdocs = new List<BatchEdoc>(items.Count * photoBizTypeCodes.Count + 1);
        foreach (var item in items)
        {
            var photos = item.PhotoPaths.OrderBy(Path.GetFileName, StringComparer.OrdinalIgnoreCase).ToList();
            for (var index = 0; index < photos.Count; index++)
            {
                var path = Path.GetFullPath(photos[index]);
                var source = new EdocSource(
                    Path.GetFileName(path),
                    path,
                    photoBizTypeCodes[index],
                    GetAttachmentType(path),
                    false);
                batchEdocs.Add(new BatchEdoc(item.LotId, source, CreateEdocId(source, generatedAt, options)));
            }
        }

        if (options.IncludeP0)
        {
            var p0Path = Path.GetFullPath(options.P0FilePath);
            var source = new EdocSource(Path.GetFileName(p0Path), p0Path, "P0", GetAttachmentType(p0Path), true);
            batchEdocs.Add(new BatchEdoc("", source, CreateEdocId(source, generatedAt, options)));
        }

        var lists = items.Select(item => new BatchList(item.GNo.Trim(), item.LotId)).ToList();
        var sharedPending = new List<(XDocument Document, string FileName)>();
        var lotPending = items.ToDictionary(
            item => item.LotId,
            _ => new List<(XDocument Document, string FileName)>(),
            StringComparer.Ordinal);

        var elbp004ClientSeqNo = CreateClientSequenceNo(generatedAt);
        sharedPending.Add((
            BuildElbp004(lists, batchEdocs, seqNo, proBatchNumber, options, elbp004ClientSeqNo),
            $"ELBP004_{elbp004ClientSeqNo}_{CreateFileTimestamp()}.xml"));

        foreach (var edoc in batchEdocs)
        {
            var document = BuildElbp005(
                edoc.Source,
                edoc.EdocId,
                options,
                CreateClientSequenceNo(generatedAt));
            if (edoc.Source.IsP0)
            {
                sharedPending.Add((document, $"0_P0_{CreateFileTimestamp()}.xml"));
            }
            else
            {
                lotPending[edoc.LotId].Add((
                    document,
                    $"{SanitizeFileName(edoc.LotId.Trim())}_{edoc.Source.BizTypeCode}_{CreateFileTimestamp()}.xml"));
            }
        }

        var allPending = sharedPending.Concat(lotPending.Values.SelectMany(value => value)).ToList();
        EnsureCanWrite(allPending, outputFolderPath, overwrite);

        var sharedResults = sharedPending
            .Select(item => WriteDocument(item.Document, outputFolderPath, item.FileName, overwrite))
            .ToList();
        var lotResults = lotPending.ToDictionary(
            pair => pair.Key,
            pair => (IReadOnlyList<XmlGenerationResult>)pair.Value
                .Select(item => WriteDocument(item.Document, outputFolderPath, item.FileName, overwrite))
                .ToList(),
            StringComparer.Ordinal);
        return new BatchXmlGenerationResult(sharedResults, lotResults);
    }

    public static IReadOnlyList<string> GetBatchPhotoBizTypeCodes(int photoCount) =>
        photoCount switch
        {
            3 => BatchPhotoBizTypeCodes[..3],
            4 => BatchPhotoBizTypeCodes[..4],
            5 => BatchPhotoBizTypeCodes[..5],
            _ => throw new ArgumentOutOfRangeException(
                nameof(photoCount),
                photoCount,
                "批量图片张数只支持 3、4 或 5 张。")
        };

    public XmlGenerationResult GenerateP0ToFile(
        TradeXmlOptions options,
        string outputFolderPath,
        string seqNo,
        DateTimeOffset generatedAt,
        bool overwrite)
    {
        var errors = new List<string>();
        ValidateOutputFolder(outputFolderPath, errors);
        ValidateOperatorAndEnterprises(options, errors);
        ValidateP0(options, errors);
        if (errors.Count > 0)
        {
            throw new XmlGenerationException(errors);
        }

        var fullPath = Path.GetFullPath(options.P0FilePath);
        ValidateAttachmentSizes([fullPath], options, errors);
        if (errors.Count > 0)
        {
            throw new XmlGenerationException(errors);
        }

        var source = new EdocSource(Path.GetFileName(fullPath), fullPath, "P0", GetAttachmentType(fullPath), true);
        var edocId = CreateEdocId(source, generatedAt, options);
        var document = BuildElbp005(source, edocId, options, CreateClientSequenceNo(generatedAt));
        return WriteDocument(document, outputFolderPath, $"0_P0_{CreateFileTimestamp()}.xml", overwrite);
    }

    public static IReadOnlyList<EdocSource> ScanEdocs(string sourceFolder, string p0FilePath, bool includeP0)
    {
        var result = new List<EdocSource>();
        string? normalizedP0 = null;

        if (includeP0 && !string.IsNullOrWhiteSpace(p0FilePath))
        {
            normalizedP0 = TryGetFullPath(p0FilePath) ?? p0FilePath;
            result.Add(new EdocSource(
                Path.GetFileName(p0FilePath),
                normalizedP0,
                "P0",
                GetAttachmentType(p0FilePath),
                true));
        }

        if (!Directory.Exists(sourceFolder))
        {
            return result;
        }

        var photos = Directory.EnumerateFiles(sourceFolder)
            .Where(IsSupportedPhoto)
            .OrderBy(Path.GetFileName, StringComparer.OrdinalIgnoreCase);
        var index = 0;
        foreach (var photo in photos)
        {
            var fullPath = TryGetFullPath(photo) ?? photo;
            if (normalizedP0 is not null && string.Equals(fullPath, normalizedP0, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            result.Add(new EdocSource(
                Path.GetFileName(photo),
                fullPath,
                index < PhotoBizTypeCodes.Length ? PhotoBizTypeCodes[index] : "",
                GetAttachmentType(photo),
                false));
            index++;
        }

        return result;
    }

    public static string ToXmlText(XDocument document)
    {
        var root = document.Root ?? throw new InvalidOperationException("XML 缺少根节点。");
        return $"<?xml version=\"1.0\" encoding=\"UTF-8\"?>{Environment.NewLine}{root}";
    }

    private static IReadOnlyList<XmlGenerationResult> GeneratePackage(
        IReadOnlyList<EdocSource> sources,
        XmlGenerationRequest request,
        TradeXmlOptions options,
        bool overwrite)
    {
        var orderedSources = sources.OrderBy(source => source.IsP0).ToList();
        var batchEdocs = orderedSources
            .Select(source => new BatchEdoc(
                source.IsP0 ? "" : request.LotId.Trim(),
                source,
                CreateEdocId(source, request.GeneratedAt, options)))
            .ToList();
        var clientSeqNo = CreateClientSequenceNo(request.GeneratedAt);
        var pending = new List<(XDocument Document, string FileName)>
        {
            (BuildElbp004(
                [new BatchList(request.GNo.Trim(), request.LotId.Trim())],
                batchEdocs,
                request.SeqNo,
                request.ProBatchNumber,
                options,
                clientSeqNo),
                $"ELBP004_{clientSeqNo}_{CreateFileTimestamp()}.xml")
        };

        foreach (var edoc in batchEdocs)
        {
            pending.Add((
                BuildElbp005(
                    edoc.Source,
                    edoc.EdocId,
                    options,
                    CreateClientSequenceNo(request.GeneratedAt)),
                edoc.Source.IsP0
                    ? $"0_P0_{CreateFileTimestamp()}.xml"
                    : $"{SanitizeFileName(request.LotId.Trim())}_{edoc.Source.BizTypeCode}_{CreateFileTimestamp()}.xml"));
        }

        EnsureCanWrite(pending, request.OutputFolderPath, overwrite);

        return pending
            .Select(item => WriteDocument(item.Document, request.OutputFolderPath, item.FileName, overwrite))
            .ToList();
    }

    private static XDocument BuildElbp004(
        IReadOnlyList<BatchList> lists,
        IReadOnlyList<BatchEdoc> batchEdocs,
        string seqNo,
        string proBatchNumber,
        TradeXmlOptions options,
        string clientSeqNo)
    {
        var edocs = batchEdocs.Select(edoc =>
            new XElement(Ns + "Edoc",
                new XElement(Ns + "LotId", edoc.LotId),
                new XElement(Ns + "EdocID", edoc.EdocId),
                new XElement(Ns + "BizTypeCode", edoc.Source.BizTypeCode),
                new XElement(Ns + "AttFmtTypeCode", "US"),
                new XElement(Ns + "AttTypeCode", edoc.Source.AttTypeCode),
                new XElement(Ns + "AttEdocName", edoc.Source.FileName)));

        var root = new XElement(Ns + "ELBP004Request",
            BuildOperatorInfo("ELBP004", options, clientSeqNo, includeOperType: true),
            new XElement(Ns + "Head",
                new XElement(Ns + "NnNo", seqNo.Trim()),
                new XElement(Ns + "ProBatchNumber", proBatchNumber.Trim()),
                new XElement(Ns + "InputEtpsName", options.ExportEnterprise.Name.Trim()),
                new XElement(Ns + "InputEtpsCode", options.ExportEnterprise.CustomsCode.Trim()),
                new XElement(Ns + "InputEtpsScc", options.ExportEnterprise.SocialCreditCode.Trim()),
                new XElement(Ns + "AgentName", options.ApplicantEnterprise.Name.Trim()),
                new XElement(Ns + "AgentCode", options.ApplicantEnterprise.CustomsCode.Trim()),
                new XElement(Ns + "AgentScc", options.ApplicantEnterprise.SocialCreditCode.Trim()),
                new XElement(Ns + "Note", "")),
            new XElement(Ns + "Lists", lists.Select(list =>
                new XElement(Ns + "List",
                    new XElement(Ns + "GNo", list.GNo),
                    new XElement(Ns + "LotId", list.LotId)))),
            new XElement(Ns + "Edocs", edocs));
        return new XDocument(new XDeclaration("1.0", "UTF-8", null), root);
    }

    private static XDocument BuildElbp005(
        EdocSource source,
        string edocId,
        TradeXmlOptions options,
        string clientSeqNo)
    {
        var root = new XElement(Ns + "ELBP005Request",
            BuildOperatorInfo("ELBP005", options, clientSeqNo, includeOperType: false),
            new XElement(Ns + "Edoc",
                new XElement(Ns + "EdocID", edocId),
                new XElement(Ns + "BizTypeCode", source.BizTypeCode),
                new XElement(Ns + "AttTypeCode", source.AttTypeCode),
                new XElement(Ns + "AttFmtTypeCode", "US"),
                new XElement(Ns + "AttEdocName", source.FileName),
                new XElement(Ns + "UploadTypeCode", options.UploadTypeCode.Trim()),
                new XElement(Ns + "FileContent", Convert.ToBase64String(File.ReadAllBytes(source.FullPath)))));
        return new XDocument(new XDeclaration("1.0", "UTF-8", null), root);
    }

    private static XElement BuildOperatorInfo(
        string messageType,
        TradeXmlOptions options,
        string clientSeqNo,
        bool includeOperType)
    {
        var operInfo = new XElement(Ns + "OperInfo");
        if (includeOperType)
        {
            operInfo.Add(new XElement(Ns + "OperType", options.InformationEntryOperType.Trim()));
        }

        operInfo.Add(
            new XElement(Ns + "MessageType", messageType),
            new XElement(Ns + "Version", "1.0"),
            new XElement(Ns + "ICCode", options.Operator.ICCode.Trim()),
            new XElement(Ns + "CopCode", includeOperType ? options.Operator.CopCode.Trim() : ""),
            new XElement(Ns + "OperName", options.Operator.OperName.Trim()),
            new XElement(Ns + "ClientSeqNo", clientSeqNo),
            // 单一窗口导入客户端会在发送前填写签名及签名时间。
            new XElement(Ns + "Sign", ""),
            new XElement(Ns + "SignDate", ""),
            new XElement(Ns + "Note", ""));
        return operInfo;
    }

    private static string CreateClientSequenceNo(DateTimeOffset generatedAt) =>
        generatedAt.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture)
        + NextCounter(ref _clientSequenceCounter, 10_000, 4);

    private static string CreateEdocId(
        EdocSource source,
        DateTimeOffset generatedAt,
        TradeXmlOptions options)
    {
        var serial = (Interlocked.Increment(ref _edocSequenceCounter) & long.MaxValue)
            .ToString("D20", CultureInfo.InvariantCulture);
        return options.ExportEnterprise.SocialCreditCode.Trim()
            + options.SupervisingCustomsCode.Trim()
            + source.BizTypeCode.Trim()
            + generatedAt.ToString("yyyyMMddHHmmssfff", CultureInfo.InvariantCulture)
            + serial;
    }

    private static string CreateFileTimestamp()
    {
        while (true)
        {
            var previous = Interlocked.Read(ref _fileTimestampTicks);
            var now = DateTime.Now.Ticks;
            var next = Math.Max(now, previous + 1);
            if (Interlocked.CompareExchange(ref _fileTimestampTicks, next, previous) == previous)
            {
                return new DateTime(next, DateTimeKind.Local)
                    .ToString("yyyyMMddHHmmssfffffff", CultureInfo.InvariantCulture);
            }
        }
    }

    private static void EnsureCanWrite(
        IReadOnlyList<(XDocument Document, string FileName)> pending,
        string outputFolderPath,
        bool overwrite)
    {
        if (overwrite)
        {
            return;
        }

        var existing = pending
            .Select(item => Path.Combine(outputFolderPath, item.FileName))
            .FirstOrDefault(File.Exists);
        if (existing is not null)
        {
            throw new IOException($"输出文件已存在：{existing}");
        }
    }

    private static XmlGenerationResult WriteDocument(
        XDocument document,
        string outputFolder,
        string fileName,
        bool overwrite)
    {
        var outputPath = Path.Combine(outputFolder, fileName);
        WriteAtomically(outputPath, ToXmlText(document), overwrite);
        return new XmlGenerationResult(outputPath, document);
    }

    private static void WriteAtomically(string outputPath, string content, bool overwrite)
    {
        if (File.Exists(outputPath) && !overwrite)
        {
            throw new IOException($"输出文件已存在：{outputPath}");
        }

        var outputDirectory = Path.GetDirectoryName(outputPath)
            ?? throw new InvalidOperationException("输出路径无效。");
        var tempDirectory = Path.Combine(outputDirectory, ".trade-xml-tmp");
        Directory.CreateDirectory(tempDirectory);
        var tempPath = Path.Combine(tempDirectory, $"{Path.GetFileName(outputPath)}.{Guid.NewGuid():N}.tmp");

        try
        {
            File.WriteAllText(tempPath, content, Utf8WithoutBom);
            File.Move(tempPath, outputPath, true);
        }
        finally
        {
            if (File.Exists(tempPath))
            {
                File.Delete(tempPath);
            }
        }
    }

    private static List<string> ValidateFolderRequest(XmlGenerationRequest request, TradeXmlOptions options)
    {
        var errors = new List<string>();
        if (string.IsNullOrWhiteSpace(request.SourceFolderPath))
        {
            errors.Add("照片文件夹不能为空。");
        }
        else if (!Directory.Exists(request.SourceFolderPath))
        {
            errors.Add($"照片文件夹不存在：{request.SourceFolderPath}");
        }

        ValidateOutputFolder(request.OutputFolderPath, errors);
        ValidateBusinessFields(request, options, errors);
        if (options.IncludeP0)
        {
            ValidateP0(options, errors);
        }

        if (Directory.Exists(request.SourceFolderPath))
        {
            var sources = ScanEdocs(request.SourceFolderPath, options.P0FilePath, options.IncludeP0);
            var photos = sources.Where(source => !source.IsP0).ToList();
            ValidatePhotoCount(photos.Count, errors);
            ValidateAttachmentSizes(sources.Where(source => File.Exists(source.FullPath)).Select(source => source.FullPath), options, errors);
        }

        return errors;
    }

    private static List<string> ValidatePhotoRequest(
        XmlGenerationRequest request,
        TradeXmlOptions options,
        IReadOnlyList<string> photos)
    {
        var errors = new List<string>();
        ValidateOutputFolder(request.OutputFolderPath, errors);
        ValidateBusinessFields(request, options, errors);
        if (options.IncludeP0)
        {
            ValidateP0(options, errors);
        }
        ValidatePhotoCount(photos.Count, errors);
        foreach (var photo in photos.Where(photo => !File.Exists(photo)))
        {
            errors.Add($"照片不存在：{photo}");
        }

        var attachments = photos.Where(File.Exists).ToList();
        if (options.IncludeP0 && File.Exists(options.P0FilePath))
        {
            attachments.Add(options.P0FilePath);
        }
        ValidateAttachmentSizes(attachments, options, errors);
        return errors;
    }

    private static List<string> ValidateBatchRequest(
        IReadOnlyList<BatchXmlGenerationItem> items,
        string outputFolderPath,
        string seqNo,
        string proBatchNumber,
        TradeXmlOptions options,
        int expectedPhotoCount)
    {
        var errors = new List<string>();
        ValidateOutputFolder(outputFolderPath, errors);
        ValidateRequiredLength(seqNo, "通知编号", 50, errors);
        ValidateRequiredLength(proBatchNumber, "生产批次号", 50, errors);
        ValidateOperatorAndEnterprises(options, errors);

        var hasSupportedPhotoCount = expectedPhotoCount is 3 or 4 or 5;
        if (!hasSupportedPhotoCount)
        {
            errors.Add("批量图片张数只支持 3、4 或 5 张。");
        }

        if (items.Count == 0)
        {
            errors.Add("没有可生成的箱号。");
        }

        foreach (var duplicate in items.GroupBy(item => item.LotId.Trim(), StringComparer.Ordinal).Where(group => group.Count() > 1))
        {
            errors.Add($"箱号重复：{duplicate.Key}");
        }

        if (options.IncludeP0)
        {
            ValidateP0(options, errors);
        }

        foreach (var item in items)
        {
            ValidateRequiredLength(item.GNo, $"箱号 {item.LotId} 的商品项号", 18, errors);
            if (!string.IsNullOrWhiteSpace(item.GNo) && !item.GNo.Trim().All(char.IsDigit))
            {
                errors.Add($"箱号 {item.LotId} 的商品项号必须为数字。");
            }
            ValidateRequiredLength(item.LotId, "箱号", 255, errors);
            if (hasSupportedPhotoCount)
            {
                ValidateBatchPhotoCount(item.PhotoPaths.Count, expectedPhotoCount, errors);
            }
            foreach (var photo in item.PhotoPaths.Where(photo => !File.Exists(photo)))
            {
                errors.Add($"照片不存在：{photo}");
            }
        }

        var attachments = items.SelectMany(item => item.PhotoPaths).Where(File.Exists).ToList();
        if (options.IncludeP0 && File.Exists(options.P0FilePath))
        {
            attachments.Add(options.P0FilePath);
        }
        ValidateAttachmentSizes(attachments, options, errors);
        return errors;
    }

    private static void ValidateBatchPhotoCount(int count, int expectedPhotoCount, List<string> errors)
    {
        var bizTypeCodes = string.Join("、", GetBatchPhotoBizTypeCodes(expectedPhotoCount));
        if (count == 0)
        {
            errors.Add($"没有匹配到照片（需要 {expectedPhotoCount} 张：{bizTypeCodes}）。");
        }
        else if (count != expectedPhotoCount)
        {
            errors.Add($"照片数量必须为 {expectedPhotoCount} 张（{bizTypeCodes}），当前 {count} 张。");
        }
    }

    private static void ValidatePhotoCount(int count, List<string> errors)
    {
        if (count == 0)
        {
            errors.Add("没有匹配到照片（需要 4 张 A1-A4）。");
        }
        else if (count != 4)
        {
            errors.Add($"照片数量必须为 4 张（A1-A4），当前 {count} 张。");
        }
    }

    private static void ValidateBusinessFields(
        XmlGenerationRequest request,
        TradeXmlOptions options,
        List<string> errors)
    {
        ValidateRequiredLength(request.SeqNo, "通知编号", 50, errors);
        ValidateRequiredLength(request.ProBatchNumber, "生产批次号", 50, errors);
        ValidateRequiredLength(request.GNo, "商品项号", 18, errors);
        if (!string.IsNullOrWhiteSpace(request.GNo) && !request.GNo.Trim().All(char.IsDigit))
        {
            errors.Add("商品项号必须为数字。");
        }
        ValidateRequiredLength(request.LotId, "箱号", 255, errors);
        ValidateOperatorAndEnterprises(options, errors);
    }

    private static void ValidateOperatorAndEnterprises(TradeXmlOptions options, List<string> errors)
    {
        ValidateRequiredLength(options.Operator.ICCode, "操作人代码", 13, errors);
        ValidateOptionalLength(options.Operator.CopCode, "企业组织机构代码", 18, errors);
        ValidateRequiredLength(options.Operator.OperName, "操作人姓名", 30, errors);

        ValidateEnterprise(options.ExportEnterprise, "出口企业", errors);
        ValidateEnterprise(options.ApplicantEnterprise, "申请单位", errors);
        ValidateExactLength(options.SupervisingCustomsCode, "4位主管海关代码", 4, errors);

        if (string.IsNullOrWhiteSpace(options.InformationEntryOperType)
            || options.InformationEntryOperType.Trim() is not ("C" or "G"))
        {
            errors.Add("信息补录导入方式必须为 C（申报）或 G（暂存）。");
        }
        if (string.IsNullOrWhiteSpace(options.UploadTypeCode)
            || options.UploadTypeCode.Trim() is not ("F" or "P"))
        {
            errors.Add("上传类型代码必须为 F（首次上传）或 P（重传/补传）。");
        }
        if (options.MaxImageBytes <= 0)
        {
            errors.Add("单个附件大小限制必须大于 0。");
        }
    }

    private static void ValidateEnterprise(EnterpriseOptions enterprise, string label, List<string> errors)
    {
        ValidateRequiredLength(enterprise.Name, $"{label}名称", 70, errors);
        ValidateRequiredLength(enterprise.CustomsCode, $"{label}海关十位编码", 10, errors);
        ValidateExactLength(enterprise.SocialCreditCode, $"{label}统一社会信用代码", 18, errors);
    }

    private static void ValidateP0(TradeXmlOptions options, List<string> errors)
    {
        if (string.IsNullOrWhiteSpace(options.P0FilePath))
        {
            errors.Add("包含代理委托协议（P0）时必须选择 P0 文件。");
        }
        else if (!File.Exists(options.P0FilePath))
        {
            errors.Add($"P0 文件不存在：{options.P0FilePath}");
        }
    }

    private static void ValidateOutputFolder(string outputFolder, List<string> errors)
    {
        if (string.IsNullOrWhiteSpace(outputFolder))
        {
            errors.Add("输出目录不能为空。");
        }
        else if (!Directory.Exists(outputFolder))
        {
            errors.Add($"输出目录不存在：{outputFolder}");
        }
    }

    private static void ValidateAttachmentSizes(
        IEnumerable<string> paths,
        TradeXmlOptions options,
        List<string> errors)
    {
        if (options.MaxImageBytes <= 0)
        {
            return;
        }

        foreach (var path in paths)
        {
            var rawLength = new FileInfo(path).Length;
            if (rawLength > options.MaxImageBytes)
            {
                errors.Add(
                    $"附件超过大小限制：{Path.GetFileName(path)}，原文件 "
                    + $"{ToMegabytes(rawLength):0.00} MB，限制 {ToMegabytes(options.MaxImageBytes):0.00} MB。");
            }
        }
    }

    private static void ValidateRequiredLength(string value, string label, int maxLength, List<string> errors)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            errors.Add($"{label}不能为空。");
        }
        else if (value.Trim().Length > maxLength)
        {
            errors.Add($"{label}长度不能超过 {maxLength}。");
        }
    }

    private static void ValidateOptionalLength(string value, string label, int maxLength, List<string> errors)
    {
        if (!string.IsNullOrWhiteSpace(value) && value.Trim().Length > maxLength)
        {
            errors.Add($"{label}长度不能超过 {maxLength}。");
        }
    }

    private static void ValidateExactLength(string value, string label, int length, List<string> errors)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            errors.Add($"{label}不能为空。");
        }
        else if (value.Trim().Length != length)
        {
            errors.Add($"{label}必须为 {length} 位。");
        }
    }

    private static bool IsSupportedPhoto(string path) =>
        TradeXmlOptions.PhotoExtensions.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase);

    private static string GetAttachmentType(string path) => Path.GetExtension(path).TrimStart('.').ToLowerInvariant();

    private static string? TryGetFullPath(string path)
    {
        try
        {
            return Path.GetFullPath(path);
        }
        catch
        {
            return null;
        }
    }

    private static string SanitizeFileName(string value)
    {
        var invalid = Path.GetInvalidFileNameChars();
        return new string(value.Select(character => invalid.Contains(character) ? '_' : character).ToArray());
    }

    private static string NextCounter(ref int counter, int modulus, int digits) =>
        ((Interlocked.Increment(ref counter) & int.MaxValue) % modulus)
        .ToString($"D{digits}", CultureInfo.InvariantCulture);

    private static double ToMegabytes(long bytes) => bytes / 1024d / 1024d;

    private sealed record BatchList(string GNo, string LotId);

    private sealed record BatchEdoc(string LotId, EdocSource Source, string EdocId);
}
