using System.Globalization;
using System.Text;
using System.Xml.Linq;

namespace TradeXmlStudio.Core;

public sealed class TradeXmlGenerator
{
    public const string ContractNamespace = "http://www.w3.org/2000/09/xmldsig#";

    private static readonly XNamespace Ns = ContractNamespace;
    private static readonly UTF8Encoding Utf8WithoutBom = new(false);
    private static readonly string[] PhotoBizTypeCodes = ["A1", "A2", "A3", "A4"];
    private static int _clientSequenceCounter = Environment.TickCount & int.MaxValue;
    private static long _edocSequenceCounter = DateTime.UtcNow.Ticks & long.MaxValue;

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
        var document = BuildElbp005(source, edocId, generatedAt, options);
        return WriteDocument(document, outputFolderPath, "0_P0_ELBP005.xml", overwrite);
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
        var ids = sources.Select(source => CreateEdocId(source, request.GeneratedAt, options)).ToList();
        var prefix = SanitizeFileName(request.LotId.Trim());
        var pending = new List<(XDocument Document, string FileName)>
        {
            (BuildElbp004(sources, ids, request, options), $"{prefix}_ELBP004.xml")
        };

        for (var index = 0; index < sources.Count; index++)
        {
            var source = sources[index];
            pending.Add((
                BuildElbp005(source, ids[index], request.GeneratedAt, options),
                $"{prefix}_{source.BizTypeCode}_ELBP005.xml"));
        }

        var outputPaths = pending
            .Select(item => Path.Combine(request.OutputFolderPath, item.FileName))
            .ToList();
        if (!overwrite)
        {
            var existing = outputPaths.FirstOrDefault(File.Exists);
            if (existing is not null)
            {
                throw new IOException($"输出文件已存在：{existing}");
            }
        }

        return pending
            .Select(item => WriteDocument(item.Document, request.OutputFolderPath, item.FileName, overwrite))
            .ToList();
    }

    private static XDocument BuildElbp004(
        IReadOnlyList<EdocSource> sources,
        IReadOnlyList<string> edocIds,
        XmlGenerationRequest request,
        TradeXmlOptions options)
    {
        var edocs = sources.Select((source, index) =>
            new XElement(Ns + "Edoc",
                new XElement(Ns + "LotId", source.IsP0 ? "" : request.LotId.Trim()),
                new XElement(Ns + "EdocID", edocIds[index]),
                new XElement(Ns + "BizTypeCode", source.BizTypeCode),
                new XElement(Ns + "AttFmtTypeCode", "US"),
                new XElement(Ns + "AttTypeCode", source.AttTypeCode),
                new XElement(Ns + "AttEdocName", source.FileName)));

        var root = new XElement(Ns + "ELBP004Request",
            BuildOperatorInfo("ELBP004", request.GeneratedAt, options, includeOperType: true),
            new XElement(Ns + "Head",
                new XElement(Ns + "NnNo", request.SeqNo.Trim()),
                new XElement(Ns + "ProBatchNumber", request.ProBatchNumber.Trim()),
                new XElement(Ns + "InputEtpsName", options.ExportEnterprise.Name.Trim()),
                new XElement(Ns + "InputEtpsCode", options.ExportEnterprise.CustomsCode.Trim()),
                new XElement(Ns + "InputEtpsScc", options.ExportEnterprise.SocialCreditCode.Trim()),
                new XElement(Ns + "AgentName", options.ApplicantEnterprise.Name.Trim()),
                new XElement(Ns + "AgentCode", options.ApplicantEnterprise.CustomsCode.Trim()),
                new XElement(Ns + "AgentScc", options.ApplicantEnterprise.SocialCreditCode.Trim()),
                new XElement(Ns + "Note", "")),
            new XElement(Ns + "Lists",
                new XElement(Ns + "List",
                    new XElement(Ns + "GNo", request.GNo.Trim()),
                    new XElement(Ns + "LotId", request.LotId.Trim()))),
            new XElement(Ns + "Edocs", edocs));
        return new XDocument(new XDeclaration("1.0", "UTF-8", null), root);
    }

    private static XDocument BuildElbp005(
        EdocSource source,
        string edocId,
        DateTimeOffset generatedAt,
        TradeXmlOptions options)
    {
        var root = new XElement(Ns + "ELBP005Request",
            BuildOperatorInfo("ELBP005", generatedAt, options, includeOperType: false),
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
        DateTimeOffset generatedAt,
        TradeXmlOptions options,
        bool includeOperType)
    {
        var operInfo = new XElement(Ns + "OperInfo");
        if (includeOperType)
        {
            operInfo.Add(new XElement(Ns + "OperType", "C"));
        }

        operInfo.Add(
            new XElement(Ns + "MessageType", messageType),
            new XElement(Ns + "Version", "1.0"),
            new XElement(Ns + "ICCode", options.Operator.ICCode.Trim()),
            string.IsNullOrWhiteSpace(options.Operator.CopCode)
                ? null
                : new XElement(Ns + "CopCode", options.Operator.CopCode.Trim()),
            new XElement(Ns + "OperName", options.Operator.OperName.Trim()),
            new XElement(Ns + "ClientSeqNo", CreateClientSequenceNo(generatedAt)),
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
            .ToString("D23", CultureInfo.InvariantCulture);
        return options.ExportEnterprise.SocialCreditCode.Trim()
            + options.SupervisingCustomsCode.Trim()
            + source.BizTypeCode.Trim()
            + generatedAt.ToString("yyyyMMddHHmmssfff", CultureInfo.InvariantCulture)
            + serial;
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
        ValidateRequiredLength(request.SeqNo, "通知编号", 18, errors);
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
}
