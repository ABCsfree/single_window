using System.Xml.Linq;

namespace TradeXmlStudio.Core;

public sealed record OperatorOptions
{
    public string ICCode { get; set; } = "";
    public string CopCode { get; set; } = "";
    public string OperName { get; set; } = "";
}

public sealed record EnterpriseOptions
{
    public string Name { get; set; } = "";
    public string CustomsCode { get; set; } = "";
    public string SocialCreditCode { get; set; } = "";
}

public sealed record TradeXmlOptions
{
    public const long DefaultMaxImageBytes = 2_359_296;

    public static readonly string[] PhotoExtensions = [".jpg", ".jpeg", ".png", ".bmp"];

    public OperatorOptions Operator { get; set; } = new();
    public EnterpriseOptions ExportEnterprise { get; set; } = new();
    public EnterpriseOptions ApplicantEnterprise { get; set; } = new();
    public string SupervisingCustomsCode { get; set; } = "";
    public long MaxImageBytes { get; set; } = DefaultMaxImageBytes;
    public string UploadTypeCode { get; set; } = "F";
    public string P0FilePath { get; set; } = "";
    public bool IncludeP0 { get; set; }
}

public sealed record XmlGenerationRequest(
    string SourceFolderPath,
    string OutputFolderPath,
    string SeqNo,
    string ProBatchNumber,
    string GNo,
    string LotId,
    DateTimeOffset GeneratedAt);

public sealed record XmlGenerationResult(string OutputPath, XDocument Document);

public sealed record EdocSource(
    string FileName,
    string FullPath,
    string BizTypeCode,
    string AttTypeCode,
    bool IsP0);

public sealed class XmlGenerationException(IReadOnlyList<string> errors)
    : Exception(string.Join(Environment.NewLine, errors))
{
    public IReadOnlyList<string> Errors { get; } = errors;
}

public enum BatchFolderMode
{
    SmallFolders,
    SerialSmallFolders,
    SingleBigFolder
}

public sealed record ExcelBatchEntry(int Serial, string LotId);

public sealed record ExcelBatchItemResult(
    int Serial,
    string LotId,
    string PhotoFolderPath,
    int PhotoCount,
    string Status,
    string Detail,
    bool Success);
