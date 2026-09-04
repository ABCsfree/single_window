using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using TradeXmlStudio.Core;

var testRoot = Path.Combine(Path.GetTempPath(), $"trade-xml-smoke-{Guid.NewGuid():N}");
Directory.CreateDirectory(testRoot);

try
{
    TestLotIdMatching();
    TestWorkbookReading(testRoot);
    TestSerialFolderMode(testRoot);
    TestConfigurationRoundTrip(testRoot);
    TestXmlGeneration(testRoot);
    Console.WriteLine("All smoke tests passed.");
    return 0;
}
finally
{
    if (Directory.Exists(testRoot))
    {
        Directory.Delete(testRoot, true);
    }
}

static void TestLotIdMatching()
{
    Assert(ExcelBatchGenerator.MatchesLotId("123-A1", "BOX00123"), "numeric prefix should match lot-id suffix");
    Assert(ExcelBatchGenerator.MatchesLotId("00123-A4", "箱号123"), "leading zeroes should not affect matching");
    Assert(!ExcelBatchGenerator.MatchesLotId("photo-A1", "BOX00123"), "non-numeric prefix must not match");
    Assert(!ExcelBatchGenerator.MatchesLotId("124-A1", "BOX00123"), "different numbers must not match");
}

static void TestWorkbookReading(string root)
{
    var workbookPath = Path.Combine(root, "batch.xlsx");
    CreateWorkbook(workbookPath);
    var reader = new OpenXmlWorkbookReader();
    var sheets = reader.ListSheets(workbookPath);
    Assert(sheets.SequenceEqual(["批次一"]), "sheet list should be read from workbook.xml");

    var entries = reader.ReadEntries(workbookPath, "批次一");
    Assert(entries.Count == 2, "header and blank lot IDs should be skipped");
    Assert(entries[0] == new ExcelBatchEntry(1, "BOX00123"), "B/C values should map to serial and lot ID");
    Assert(entries[1] == new ExcelBatchEntry(7, "BOX00124"), "row number should be fallback serial");
}

static void TestSerialFolderMode(string root)
{
    var batchRoot = Path.Combine(root, "serial-folders");
    var serialFolder = Path.Combine(batchRoot, "7");
    Directory.CreateDirectory(serialFolder);
    for (var index = 1; index <= 4; index++)
    {
        File.WriteAllBytes(Path.Combine(serialFolder, $"{index}.jpg"), [(byte)'s', (byte)index]);
    }

    var entry = new ExcelBatchEntry(7, "BOX00123");
    var preview = new ExcelBatchGenerator()
        .Preview([entry], batchRoot, BatchFolderMode.SerialSmallFolders)
        .Single();

    Assert(preview.PhotoFolderPath == serialFolder, "serial folder mode should resolve the folder from the Excel serial");
    Assert(preview.PhotoCount == 4 && preview.Status == "就绪", "serial folder mode should find four photos in the serial-named folder");
}

static void TestConfigurationRoundTrip(string root)
{
    var path = Path.Combine(root, "trade-xml-config.json");
    var options = new TradeXmlOptions
    {
        Operator = new OperatorOptions { ICCode = "1000401089921", CopCode = "COP-CODE", OperName = "操作员" },
        ExportEnterprise = new EnterpriseOptions
        {
            Name = "出口企业",
            CustomsCode = "35079606C2",
            SocialCreditCode = "91350781MA2YGJK59J"
        },
        ApplicantEnterprise = new EnterpriseOptions
        {
            Name = "申请单位",
            CustomsCode = "1108919038",
            SocialCreditCode = "91110108MA0012345X"
        },
        SupervisingCustomsCode = "3503"
    };

    ConfigurationStore.Save(path, options);
    var loaded = ConfigurationStore.Load(path);
    Assert(loaded.ExportEnterprise.Name == "出口企业", "exporter configuration should round-trip independently");
    Assert(loaded.ApplicantEnterprise.Name == "申请单位", "applicant configuration should round-trip independently");
    Assert(loaded.SupervisingCustomsCode == "3503", "supervising customs code should round-trip");
}

static void TestXmlGeneration(string root)
{
    var photoFolder = Path.Combine(root, "BOX00123");
    var outputFolder = Path.Combine(root, "output");
    Directory.CreateDirectory(photoFolder);
    Directory.CreateDirectory(outputFolder);
    for (var index = 1; index <= 4; index++)
    {
        File.WriteAllBytes(Path.Combine(photoFolder, $"{index}.jpg"), [(byte)'x', (byte)index]);
    }
    var p0Path = Path.Combine(root, "代理报检委托书.pdf");
    File.WriteAllBytes(p0Path, [(byte)'p', (byte)'0']);

    var options = new TradeXmlOptions
    {
        Operator = new OperatorOptions
        {
            ICCode = "1000401089921",
            CopCode = "91350781MA2YGJK59J",
            OperName = "Tester"
        },
        ExportEnterprise = new EnterpriseOptions
        {
            Name = "福建出口企业有限公司",
            CustomsCode = "35079606C2",
            SocialCreditCode = "91350781MA2YGJK59J"
        },
        ApplicantEnterprise = new EnterpriseOptions
        {
            Name = "北京申请单位有限公司",
            CustomsCode = "1108919038",
            SocialCreditCode = "91110108MA0012345X"
        },
        SupervisingCustomsCode = "3503",
        MaxImageBytes = 2,
        UploadTypeCode = "F",
        IncludeP0 = true,
        P0FilePath = p0Path
    };
    var request = new XmlGenerationRequest(
        photoFolder,
        outputFolder,
        "NN20260804000001",
        "PB202608040001",
        "7",
        "BOX00123",
        DateTimeOffset.Parse("2026-08-04T12:34:56.789-07:00"));
    var results = new TradeXmlGenerator().GenerateToFiles(request, options, false);

    Assert(results.Count == 6, "one ELBP004 and five ELBP005 files should be generated");
    Assert(Path.GetFileName(results[0].OutputPath) == "BOX00123_ELBP004.xml", "the first XML should be the ELBP004 request");

    XNamespace ns = TradeXmlGenerator.ContractNamespace;
    var main = XDocument.Load(results[0].OutputPath);
    Assert(main.Root?.Name == ns + "ELBP004Request", "main root should use the qualified ELBP004 contract");
    Assert(main.Root?.Element(ns + "Head")?.Element(ns + "ProBatchNumber")?.Value == "PB202608040001",
        "ELBP004 should contain the production batch number");
    Assert(main.Root?.Element(ns + "Head")?.Element(ns + "InputEtpsScc")?.Value == "91350781MA2YGJK59J",
        "ELBP004 should contain the independently configured exporter");
    Assert(main.Root?.Element(ns + "Head")?.Element(ns + "AgentScc")?.Value == "91110108MA0012345X",
        "ELBP004 should contain the independently configured applicant");
    Assert(main.Root?.Element(ns + "Lists")?.Element(ns + "List")?.Element(ns + "LotId")?.Value == "BOX00123",
        "ELBP004 should contain the lot list");

    var metadata = main.Root?.Element(ns + "Edocs")?.Elements(ns + "Edoc").ToList() ?? [];
    Assert(metadata.Count == 5, "ELBP004 should reference P0 and A1-A4 metadata");
    Assert(metadata.Select(edoc => edoc.Element(ns + "BizTypeCode")?.Value).SequenceEqual(["P0", "A1", "A2", "A3", "A4"]),
        "ELBP004 attachment order should be P0 followed by A1-A4");
    Assert(metadata.All(edoc => edoc.Element(ns + "EdocID")?.Value.Length == 64),
        "every EdocID should use the updated 64-character rule");
    Assert(metadata[0].Element(ns + "LotId")?.Value == "" && metadata.Skip(1).All(edoc => edoc.Element(ns + "LotId")?.Value == "BOX00123"),
        "P0 should be batch-level while photos should reference the lot ID");

    var attachmentDocuments = results.Skip(1).Select(result => XDocument.Load(result.OutputPath)).ToList();
    Assert(attachmentDocuments.All(document => document.Root?.Name == ns + "ELBP005Request"),
        "attachment roots should use the qualified ELBP005 contract");
    Assert(attachmentDocuments.All(document => document.Root?.Elements(ns + "Edoc").Count() == 1),
        "each ELBP005 should contain exactly one Edoc");
    Assert(attachmentDocuments.Select(document => document.Root!.Element(ns + "Edoc")!.Element(ns + "EdocID")!.Value)
            .SequenceEqual(metadata.Select(edoc => edoc.Element(ns + "EdocID")!.Value)),
        "ELBP004 and ELBP005 should share identical EdocIDs");
    Assert(attachmentDocuments.All(document => document.Root!.Element(ns + "Edoc")!.Element(ns + "FileContent") is not null),
        "each ELBP005 should carry Base64 file content");
    Assert(attachmentDocuments.All(document => document.Root!.Element(ns + "Edoc")!.Element(ns + "SeqNo") is null),
        "ELBP005 should not contain the removed SeqNo field");
    Assert(main.Root?.Element(ns + "OperInfo")?.Element(ns + "Sign")?.Value == "",
        "signature should remain empty for the import client");
}

static void CreateWorkbook(string path)
{
    using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
    AddXml(archive, "xl/workbook.xml", """
        <?xml version="1.0" encoding="UTF-8"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets><sheet name="批次一" sheetId="1" r:id="rId1" /></sheets>
        </workbook>
        """);
    AddXml(archive, "xl/_rels/workbook.xml.rels", """
        <?xml version="1.0" encoding="UTF-8"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml" />
        </Relationships>
        """);
    AddXml(archive, "xl/sharedStrings.xml", """
        <?xml version="1.0" encoding="UTF-8"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">
          <si><t>序号</t></si><si><t>箱号</t></si><si><t>BOX00123</t></si><si><t>BOX00124</t></si>
        </sst>
        """);
    AddXml(archive, "xl/worksheets/sheet1.xml", """
        <?xml version="1.0" encoding="UTF-8"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
          <row r="1"><c r="B1" t="s"><v>0</v></c><c r="C1" t="s"><v>1</v></c></row>
          <row r="2"><c r="B2"><v>1</v></c><c r="C2" t="s"><v>2</v></c></row>
          <row r="3"><c r="B3"><v>2</v></c><c r="C3" t="inlineStr"><is><t></t></is></c></row>
          <row r="7"><c r="B7" t="inlineStr"><is><t>不是数字</t></is></c><c r="C7" t="s"><v>3</v></c></row>
        </sheetData></worksheet>
        """);
}

static void AddXml(ZipArchive archive, string name, string content)
{
    var entry = archive.CreateEntry(name);
    using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
    writer.Write(content);
}

static void Assert(bool condition, string message)
{
    if (!condition)
    {
        throw new InvalidOperationException($"Assertion failed: {message}");
    }
    Console.WriteLine($"PASS: {message}");
}
