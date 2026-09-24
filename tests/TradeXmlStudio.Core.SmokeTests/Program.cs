using System.IO.Compression;
using System.IO;
using System.Windows.Media;
using System.Windows.Media.Imaging;
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
    TestExportEnterpriseProfiles(testRoot);
    TestExportEnterpriseProfileDeletion(testRoot);
    TestXmlGeneration(testRoot);
    TestBatchXmlGeneration(testRoot);
    TestAttachmentConversion(testRoot, args.FirstOrDefault());
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
    Assert(ExcelBatchGenerator.MatchesLotId("123-A1", "  BOX00123  "), "lot-id matching should ignore surrounding spaces");
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
    Assert(entries[0] == new ExcelBatchEntry(1, "  BOX00123  "), "lot ID should preserve surrounding spaces from Excel");
    Assert(entries[1] == new ExcelBatchEntry(42, "BOX00124"), "the second GNo should come directly from column B");

    var generator = new ExcelBatchGenerator();
    Assert(generator.ReadEntries(workbookPath, "批次一").SequenceEqual(entries),
        "batch reader should default to BC columns");
    foreach (var mode in new[] { ExcelReadMode.BC, ExcelReadMode.AB })
    {
        var modePath = Path.Combine(root, $"batch-{mode}.xlsx");
        CreateWorkbook(modePath, mode == ExcelReadMode.AB);
        Assert(generator.ReadEntries(modePath, "批次一", mode).SequenceEqual(entries),
            $"{mode} should read the selected serial and lot columns, skip headers and blanks, and preserve lot spaces");

        var serialColumn = mode == ExcelReadMode.AB ? "A" : "B";
        using (var archive = ZipFile.Open(modePath, ZipArchiveMode.Update))
        {
            var sheet = archive.GetEntry("xl/worksheets/sheet1.xml")!;
            XDocument document;
            using (var stream = sheet.Open())
            {
                document = XDocument.Load(stream);
            }
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            document.Descendants(ns + "c").Single(cell => (string?)cell.Attribute("r") == $"{serialColumn}7")
                .Element(ns + "v")!.Value = "invalid";
            sheet.Delete();
            AddXml(archive, "xl/worksheets/sheet1.xml", document.ToString());
        }

        var rejected = false;
        try
        {
            generator.ReadEntries(modePath, "批次一", mode);
        }
        catch (InvalidDataException ex)
        {
            rejected = ex.Message.Contains($"第 7 行 {serialColumn} 栏序号", StringComparison.Ordinal);
        }
        Assert(rejected, $"{mode} should report the actual row and serial column for invalid serials");
    }
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
        SupervisingCustomsCode = "3503",
        InformationEntryOperType = "G"
    };

    ConfigurationStore.Save(path, options);
    var loaded = ConfigurationStore.Load(path);
    Assert(loaded.ExportEnterprise.Name == "出口企业", "exporter configuration should round-trip independently");
    Assert(loaded.ApplicantEnterprise.Name == "申请单位", "applicant configuration should round-trip independently");
    Assert(loaded.SupervisingCustomsCode == "3503", "supervising customs code should round-trip");
    Assert(loaded.InformationEntryOperType == "G", "information-entry operation type should round-trip");

    var legacyPath = Path.Combine(root, "legacy-trade-xml-config.json");
    File.WriteAllText(legacyPath, "{}");
    var legacy = ConfigurationStore.Load(legacyPath);
    Assert(legacy.InformationEntryOperType == "C", "legacy configuration without an operation type should default to C");
}

static void TestExportEnterpriseProfiles(string root)
{
    var path = Path.Combine(root, "export-profiles.json");
    var options = new TradeXmlOptions
    {
        ExportEnterprise = new EnterpriseOptions { Name = "原出口企业", CustomsCode = "1234567890", SocialCreditCode = "原信用代码" },
        SupervisingCustomsCode = "2301",
        ApplicantEnterprise = new EnterpriseOptions { Name = "申请单位" },
        Operator = new OperatorOptions { OperName = "操作员" }
    };
    ConfigurationStore.Save(path, options);
    options = ConfigurationStore.Load(path);
    var manager = new ExportEnterpriseProfileManager(options);
    var originalId = manager.Selected.Id;
    Assert(manager.Profiles.Count == 1 && manager.Selected.Name == "默认方案"
        && manager.Selected.Enterprise == options.ExportEnterprise && manager.Selected.SupervisingCustomsCode == "2301",
        "legacy exporter settings should migrate intact into a default profile");

    manager.UpdateCurrent(options.ExportEnterprise with { Name = "已编辑出口企业" }, "4403");
    manager.Add("  第二方案  ");
    var secondId = manager.Selected.Id;
    Assert(manager.Selected.Name == "第二方案" && manager.Selected.Enterprise.Name == ""
        && manager.Selected.SupervisingCustomsCode == "", "adding a profile should select a blank independent profile");
    manager.UpdateCurrent(new EnterpriseOptions { Name = "新出口企业", CustomsCode = "0987654321", SocialCreditCode = "新信用代码" }, "3503");
    manager.Rename("新方案名");
    Assert(manager.Selected.Id == secondId && manager.Selected.Enterprise.Name == "新出口企业",
        "renaming a profile should preserve its identity and enterprise data");
    manager.Select(originalId);
    Assert(manager.Selected.Enterprise.Name == "已编辑出口企业" && manager.Selected.SupervisingCustomsCode == "4403",
        "switching back should preserve edits and restore the profile customs code");

    foreach (var name in new[] { "   ", "新方案名" })
    {
        foreach (var rename in new[] { false, true })
        {
            var rejected = false;
            try
            {
                if (rename) manager.Rename(name); else manager.Add(name);
            }
            catch (ArgumentException) { rejected = true; }
            Assert(rejected && manager.Profiles.Count == 2 && manager.Selected.Id == originalId,
                "blank or duplicate profile names should be rejected without changing profiles or selection");
        }
    }

    manager.Select(secondId);
    manager.ApplyTo(options);
    ConfigurationStore.Save(path, options);
    var loaded = ConfigurationStore.Load(path);
    var restored = new ExportEnterpriseProfileManager(loaded);
    Assert(restored.Profiles.Count == 2 && restored.Selected.Id == secondId && restored.Selected.Name == "新方案名",
        "all profiles, renamed names and the current selection should survive a configuration reload");
    Assert(loaded.ExportEnterprise == manager.Selected.Enterprise && loaded.SupervisingCustomsCode == "3503",
        "XML generation options should use the selected exporter and customs code");
    Assert(loaded.ApplicantEnterprise.Name == "申请单位" && loaded.Operator.OperName == "操作员",
        "exporter profile changes should preserve the applicant and operator settings");
    restored.Select(originalId);
    Assert(restored.Selected.Enterprise.Name == "已编辑出口企业", "inactive profile edits should also be persisted");
}

static void TestExportEnterpriseProfileDeletion(string root)
{
    var options = new TradeXmlOptions();
    var manager = new ExportEnterpriseProfileManager(options);
    var firstId = manager.Selected.Id;
    manager.Add("第二方案");
    var secondId = manager.Selected.Id;
    manager.UpdateCurrent(new EnterpriseOptions { Name = "保留企业" }, "3503");
    manager.Add("第三方案");
    var thirdId = manager.Selected.Id;

    manager.Select(firstId);
    manager.DeleteCurrent();
    Assert(manager.Profiles.Count == 2 && manager.Selected.Id == secondId
        && manager.Selected.Enterprise.Name == "保留企业" && manager.Selected.SupervisingCustomsCode == "3503",
        "deleting the first profile should select the next profile and preserve its enterprise data");
    manager.Select(thirdId);
    manager.DeleteCurrent();
    Assert(manager.Profiles.Count == 1 && manager.Selected.Id == secondId,
        "deleting the last profile in the list should select the preceding profile");

    manager.ApplyTo(options);
    var path = Path.Combine(root, "deleted-export-profiles.json");
    ConfigurationStore.Save(path, options);
    var loaded = ConfigurationStore.Load(path);
    Assert(loaded.ExportEnterpriseProfiles.Count == 1 && loaded.SelectedExportEnterpriseProfileId == secondId
        && loaded.ExportEnterprise.Name == "保留企业" && loaded.SupervisingCustomsCode == "3503",
        "saving and reloading should preserve deletion and use the remaining exporter for XML generation");

    var rejected = false;
    try { manager.DeleteCurrent(); }
    catch (InvalidOperationException) { rejected = true; }
    Assert(rejected && manager.Profiles.Count == 1 && manager.Selected.Id == secondId,
        "deleting the only remaining profile should be rejected without changing its data or selection");
}

static void TestXmlGeneration(string root)
{
    var photoFolder = Path.Combine(root, "BOX00123");
    var outputFolder = Path.Combine(root, "output");
    Directory.CreateDirectory(photoFolder);
    Directory.CreateDirectory(outputFolder);
    for (var index = 1; index <= 4; index++)
    {
        File.WriteAllBytes(Path.Combine(photoFolder, $"{index}.jpg"), CreateImage(false));
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
        MaxImageBytes = 1024,
        UploadTypeCode = "F",
        IncludeP0 = true,
        P0FilePath = p0Path
    };
    var request = new XmlGenerationRequest(
        photoFolder,
        outputFolder,
        "20260813164003000645181692",
        "PB202608040001",
        "7",
        "BOX00123",
        DateTimeOffset.Parse("2026-08-04T12:34:56.789-07:00"));
    var results = new TradeXmlGenerator().GenerateToFiles(request, options, false);

    Assert(results.Count == 6, "one ELBP004 and five ELBP005 files should be generated");
    Assert(Path.GetFileName(results[0].OutputPath).StartsWith("ELBP004_", StringComparison.Ordinal),
        "the first XML should use the reference ELBP004 filename pattern");

    XNamespace ns = TradeXmlGenerator.ContractNamespace;
    var main = XDocument.Load(results[0].OutputPath);
    Assert(main.Root?.Name == ns + "ELBP004Request" && main.Root.Name.NamespaceName == "",
        "main root should use the unqualified ELBP004 contract from the reference message");
    Assert(main.Root?.Element(ns + "Head")?.Element(ns + "ProBatchNumber")?.Value == "PB202608040001",
        "ELBP004 should contain the production batch number");
    Assert(main.Root?.Element(ns + "OperInfo")?.Element(ns + "OperType")?.Value == "C",
        "ELBP004 should default to declaration operation type C");
    Assert(main.Root?.Element(ns + "Head")?.Element(ns + "InputEtpsScc")?.Value == "91350781MA2YGJK59J",
        "ELBP004 should contain the independently configured exporter");
    Assert(main.Root?.Element(ns + "Head")?.Element(ns + "AgentScc")?.Value == "91110108MA0012345X",
        "ELBP004 should contain the independently configured applicant");
    Assert(main.Root?.Element(ns + "Lists")?.Element(ns + "List")?.Element(ns + "LotId")?.Value == "BOX00123",
        "ELBP004 should contain the lot list");
    Assert(main.Root?.Element(ns + "Lists")?.Element(ns + "List")?.Element(ns + "GNo")?.Value == "7",
        "single-entry ELBP004 should use the entered GNo");

    var metadata = main.Root?.Element(ns + "Edocs")?.Elements(ns + "Edoc").ToList() ?? [];
    Assert(metadata.Count == 5, "ELBP004 should reference P0 and A1-A4 metadata");
    Assert(metadata.Select(edoc => edoc.Element(ns + "BizTypeCode")?.Value).SequenceEqual(["A1", "A2", "A3", "A4", "P0"]),
        "ELBP004 attachment order should be A1-A4 followed by the shared P0");
    var edocIds = metadata
        .Select(edoc => edoc.Element(ns + "EdocID")?.Value ?? string.Empty)
        .ToList();
    var expectedEdocPrefixes = new[] { "A1", "A2", "A3", "A4", "P0" }
        .Select(code => $"91350781MA2YGJK59J3503{code}20260804123456789")
        .ToList();
    Assert(edocIds.Zip(expectedEdocPrefixes).All(pair =>
            pair.First.StartsWith(pair.Second, StringComparison.Ordinal)
            && pair.First.Length == pair.Second.Length + 20
            && pair.First[^20..].All(char.IsDigit)),
        "every EdocID should contain enterprise code, customs code, attachment code, timestamp, and a 20-digit serial number");
    Assert(edocIds.Distinct(StringComparer.Ordinal).Count() == edocIds.Count,
        "every EdocID should be unique");
    Assert(metadata[^1].Element(ns + "LotId")?.Value == "" && metadata.Take(4).All(edoc => edoc.Element(ns + "LotId")?.Value == "BOX00123"),
        "P0 should be batch-level while photos should reference the lot ID");

    var attachmentDocuments = results.Skip(1).Select(result => XDocument.Load(result.OutputPath)).ToList();
    Assert(attachmentDocuments.All(document => document.Root?.Name == ns + "ELBP005Request"),
        "attachment roots should use the unqualified ELBP005 reference contract");
    Assert(attachmentDocuments.All(document => document.Root?.Elements(ns + "Edoc").Count() == 1),
        "each ELBP005 should contain exactly one Edoc");
    Assert(attachmentDocuments.Select(document => document.Root!.Element(ns + "Edoc")!.Element(ns + "EdocID")!.Value)
            .SequenceEqual(edocIds),
        "ELBP004 and ELBP005 should share identical EdocIDs");
    Assert(attachmentDocuments.All(document => document.Root!.Element(ns + "Edoc")!.Element(ns + "FileContent") is not null),
        "each ELBP005 should carry Base64 file content");
    Assert(attachmentDocuments.All(document => document.Root!.Element(ns + "Edoc")!.Element(ns + "SeqNo") is null),
        "ELBP005 should not contain the removed SeqNo field");
    Assert(main.Root?.Element(ns + "OperInfo")?.Element(ns + "Sign")?.Value == "",
        "signature should remain empty for the import client");
    Assert(attachmentDocuments.All(document => document.Root?.Element(ns + "OperInfo")?.Element(ns + "CopCode")?.Value == ""),
        "ELBP005 should retain an empty CopCode element like the reference message");
}

static void TestBatchXmlGeneration(string root)
{
    var batchRoot = Path.Combine(root, "batch-xml-photos");
    var outputFolder = Path.Combine(root, "batch-xml-output");
    Directory.CreateDirectory(batchRoot);
    Directory.CreateDirectory(outputFolder);

    var entries = new[]
    {
        new ExcelBatchEntry(501, "NPS-R288 260803N100US0019"),
        new ExcelBatchEntry(702, "  NPS-R288 260803N100US0020  ")
    };
    foreach (var entry in entries)
    {
        var folder = Path.Combine(batchRoot, entry.LotId.Trim());
        Directory.CreateDirectory(folder);
        for (var index = 1; index <= 4; index++)
        {
            File.WriteAllBytes(Path.Combine(folder, $"{entry.Serial}-{index}.jpg"), CreateImage(false));
        }
    }

    var p0Path = Path.Combine(root, "batch-p0.pdf");
    File.WriteAllBytes(p0Path, [(byte)'p', (byte)'0']);
    var options = new TradeXmlOptions
    {
        Operator = new OperatorOptions
        {
            ICCode = "2026332131",
            CopCode = "91440300MA5EF74B5W",
            OperName = "Tester"
        },
        ExportEnterprise = new EnterpriseOptions
        {
            Name = "深圳市省油灯网络科技有限公司",
            CustomsCode = "4403961H56",
            SocialCreditCode = "914403003427277794"
        },
        ApplicantEnterprise = new EnterpriseOptions
        {
            Name = "深圳市森灏物流有限公司",
            CustomsCode = "44039809NU",
            SocialCreditCode = "91440300MA5EF74B5W"
        },
        SupervisingCustomsCode = "2301",
        MaxImageBytes = 1024,
        UploadTypeCode = "F",
        IncludeP0 = true,
        P0FilePath = p0Path
    };

    var results = new ExcelBatchGenerator().RunBatch(
        entries,
        batchRoot,
        outputFolder,
        BatchFolderMode.SmallFolders,
        "20260813164003000645181692",
        "L20260813163543000645178869",
        DateTimeOffset.Parse("2026-08-14T01:37:30-07:00"),
        options,
        false);

    Assert(results.Count == 2 && results.All(result => result.Success), "both batch rows should generate successfully");
    Assert(results.All(result => result.Detail.StartsWith("4 个 ELBP005", StringComparison.Ordinal)),
        "each lot should report exactly four A1-A4 ELBP005 files");

    var files = Directory.GetFiles(outputFolder, "*.xml");
    Assert(files.Length == 10, "two lots should create one ELBP004, one P0 and eight A1-A4 ELBP005 files");
    Assert(files.Count(path => Path.GetFileName(path).StartsWith("ELBP004_", StringComparison.Ordinal)) == 1,
        "a batch should create exactly one ELBP004");
    Assert(files.Count(path => Path.GetFileName(path).StartsWith("0_P0_", StringComparison.Ordinal)) == 1,
        "a batch should create exactly one shared P0 ELBP005");
    foreach (var entry in entries)
    {
        Assert(files.Count(path => Path.GetFileName(path).StartsWith($"{entry.LotId.Trim()}_A", StringComparison.Ordinal)) == 4,
            $"lot {entry.LotId} should create A1-A4");
    }

    var mainPath = files.Single(path => Path.GetFileName(path).StartsWith("ELBP004_", StringComparison.Ordinal));
    var main = XDocument.Load(mainPath);
    XNamespace ns = TradeXmlGenerator.ContractNamespace;
    var lists = main.Root?.Element(ns + "Lists")?.Elements(ns + "List").ToList() ?? [];
    var metadata = main.Root?.Element(ns + "Edocs")?.Elements(ns + "Edoc").ToList() ?? [];
    Assert(lists.Count == 2, "the shared ELBP004 should contain both lot list rows");
    Assert(lists.Select(list => list.Element(ns + "GNo")?.Value).SequenceEqual(["501", "702"]),
        "batch GNo values should come directly from Excel column B");
    Assert(lists.Select(list => list.Element(ns + "LotId")?.Value).SequenceEqual(entries.Select(entry => entry.LotId)),
        "batch ELBP004 lot IDs should preserve surrounding spaces from Excel");
    Assert(metadata.Where(edoc => edoc.Element(ns + "BizTypeCode")?.Value != "P0")
            .Select(edoc => edoc.Element(ns + "LotId")?.Value)
            .Distinct()
            .SequenceEqual(entries.Select(entry => entry.LotId)),
        "batch attachment metadata should preserve surrounding lot-id spaces");
    Assert(metadata.Count == 9, "the shared ELBP004 should contain eight photo entries and one P0 entry");
    Assert(metadata.Count(edoc => edoc.Element(ns + "BizTypeCode")?.Value == "P0") == 1,
        "the shared ELBP004 should reference P0 once");
    Assert(metadata[^1].Element(ns + "BizTypeCode")?.Value == "P0" && metadata[^1].Element(ns + "LotId")?.Value == "",
        "the shared P0 metadata should be last and have an empty lot ID");

    var metadataIds = metadata.Select(edoc => edoc.Element(ns + "EdocID")!.Value).Order().ToList();
    var attachmentIds = files
        .Where(path => path != mainPath)
        .Select(path => XDocument.Load(path).Root!.Element(ns + "Edoc")!.Element(ns + "EdocID")!.Value)
        .Order()
        .ToList();
    Assert(metadataIds.SequenceEqual(attachmentIds), "ELBP004 metadata IDs should match all batch ELBP005 IDs");

    foreach (var photoCount in new[] { 3, 5 })
    {
        var expectedOperType = photoCount == 3 ? "G" : "C";
        options.InformationEntryOperType = expectedOperType;
        var variableEntry = new ExcelBatchEntry(800 + photoCount, $"PHOTO-COUNT-{photoCount}");
        var variableRoot = Path.Combine(root, $"batch-{photoCount}-photos");
        var variableFolder = Path.Combine(variableRoot, variableEntry.LotId);
        var variableOutput = Path.Combine(root, $"batch-{photoCount}-output");
        Directory.CreateDirectory(variableFolder);
        Directory.CreateDirectory(variableOutput);
        for (var index = 1; index <= photoCount; index++)
        {
            File.WriteAllBytes(Path.Combine(variableFolder, $"{index}.jpg"), CreateImage(false));
        }

        var batchGenerator = new ExcelBatchGenerator();
        var preview = batchGenerator.Preview(
            [variableEntry],
            variableRoot,
            BatchFolderMode.SmallFolders,
            photoCount).Single();
        Assert(preview.Success && preview.PhotoCount == photoCount,
            $"the {photoCount}-photo batch option should be ready when the folder has exactly {photoCount} photos");

        var variableResults = batchGenerator.RunBatch(
            [variableEntry],
            variableRoot,
            variableOutput,
            BatchFolderMode.SmallFolders,
            "20260813164003000645181692",
            "L20260813163543000645178869",
            DateTimeOffset.Parse("2026-08-14T01:37:30-07:00"),
            options,
            false,
            photoCount);
        Assert(variableResults.Single().Success, $"the {photoCount}-photo batch should generate successfully");

        var expectedCodes = photoCount == 3
            ? new[] { "A1", "A2", "A3" }
            : ["A1", "A2", "A3", "A4", "B1"];
        var variableFiles = Directory.GetFiles(variableOutput, "*.xml");
        Assert(variableFiles.Length == photoCount + 2,
            $"the {photoCount}-photo batch should create ELBP004, P0 and {photoCount} photo ELBP005 files");
        Assert(expectedCodes.All(code => variableFiles.Any(path =>
                Path.GetFileName(path).StartsWith($"{variableEntry.LotId}_{code}_", StringComparison.Ordinal))),
            $"the {photoCount}-photo batch filenames should use the requested business type codes");

        var variableMainPath = variableFiles.Single(path =>
            Path.GetFileName(path).StartsWith("ELBP004_", StringComparison.Ordinal));
        var variableMain = XDocument.Load(variableMainPath);
        Assert(variableMain.Root?.Element(ns + "OperInfo")?.Element(ns + "OperType")?.Value == expectedOperType,
            $"the {photoCount}-photo ELBP004 should use operation type {expectedOperType}");
        var variableMetadata = variableMain.Root?
            .Element(ns + "Edocs")?
            .Elements(ns + "Edoc")
            .Select(edoc => edoc.Element(ns + "BizTypeCode")?.Value)
            .ToList() ?? [];
        Assert(variableMetadata.SequenceEqual(expectedCodes.Append("P0")),
            $"the {photoCount}-photo ELBP004 metadata should use the requested business type codes followed by P0");
    }
}

static byte[] CreateImage(bool large)
{
    var size = large ? 1024 : 8;
    var pixels = new byte[size * size * 3];
    for (var y = 0; y < size; y++)
        for (var x = 0; x < size; x++)
        {
            var offset = (y * size + x) * 3;
            pixels[offset] = (byte)(x / 4);
            pixels[offset + 1] = (byte)(y / 4);
            pixels[offset + 2] = 128;
        }
    var bitmap = BitmapSource.Create(size, size, 96, 96, PixelFormats.Bgr24, null, pixels, size * 3);
    BitmapEncoder encoder = large ? new BmpBitmapEncoder() : new PngBitmapEncoder();
    encoder.Frames.Add(BitmapFrame.Create(bitmap));
    using var output = new MemoryStream();
    encoder.Save(output);
    return output.ToArray();
}

static void TestAttachmentConversion(string root, string? samplePath)
{
    var input = Path.Combine(root, "conversion-input");
    var output = Path.Combine(root, "conversion-output");
    Directory.CreateDirectory(input);
    Directory.CreateDirectory(output);
    var original = CreateImage(true);
    if (samplePath is not null)
        original = Convert.FromBase64String(XDocument.Load(samplePath).Root!.Element("Edoc")!.Element("FileContent")!.Value);
    var small = CreateImage(false);
    for (var i = 1; i <= 4; i++)
        File.WriteAllBytes(Path.Combine(input, $"{i}.jpg"), i == 1 ? original : small);
    var options = new TradeXmlOptions
    {
        Operator = new() { ICCode = "1234567890", OperName = "Test" },
        ExportEnterprise = new() { Name = "Test", CustomsCode = "1234567890", SocialCreditCode = "123456789012345678" },
        ApplicantEnterprise = new() { Name = "Test", CustomsCode = "1234567890", SocialCreditCode = "123456789012345678" },
        SupervisingCustomsCode = "1234"
    };
    var request = new XmlGenerationRequest(input, output, "TEST", "BATCH", "1", "LOT", DateTimeOffset.Now);
    var generator = new TradeXmlGenerator();
    var results = generator.GenerateToFiles(request, options, false);
    var metadata = results[0].Document.Root!.Element("Edocs")!.Elements("Edoc").ToList();
    for (var i = 1; i <= 4; i++)
    {
        var edoc = XDocument.Load(results[i].OutputPath).Root!.Element("Edoc")!;
        var bytes = Convert.FromBase64String(edoc.Element("FileContent")!.Value);
        using var stream = new MemoryStream(bytes);
        var decoder = BitmapDecoder.Create(stream, BitmapCreateOptions.None, BitmapCacheOption.OnLoad);
        Assert(i == 1 ? decoder is JpegBitmapDecoder : decoder is PngBitmapDecoder,
            "oversized photo should become JPEG; small PNG with jpg suffix should retain PNG bytes");
        foreach (var field in new[] { "EdocID", "AttTypeCode", "AttEdocName" })
            Assert(edoc.Element(field)!.Value == metadata[i - 1].Element(field)!.Value,
                $"prepared attachment {field} should match in both messages");
        Assert(edoc.Element("AttTypeCode")!.Value == (i == 1 ? "jpg" : "png")
            && edoc.Element("AttEdocName")!.Value == $"{i}." + (i == 1 ? "jpg" : "png"),
            "declared type and filename should match the encoded image");
        Assert(bytes.Length <= options.MaxImageBytes && new FileInfo(results[i].OutputPath).Length < 3 * 1024 * 1024 - 64 * 1024,
            "attachment and actual XML sizes should fit the configured and conservative budgets");
        if (i == 1)
        {
            using var beforeStream = new MemoryStream(original);
            var before = BitmapDecoder.Create(beforeStream, BitmapCreateOptions.None, BitmapCacheOption.OnLoad).Frames[0];
            Assert(decoder.Frames[0].PixelWidth == before.PixelWidth && decoder.Frames[0].PixelHeight == before.PixelHeight,
                "compression should preserve pixel dimensions");
            Console.WriteLine($"Conversion: {original.Length} -> {bytes.Length} bytes; XML {new FileInfo(results[i].OutputPath).Length} bytes.");
        }
        else Assert(bytes.SequenceEqual(small), "small images should not be re-encoded");
    }
    Assert(File.ReadAllBytes(Path.Combine(input, "1.jpg")).SequenceEqual(original), "original photo should remain unchanged");
    var batchOutput = Path.Combine(root, "conversion-batch");
    Directory.CreateDirectory(batchOutput);
    var batch = new ExcelBatchGenerator().RunBatch([new(1, "conversion-input")], root, batchOutput,
        BatchFolderMode.SmallFolders, "TEST", "BATCH", DateTimeOffset.Now, options, false);
    Assert(batch.Single().Success, "batch path should allow oversized originals to reach compression");

    var blockedOutput = Path.Combine(root, "conversion-blocked");
    Directory.CreateDirectory(blockedOutput);
    options.MaxImageBytes = 1;
    var rejected = false;
    try { generator.GenerateToFiles(request with { OutputFolderPath = blockedOutput }, options, false); }
    catch (XmlGenerationException ex) { rejected = ex.Message.Contains("质量降至 80"); }
    Assert(rejected && !Directory.EnumerateFiles(blockedOutput, "*.xml").Any(),
        "impossible compression target should fail before writing any XML");
    options.MaxImageBytes = TradeXmlOptions.DefaultMaxImageBytes;
    File.WriteAllBytes(Path.Combine(input, "1.jpg"), [1, 2, 3]);
    rejected = false;
    try { generator.GenerateToFiles(request with { OutputFolderPath = blockedOutput }, options, false); }
    catch (XmlGenerationException ex) { rejected = ex.Message.Contains("图片读取或转换失败"); }
    Assert(rejected && !Directory.EnumerateFiles(blockedOutput, "*.xml").Any(), "corrupt photo should fail before writing any XML");
}

static void CreateWorkbook(string path, bool useAbColumns = false)
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
          <si><t>序号</t></si><si><t>箱号</t></si><si><t xml:space="preserve">  BOX00123  </t></si><si><t>BOX00124</t></si>
        </sst>
        """);
    AddXml(archive, "xl/worksheets/sheet1.xml", """
        <?xml version="1.0" encoding="UTF-8"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
          <row r="1"><c r="B1" t="s"><v>0</v></c><c r="C1" t="s"><v>1</v></c></row>
          <row r="2"><c r="B2"><v>1</v></c><c r="C2" t="s"><v>2</v></c><c r="D2" t="inlineStr"><is><t>忽略其他栏</t></is></c></row>
          <row r="3"><c r="B3"><v>2</v></c><c r="C3" t="inlineStr"><is><t></t></is></c></row>
          <row r="7"><c r="B7"><v>42</v></c><c r="C7" t="s"><v>3</v></c></row>
        </sheetData></worksheet>
        """.Replace("r=\"B", useAbColumns ? "r=\"A" : "r=\"B")
            .Replace("r=\"C", useAbColumns ? "r=\"B" : "r=\"C"));
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
