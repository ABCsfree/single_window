using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using TradeXmlStudio.Core;

namespace TradeXmlStudio.App;

public partial class MainWindow : Window
{
    private sealed record PhotoPreviewRow(
        string FileName,
        string Status,
        string SizeText,
        string BizTypeCode,
        string FullPath);

    private readonly TradeXmlGenerator _xmlGenerator = new();
    private readonly ExcelBatchGenerator _batchGenerator = new();
    private readonly string _configPath = Path.Combine(AppContext.BaseDirectory, "trade-xml-config.json");
    private List<ExcelBatchEntry> _batchEntries = [];
    private string? _lastOutputFolder;
    private bool _isReady;

    public MainWindow()
    {
        InitializeComponent();
        LoadConfiguration();
        _isReady = true;
        RefreshPhotoPreview();
    }

    private void BrowseSourceFolder_Click(object sender, RoutedEventArgs e)
    {
        if (TryPickFolder("选择包含 4 张照片的文件夹", SourceFolderTextBox.Text, out var folder))
        {
            SourceFolderTextBox.Text = folder;
            if (string.IsNullOrWhiteSpace(OutputFolderTextBox.Text))
            {
                OutputFolderTextBox.Text = folder;
            }
        }
    }

    private void BrowseOutputFolder_Click(object sender, RoutedEventArgs e)
    {
        if (TryPickFolder("选择 XML 输出目录", OutputFolderTextBox.Text, out var folder))
        {
            OutputFolderTextBox.Text = folder;
        }
    }

    private void BrowseP0File_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "选择代理委托协议（P0）文件",
            Filter = "图片及 PDF (*.jpg;*.jpeg;*.png;*.bmp;*.pdf)|*.jpg;*.jpeg;*.png;*.bmp;*.pdf|所有文件 (*.*)|*.*",
            CheckFileExists = true
        };
        ApplyExistingFilePath(dialog, P0FilePathTextBox.Text);
        if (dialog.ShowDialog(this) == true)
        {
            P0FilePathTextBox.Text = dialog.FileName;
        }
    }

    private void SaveConfig_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            ConfigurationStore.Save(_configPath, BuildOptions());
            SetStatus($"配置已保存：{_configPath}");
            MessageBox.Show(this, "配置已保存。", "保存成功", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            ShowError($"配置保存失败：{ex.Message}");
        }
    }

    private void GenerateXml_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var request = new XmlGenerationRequest(
                SourceFolderTextBox.Text.Trim(),
                OutputFolderTextBox.Text.Trim(),
                SeqNoTextBox.Text,
                ProBatchNumberTextBox.Text,
                GNoTextBox.Text,
                LotIdTextBox.Text,
                DateTimeOffset.Now);
            var overwrite = ConfirmSingleOverwrite(request);
            var results = _xmlGenerator.GenerateToFiles(request, BuildOptions(), overwrite);
            _lastOutputFolder = results.Count > 0
                ? Path.GetDirectoryName(results[0].OutputPath)
                : request.OutputFolderPath;
            OpenOutputFolderButton.IsEnabled = Directory.Exists(_lastOutputFolder);
            RefreshPhotoPreview();

            var names = string.Join(Environment.NewLine, results.Select(result => Path.GetFileName(result.OutputPath)));
            SetStatus($"生成成功：{results.Count} 个 XML。");
            MessageBox.Show(this, $"已生成 {results.Count} 个 XML：{Environment.NewLine}{names}",
                "生成成功", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (OperationCanceledException)
        {
            SetStatus("已取消生成。");
        }
        catch (XmlGenerationException ex)
        {
            ShowError(ex.Message);
            SetStatus("生成失败：输入校验未通过。");
        }
        catch (Exception ex)
        {
            ShowError($"生成 XML 失败：{ex.Message}");
            SetStatus($"生成 XML 失败：{ex.Message}");
        }
    }

    private void OpenOutputFolder_Click(object sender, RoutedEventArgs e) => OpenFolder(_lastOutputFolder);

    private void SingleInputChanged(object sender, TextChangedEventArgs e)
    {
        if (_isReady)
        {
            RefreshPhotoPreview();
        }
    }

    private void P0IncludeChanged(object sender, RoutedEventArgs e)
    {
        if (_isReady)
        {
            RefreshPhotoPreview();
        }
    }

    private void BrowseBatchExcel_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "选择 Excel 文件",
            Filter = "Excel 工作簿 (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|所有文件 (*.*)|*.*",
            CheckFileExists = true
        };
        ApplyExistingFilePath(dialog, BatchExcelPathTextBox.Text);
        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        BatchExcelPathTextBox.Text = dialog.FileName;
        LoadBatchSheets(dialog.FileName);
        // Selecting a workbook is only one step of the batch setup. Do not show
        // a validation error while the user still needs to choose the photo root.
        RefreshBatchPreview();
    }

    private void LoadBatchSheets(string excelPath)
    {
        BatchSheetComboBox.Items.Clear();
        try
        {
            foreach (var sheet in _batchGenerator.ListSheets(excelPath))
            {
                BatchSheetComboBox.Items.Add(sheet);
            }

            if (BatchSheetComboBox.Items.Count > 0)
            {
                BatchSheetComboBox.SelectedIndex = 0;
            }
            else
            {
                SetStatus("Excel 中没有可读取的工作表。");
            }
        }
        catch (Exception ex)
        {
            ShowError($"读取 Excel 工作表失败：{ex.Message}");
            SetStatus($"读取 Excel 工作表失败：{ex.Message}");
        }
    }

    private void BatchSheet_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (_isReady)
        {
            RefreshBatchPreview();
        }
    }

    private void BatchInputChanged(object sender, TextChangedEventArgs e)
    {
        if (_isReady)
        {
            RefreshBatchPreview();
        }
    }

    private void BrowseBatchBigFolder_Click(object sender, RoutedEventArgs e)
    {
        if (TryPickFolder("选择影像根目录", BatchBigFolderTextBox.Text, out var folder))
        {
            BatchBigFolderTextBox.Text = folder;
        }
    }

    private void BrowseBatchOutputFolder_Click(object sender, RoutedEventArgs e)
    {
        if (TryPickFolder("选择 XML 输出目录", BatchOutputFolderTextBox.Text, out var folder))
        {
            BatchOutputFolderTextBox.Text = folder;
        }
    }

    private void BatchFolderMode_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (_isReady)
        {
            RefreshBatchPreview();
        }
    }

    private async void GenerateBatch_Click(object sender, RoutedEventArgs e)
    {
        // Always reread Excel here. This allows users to save workbook changes
        // after previewing without restarting the application.
        if (!RefreshBatchPreview(true))
        {
            return;
        }

        var bigFolder = BatchBigFolderTextBox.Text.Trim();
        var outputFolder = BatchOutputFolderTextBox.Text.Trim();
        var seqNo = BatchSeqNoTextBox.Text.Trim();
        var proBatchNumber = BatchProBatchNumberTextBox.Text.Trim();
        if (_batchEntries.Count == 0)
        {
            ShowError("没有可生成的行。请保存 Excel，并确认所选工作表的 C 列包含箱号。");
            return;
        }
        if (string.IsNullOrWhiteSpace(seqNo))
        {
            ShowError("通知编号不能为空。");
            return;
        }
        if (string.IsNullOrWhiteSpace(proBatchNumber))
        {
            ShowError("生产批次号不能为空。");
            return;
        }
        if (!Directory.Exists(bigFolder))
        {
            ShowError("影像根目录不存在。");
            return;
        }
        if (string.IsNullOrWhiteSpace(outputFolder))
        {
            ShowError("请单独选择输出目录。");
            return;
        }

        try
        {
            Directory.CreateDirectory(outputFolder);
        }
        catch (Exception ex)
        {
            ShowError($"输出目录无法创建：{ex.Message}");
            return;
        }

        if (MessageBox.Show(this,
                $"将处理 {_batchEntries.Count} 行；整批生成 1 份 ELBP004、最多 1 份 P0，每个有效箱号生成 A1-A4 共 4 份 ELBP005。是否继续？",
                "确认批量生成",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question) != MessageBoxResult.Yes)
        {
            return;
        }

        GenerateBatchButton.IsEnabled = false;
        SetStatus("批量生成中…");
        try
        {
            var entries = _batchEntries.ToList();
            var mode = GetBatchFolderMode();
            var options = BuildOptions();
            var generatedAt = DateTimeOffset.Now;
            var results = await Task.Run(() => _batchGenerator.RunBatch(
                entries, bigFolder, outputFolder, mode, seqNo, proBatchNumber, generatedAt, options, true));
            BatchResultListView.ItemsSource = results;

            var successCount = results.Count(result => result.Success);
            var failureCount = results.Count - successCount;
            OpenBatchFolderButton.IsEnabled = Directory.Exists(outputFolder);
            SetStatus($"批量完成：成功 {successCount}，失败 {failureCount}。输出目录：{outputFolder}");
            MessageBox.Show(this,
                failureCount == 0
                    ? $"全部成功：生成 {successCount} 个箱号；ELBP004 与 P0 均为整批共享。"
                    : $"成功 {successCount} 个箱号，失败 {failureCount} 个箱号；详情请查看列表。",
                failureCount == 0 ? "批量完成" : "批量完成（含失败）",
                MessageBoxButton.OK,
                failureCount == 0 ? MessageBoxImage.Information : MessageBoxImage.Warning);
        }
        catch (Exception ex)
        {
            ShowError($"批量生成失败：{ex.Message}");
            SetStatus($"批量生成失败：{ex.Message}");
        }
        finally
        {
            GenerateBatchButton.IsEnabled = true;
        }
    }

    private void OpenBatchFolder_Click(object sender, RoutedEventArgs e)
    {
        OpenFolder(BatchOutputFolderTextBox.Text.Trim());
    }

    private bool RefreshBatchPreview(bool showErrors = false)
    {
        var excelPath = BatchExcelPathTextBox.Text.Trim();
        var bigFolder = BatchBigFolderTextBox.Text.Trim();
        var sheetName = BatchSheetComboBox.SelectedItem as string;
        var validationError = !File.Exists(excelPath)
            ? "请选择存在的 Excel 文件。"
            : string.IsNullOrWhiteSpace(sheetName)
                ? "请选择 Excel 工作表。"
                : !Directory.Exists(bigFolder)
                    ? "请选择存在的影像根目录。"
                    : null;
        if (validationError is not null)
        {
            _batchEntries = [];
            BatchResultListView.ItemsSource = Array.Empty<ExcelBatchItemResult>();
            if (showErrors)
            {
                ShowError(validationError);
            }
            return false;
        }

        try
        {
            _batchEntries = _batchGenerator.ReadEntries(excelPath, sheetName!).ToList();
            BatchResultListView.ItemsSource = _batchGenerator.Preview(_batchEntries, bigFolder, GetBatchFolderMode());
            SetStatus($"已读取 {_batchEntries.Count} 行，预览就绪。");
            return true;
        }
        catch (Exception ex)
        {
            _batchEntries = [];
            BatchResultListView.ItemsSource = Array.Empty<ExcelBatchItemResult>();
            if (showErrors)
            {
                ShowError($"读取 Excel 失败：{ex.Message}");
            }
            SetStatus($"读取 Excel 失败：{ex.Message}");
            return false;
        }
    }

    private void RefreshPhotoPreview()
    {
        var options = BuildOptions();
        PhotoPreviewListView.ItemsSource = TradeXmlGenerator
            .ScanEdocs(SourceFolderTextBox.Text.Trim(), options.P0FilePath, options.IncludeP0)
            .Select(source =>
            {
                var exists = File.Exists(source.FullPath);
                return new PhotoPreviewRow(
                    source.FileName,
                    exists ? "存在" : "缺少",
                    exists ? FormatFileSize(new FileInfo(source.FullPath).Length) : "-",
                    source.BizTypeCode,
                    source.FullPath);
            })
            .ToList();
    }

    private TradeXmlOptions BuildOptions() => new()
    {
        Operator = new OperatorOptions
        {
            ICCode = ICCodeTextBox.Text.Trim(),
            CopCode = CopCodeTextBox.Text.Trim(),
            OperName = OperNameTextBox.Text.Trim()
        },
        ExportEnterprise = new EnterpriseOptions
        {
            Name = ExportNameTextBox.Text.Trim(),
            CustomsCode = ExportCustomsCodeTextBox.Text.Trim(),
            SocialCreditCode = ExportSccTextBox.Text.Trim()
        },
        ApplicantEnterprise = new EnterpriseOptions
        {
            Name = AgentNameTextBox.Text.Trim(),
            CustomsCode = AgentCustomsCodeTextBox.Text.Trim(),
            SocialCreditCode = AgentSccTextBox.Text.Trim()
        },
        SupervisingCustomsCode = SupervisingCustomsCodeTextBox.Text.Trim(),
        UploadTypeCode = GetComboBoxValue(UploadTypeCodeComboBox, "F"),
        MaxImageBytes = ParseMaxImageBytes(),
        IncludeP0 = P0IncludeCheckBox.IsChecked == true,
        P0FilePath = P0FilePathTextBox.Text.Trim()
    };

    private void LoadConfiguration()
    {
        try
        {
            ApplyOptions(ConfigurationStore.LoadOrCreateDefault(_configPath));
            SetStatus($"已加载配置：{_configPath}");
        }
        catch (Exception ex)
        {
            ApplyOptions(new TradeXmlOptions());
            ShowError(ex.Message);
            SetStatus($"配置加载失败：{ex.Message}");
        }
    }

    private void ApplyOptions(TradeXmlOptions options)
    {
        ICCodeTextBox.Text = options.Operator.ICCode;
        CopCodeTextBox.Text = options.Operator.CopCode;
        OperNameTextBox.Text = options.Operator.OperName;
        ExportNameTextBox.Text = options.ExportEnterprise.Name;
        ExportCustomsCodeTextBox.Text = options.ExportEnterprise.CustomsCode;
        ExportSccTextBox.Text = options.ExportEnterprise.SocialCreditCode;
        SupervisingCustomsCodeTextBox.Text = options.SupervisingCustomsCode;
        AgentNameTextBox.Text = options.ApplicantEnterprise.Name;
        AgentCustomsCodeTextBox.Text = options.ApplicantEnterprise.CustomsCode;
        AgentSccTextBox.Text = options.ApplicantEnterprise.SocialCreditCode;
        MaxImageMbTextBox.Text = (options.MaxImageBytes / 1024d / 1024d).ToString("0.##", CultureInfo.CurrentCulture);
        P0IncludeCheckBox.IsChecked = options.IncludeP0;
        P0FilePathTextBox.Text = options.P0FilePath;
        SelectComboBoxValue(UploadTypeCodeComboBox, options.UploadTypeCode, "F");
    }

    private long ParseMaxImageBytes()
    {
        var text = MaxImageMbTextBox.Text.Trim();
        if ((!double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out var value)
             && !double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value))
            || value <= 0)
        {
            return 0;
        }
        return (long)Math.Round(value * 1024 * 1024, MidpointRounding.AwayFromZero);
    }

    private bool ConfirmSingleOverwrite(XmlGenerationRequest request)
    {
        if (!Directory.Exists(request.SourceFolderPath) || !Directory.Exists(request.OutputFolderPath))
        {
            return false;
        }

        var prefix = request.LotId.Trim();
        var existing = Directory.EnumerateFiles(request.OutputFolderPath, $"{prefix}_*ELBP*.xml")
            .Select(Path.GetFileName)
            .ToList();
        if (existing.Count == 0)
        {
            return false;
        }

        if (MessageBox.Show(this,
                $"输出目录已有本箱号生成的 {existing.Count} 个 XML，是否覆盖？{Environment.NewLine}"
                + string.Join(Environment.NewLine, existing),
                "确认覆盖",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question) != MessageBoxResult.Yes)
        {
            throw new OperationCanceledException();
        }
        return true;
    }

    private BatchFolderMode GetBatchFolderMode() =>
        BatchFolderModeComboBox.SelectedIndex switch
        {
            1 => BatchFolderMode.SerialSmallFolders,
            2 => BatchFolderMode.SingleBigFolder,
            _ => BatchFolderMode.SmallFolders
        };

    private bool TryPickFolder(string title, string currentPath, out string folder)
    {
        var dialog = new OpenFolderDialog { Title = title };
        if (Directory.Exists(currentPath))
        {
            dialog.InitialDirectory = currentPath;
        }
        if (dialog.ShowDialog(this) == true)
        {
            folder = dialog.FolderName;
            return true;
        }
        folder = "";
        return false;
    }

    private void OpenFolder(string? folder)
    {
        if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
        {
            ShowError("输出目录不存在。");
            return;
        }
        Process.Start(new ProcessStartInfo(folder!) { UseShellExecute = true });
    }

    private static void ApplyExistingFilePath(FileDialog dialog, string path)
    {
        if (!File.Exists(path))
        {
            return;
        }
        dialog.InitialDirectory = Path.GetDirectoryName(path);
        dialog.FileName = Path.GetFileName(path);
    }

    private static string GetComboBoxValue(ComboBox comboBox, string fallback) =>
        (comboBox.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? fallback;

    private static void SelectComboBoxValue(ComboBox comboBox, string value, string fallback)
    {
        var selected = comboBox.Items.OfType<ComboBoxItem>()
            .FirstOrDefault(item => string.Equals(item.Content?.ToString(), value, StringComparison.Ordinal))
            ?? comboBox.Items.OfType<ComboBoxItem>()
                .First(item => string.Equals(item.Content?.ToString(), fallback, StringComparison.Ordinal));
        comboBox.SelectedItem = selected;
    }

    private static string FormatFileSize(long bytes) => $"{bytes / 1024d / 1024d:0.00} MB";

    private void SetStatus(string message) => StatusTextBlock.Text = message;

    private void ShowError(string message) =>
        MessageBox.Show(this, message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
}
