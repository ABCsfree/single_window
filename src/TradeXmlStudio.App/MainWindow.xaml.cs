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
    private bool _updatingExportProfiles;
    private ExportEnterpriseProfileManager _exportProfiles = new(new TradeXmlOptions());

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
            var options = BuildOptions();
            var results = _xmlGenerator.GenerateToFiles(request, options, overwrite);
            _lastOutputFolder = results.Count > 0
                ? Path.GetDirectoryName(results[0].OutputPath)
                : request.OutputFolderPath;
            OpenOutputFolderButton.IsEnabled = Directory.Exists(_lastOutputFolder);
            RefreshPhotoPreview();

            var names = string.Join(Environment.NewLine, results.Select(result => Path.GetFileName(result.OutputPath)));
            var operationText = FormatInformationEntryOperation(options.InformationEntryOperType);
            SetStatus($"生成成功：{results.Count} 个 XML；信息补录方式：{operationText}。");
            MessageBox.Show(this,
                $"信息补录方式：{operationText}{Environment.NewLine}已生成 {results.Count} 个 XML：{Environment.NewLine}{names}",
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

    private ExcelReadMode GetBatchReadMode() =>
        BatchReadModeComboBox.SelectedIndex == 1 ? ExcelReadMode.AB : ExcelReadMode.BC;

    private void BatchReadMode_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (_isReady)
        {
            BatchPreviewTitleTextBlock.Text = GetBatchReadMode() == ExcelReadMode.AB
                ? "批量预览（A 栏序号，B 栏箱号）"
                : "批量预览（B 栏序号，C 栏箱号）";
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

    private void BatchPhotoCount_Changed(object sender, SelectionChangedEventArgs e)
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
        var expectedPhotoCount = GetBatchPhotoCount();
        var photoBizTypeText = string.Join("、", TradeXmlGenerator.GetBatchPhotoBizTypeCodes(expectedPhotoCount));
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

        var options = BuildOptions();
        var operationText = FormatInformationEntryOperation(options.InformationEntryOperType);
        if (MessageBox.Show(this,
                $"信息补录方式：{operationText}。将处理 {_batchEntries.Count} 行；整批生成 1 份 ELBP004、最多 1 份 P0，"
                + $"每个有效箱号生成 {photoBizTypeText} 共 {expectedPhotoCount} 份 ELBP005。是否继续？",
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
            var generatedAt = DateTimeOffset.Now;
            var results = await Task.Run(() => _batchGenerator.RunBatch(
                entries,
                bigFolder,
                outputFolder,
                mode,
                seqNo,
                proBatchNumber,
                generatedAt,
                options,
                true,
                expectedPhotoCount));
            BatchResultListView.ItemsSource = results;

            var successCount = results.Count(result => result.Success);
            var failureCount = results.Count - successCount;
            OpenBatchFolderButton.IsEnabled = Directory.Exists(outputFolder);
            SetStatus($"批量完成：成功 {successCount}，失败 {failureCount}；信息补录方式：{operationText}。输出目录：{outputFolder}");
            MessageBox.Show(this,
                failureCount == 0
                    ? $"全部成功：生成 {successCount} 个箱号；信息补录方式为 {operationText}；ELBP004 与 P0 均为整批共享。"
                    : $"成功 {successCount} 个箱号，失败 {failureCount} 个箱号；信息补录方式为 {operationText}；详情请查看列表。",
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
            _batchEntries = _batchGenerator.ReadEntries(excelPath, sheetName!, GetBatchReadMode()).ToList();
            var expectedPhotoCount = GetBatchPhotoCount();
            BatchResultListView.ItemsSource = _batchGenerator.Preview(
                _batchEntries,
                bigFolder,
                GetBatchFolderMode(),
                expectedPhotoCount);
            SetStatus($"已读取 {_batchEntries.Count} 行；本批每箱上传 {expectedPhotoCount} 张，预览就绪。");
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

    private TradeXmlOptions BuildOptions()
    {
        CaptureExportProfile();
        var options = new TradeXmlOptions
        {
            Operator = new OperatorOptions
            {
                ICCode = ICCodeTextBox.Text.Trim(),
                CopCode = CopCodeTextBox.Text.Trim(),
                OperName = OperNameTextBox.Text.Trim()
            },
            ApplicantEnterprise = new EnterpriseOptions
            {
                Name = AgentNameTextBox.Text.Trim(),
                CustomsCode = AgentCustomsCodeTextBox.Text.Trim(),
                SocialCreditCode = AgentSccTextBox.Text.Trim()
            },
            InformationEntryOperType = GetComboBoxValue(InformationEntryOperTypeComboBox, "C"),
            UploadTypeCode = GetComboBoxValue(UploadTypeCodeComboBox, "F"),
            MaxImageBytes = ParseMaxImageBytes(),
            IncludeP0 = P0IncludeCheckBox.IsChecked == true,
            P0FilePath = P0FilePathTextBox.Text.Trim()
        };
        _exportProfiles.ApplyTo(options);
        return options;
    }

    private void CaptureExportProfile() => _exportProfiles.UpdateCurrent(new EnterpriseOptions
    {
        Name = ExportNameTextBox.Text.Trim(),
        CustomsCode = ExportCustomsCodeTextBox.Text.Trim(),
        SocialCreditCode = ExportSccTextBox.Text.Trim()
    }, SupervisingCustomsCodeTextBox.Text.Trim());

    private void RefreshExportProfiles()
    {
        _updatingExportProfiles = true;
        try
        {
            ExportProfileComboBox.ItemsSource = _exportProfiles.Profiles.ToList();
            ExportProfileComboBox.SelectedValue = _exportProfiles.Selected.Id;
            var profile = _exportProfiles.Selected;
            ExportNameTextBox.Text = profile.Enterprise.Name;
            ExportCustomsCodeTextBox.Text = profile.Enterprise.CustomsCode;
            ExportSccTextBox.Text = profile.Enterprise.SocialCreditCode;
            SupervisingCustomsCodeTextBox.Text = profile.SupervisingCustomsCode;
        }
        finally
        {
            _updatingExportProfiles = false;
        }
    }

    private void ExportProfile_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (!_isReady || _updatingExportProfiles || ExportProfileComboBox.SelectedValue is not string id)
        {
            return;
        }
        CaptureExportProfile();
        _exportProfiles.Select(id);
        RefreshExportProfiles();
        SetStatus($"已切换出口企业方案：{_exportProfiles.Selected.Name}；点击“保存配置”保存。");
    }

    private void AddExportProfile_Click(object sender, RoutedEventArgs e) => EditExportProfileName(true);

    private void RenameExportProfile_Click(object sender, RoutedEventArgs e) => EditExportProfileName(false);

    private void DeleteExportProfile_Click(object sender, RoutedEventArgs e)
    {
        if (_exportProfiles.Profiles.Count == 1)
        {
            MessageBox.Show(this, "至少需要保留一个出口企业方案，无法删除最后一个方案。",
                "无法删除方案", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var name = _exportProfiles.Selected.Name;
        if (MessageBox.Show(this,
                $"确定删除出口企业方案“{name}”及其企业信息吗？点击“保存配置”后保存删除结果。",
                "删除方案", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) != MessageBoxResult.Yes)
        {
            return;
        }

        _exportProfiles.DeleteCurrent();
        RefreshExportProfiles();
        SetStatus($"已删除方案：{name}；已切换到：{_exportProfiles.Selected.Name}；点击“保存配置”保存。");
    }

    private void EditExportProfileName(bool add)
    {
        var dialog = new Window
        {
            Owner = this,
            Title = add ? "新增出口企业方案" : "修改方案名",
            Width = 440,
            SizeToContent = SizeToContent.Height,
            ResizeMode = ResizeMode.NoResize,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            ShowInTaskbar = false
        };
        var panel = new StackPanel { Margin = new Thickness(20) };
        panel.Children.Add(new TextBlock { Text = "方案名称", Margin = new Thickness(0, 0, 0, 8) });
        var nameBox = new TextBox { Text = add ? "" : _exportProfiles.Selected.Name, MaxLength = 100 };
        panel.Children.Add(nameBox);
        var error = new TextBlock
        {
            Foreground = System.Windows.Media.Brushes.Firebrick,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 8, 0, 8)
        };
        panel.Children.Add(error);
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        var cancel = new Button { Content = "取消", IsCancel = true };
        var confirm = new Button { Content = add ? "新增" : "确定", IsDefault = true };
        buttons.Children.Add(cancel);
        buttons.Children.Add(confirm);
        panel.Children.Add(buttons);
        dialog.Content = panel;
        confirm.Click += (_, _) =>
        {
            try
            {
                CaptureExportProfile();
                if (add)
                {
                    _exportProfiles.Add(nameBox.Text);
                }
                else
                {
                    _exportProfiles.Rename(nameBox.Text);
                }
                dialog.DialogResult = true;
            }
            catch (ArgumentException ex)
            {
                error.Text = ex.Message;
            }
        };
        dialog.Loaded += (_, _) => { nameBox.Focus(); nameBox.SelectAll(); };
        if (dialog.ShowDialog() == true)
        {
            RefreshExportProfiles();
            SetStatus($"出口企业方案已{(add ? "新增" : "重命名")}：{_exportProfiles.Selected.Name}；点击“保存配置”保存。");
        }
    }

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
        _exportProfiles = new ExportEnterpriseProfileManager(options);
        RefreshExportProfiles();
        AgentNameTextBox.Text = options.ApplicantEnterprise.Name;
        AgentCustomsCodeTextBox.Text = options.ApplicantEnterprise.CustomsCode;
        AgentSccTextBox.Text = options.ApplicantEnterprise.SocialCreditCode;
        MaxImageMbTextBox.Text = (options.MaxImageBytes / 1024d / 1024d).ToString("0.##", CultureInfo.CurrentCulture);
        P0IncludeCheckBox.IsChecked = options.IncludeP0;
        P0FilePathTextBox.Text = options.P0FilePath;
        SelectComboBoxValue(InformationEntryOperTypeComboBox, options.InformationEntryOperType, "C");
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

    private int GetBatchPhotoCount() =>
        BatchPhotoCountComboBox.SelectedIndex switch
        {
            0 => 3,
            2 => 5,
            _ => 4
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
        comboBox.SelectedItem is ComboBoxItem item ? GetComboBoxItemValue(item) : fallback;

    private static void SelectComboBoxValue(ComboBox comboBox, string value, string fallback)
    {
        var selected = comboBox.Items.OfType<ComboBoxItem>()
            .FirstOrDefault(item => string.Equals(GetComboBoxItemValue(item), value, StringComparison.Ordinal))
            ?? comboBox.Items.OfType<ComboBoxItem>()
                .First(item => string.Equals(GetComboBoxItemValue(item), fallback, StringComparison.Ordinal));
        comboBox.SelectedItem = selected;
    }

    private static string GetComboBoxItemValue(ComboBoxItem item) =>
        item.Tag?.ToString() ?? item.Content?.ToString() ?? "";

    private static string FormatInformationEntryOperation(string operType) =>
        string.Equals(operType?.Trim(), "G", StringComparison.Ordinal) ? "暂存（G）" : "申报（C）";

    private static string FormatFileSize(long bytes) => $"{bytes / 1024d / 1024d:0.00} MB";

    private void SetStatus(string message) => StatusTextBlock.Text = message;

    private void ShowError(string message) =>
        MessageBox.Show(this, message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
}
