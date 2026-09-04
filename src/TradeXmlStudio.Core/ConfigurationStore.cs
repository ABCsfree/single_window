using System.Text;
using System.Text.Json;

namespace TradeXmlStudio.Core;

public static class ConfigurationStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true
    };

    public static TradeXmlOptions LoadOrCreateDefault(string path)
    {
        if (!File.Exists(path))
        {
            var defaults = new TradeXmlOptions();
            Save(path, defaults);
            return defaults;
        }

        return Load(path);
    }

    public static TradeXmlOptions Load(string path)
    {
        try
        {
            var options = JsonSerializer.Deserialize<TradeXmlOptions>(
                File.ReadAllText(path, Encoding.UTF8), JsonOptions)
                ?? throw new InvalidOperationException("配置内容为空。");
            options.Operator ??= new OperatorOptions();
            options.ExportEnterprise ??= new EnterpriseOptions();
            options.ApplicantEnterprise ??= new EnterpriseOptions();
            return options;
        }
        catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException or InvalidOperationException)
        {
            throw new InvalidOperationException($"配置文件读取失败：{ex.Message}", ex);
        }
    }

    public static void Save(string path, TradeXmlOptions options)
    {
        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        File.WriteAllText(path, JsonSerializer.Serialize(options, JsonOptions), new UTF8Encoding(false));
    }
}
