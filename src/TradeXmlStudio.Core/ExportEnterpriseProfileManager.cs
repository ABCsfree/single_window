namespace TradeXmlStudio.Core;

public sealed class ExportEnterpriseProfileManager
{
    private readonly List<ExportEnterpriseProfile> _profiles;
    private int _selectedIndex;

    public IReadOnlyList<ExportEnterpriseProfile> Profiles => _profiles;
    public ExportEnterpriseProfile Selected => _profiles[_selectedIndex];

    public ExportEnterpriseProfileManager(TradeXmlOptions options)
    {
        _profiles = (options.ExportEnterpriseProfiles ?? [])
            .Select(profile => profile with { Enterprise = (profile.Enterprise ?? new()) with { } })
            .ToList();
        if (_profiles.Count == 0)
        {
            _profiles.Add(new ExportEnterpriseProfile
            {
                Name = "默认方案",
                Enterprise = options.ExportEnterprise with { },
                SupervisingCustomsCode = options.SupervisingCustomsCode
            });
        }
        _selectedIndex = Math.Max(0, _profiles.FindIndex(profile =>
            profile.Id == options.SelectedExportEnterpriseProfileId));
    }

    public void UpdateCurrent(EnterpriseOptions enterprise, string supervisingCustomsCode)
    {
        _profiles[_selectedIndex] = Selected with
        {
            Enterprise = enterprise with { },
            SupervisingCustomsCode = supervisingCustomsCode
        };
    }

    public void Add(string name)
    {
        name = ValidateName(name);
        _profiles.Add(new ExportEnterpriseProfile { Name = name });
        _selectedIndex = _profiles.Count - 1;
    }

    public void Rename(string name)
    {
        _profiles[_selectedIndex] = Selected with { Name = ValidateName(name, Selected.Id) };
    }

    public void Select(string id)
    {
        var index = _profiles.FindIndex(profile => profile.Id == id);
        if (index < 0)
        {
            throw new ArgumentException("出口企业方案不存在。", nameof(id));
        }
        _selectedIndex = index;
    }

    public void ApplyTo(TradeXmlOptions options)
    {
        options.ExportEnterpriseProfiles = _profiles
            .Select(profile => profile with { Enterprise = profile.Enterprise with { } }).ToList();
        options.SelectedExportEnterpriseProfileId = Selected.Id;
        options.ExportEnterprise = Selected.Enterprise with { };
        options.SupervisingCustomsCode = Selected.SupervisingCustomsCode;
    }

    private string ValidateName(string name, string? currentId = null)
    {
        name = name.Trim();
        if (name.Length == 0)
        {
            throw new ArgumentException("方案名称不能为空。");
        }
        if (_profiles.Any(profile => profile.Id != currentId
            && string.Equals(profile.Name, name, StringComparison.OrdinalIgnoreCase)))
        {
            throw new ArgumentException("方案名称已存在，请使用其他名称。");
        }
        return name;
    }
}
