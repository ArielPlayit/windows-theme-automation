using System.Text.Json;
using System.Text.Json.Serialization;

namespace ThemeAutomation.Core.Configuration;

public sealed class JsonConfigurationStore : IConfigurationStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() }
    };

    public JsonConfigurationStore(string? baseDirectory = null)
    {
        BaseDirectory = baseDirectory ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "WindowsThemeAuto");
        ConfigPath = Path.Combine(BaseDirectory, "config.json");
    }

    public string BaseDirectory { get; }

    public string ConfigPath { get; }

    public AutomationConfig Load()
    {
        if (!File.Exists(ConfigPath))
        {
            var defaultConfig = AutomationConfig.CreateDefault();
            Save(defaultConfig);
            return defaultConfig;
        }

        var json = File.ReadAllText(ConfigPath);
        return (JsonSerializer.Deserialize<AutomationConfig>(json, JsonOptions)
                ?? AutomationConfig.CreateDefault())
            .Normalize();
    }

    public void Save(AutomationConfig config)
    {
        Directory.CreateDirectory(BaseDirectory);
        File.WriteAllText(ConfigPath, JsonSerializer.Serialize(config.Normalize(), JsonOptions));
    }
}
