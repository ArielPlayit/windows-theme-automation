namespace ThemeAutomation.Core.Configuration;

public interface IConfigurationStore
{
    string ConfigPath { get; }

    AutomationConfig Load();

    void Save(AutomationConfig config);
}
