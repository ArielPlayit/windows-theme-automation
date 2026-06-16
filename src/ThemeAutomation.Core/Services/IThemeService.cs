using ThemeAutomation.Core.Configuration;

namespace ThemeAutomation.Core.Services;

public interface IThemeService
{
    ApplyResult Apply(ThemeMode mode);
}
