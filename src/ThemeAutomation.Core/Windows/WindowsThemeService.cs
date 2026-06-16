using Microsoft.Win32;
using ThemeAutomation.Core.Configuration;
using ThemeAutomation.Core.Services;

namespace ThemeAutomation.Core.Windows;

public sealed class WindowsThemeService : IThemeService
{
    private const string PersonalizeKey = @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";

    public ApplyResult Apply(ThemeMode mode)
    {
        try
        {
            using var key = Registry.CurrentUser.CreateSubKey(PersonalizeKey);
            if (key is null)
            {
                return new ApplyResult(false, "Could not open the Windows theme registry key.");
            }

            var value = mode == ThemeMode.Light ? 1 : 0;
            key.SetValue("AppsUseLightTheme", value, RegistryValueKind.DWord);
            key.SetValue("SystemUsesLightTheme", value, RegistryValueKind.DWord);
            WindowsSettingsNotifier.Broadcast("ImmersiveColorSet");

            return new ApplyResult(true, $"Applied {mode} theme.");
        }
        catch (Exception ex)
        {
            return new ApplyResult(false, $"Could not apply theme: {ex.Message}");
        }
    }
}
