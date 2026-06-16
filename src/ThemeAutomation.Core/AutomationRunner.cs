using ThemeAutomation.Core.Configuration;
using ThemeAutomation.Core.NightLight;
using ThemeAutomation.Core.Scheduling;
using ThemeAutomation.Core.Services;

namespace ThemeAutomation.Core;

public sealed class AutomationRunner
{
    private readonly IConfigurationStore _configurationStore;
    private readonly IThemeService _themeService;
    private readonly INightLightService _nightLightService;

    public AutomationRunner(
        IConfigurationStore configurationStore,
        IThemeService themeService,
        INightLightService nightLightService)
    {
        _configurationStore = configurationStore;
        _themeService = themeService;
        _nightLightService = nightLightService;
    }

    public AutomationApplySummary Apply(DateTimeOffset now, AutomationTrigger trigger = AutomationTrigger.Automatic)
    {
        var config = _configurationStore.Load().Normalize();
        var profile = ScheduleEvaluator.SelectProfile(config, TimeOnly.FromDateTime(now.LocalDateTime));
        var skipReason = GetSkipReason(config, now, trigger);
        if (skipReason is not null)
        {
            return new AutomationApplySummary(
                profile.Kind,
                new ApplyResult(true, skipReason),
                new NightLightServiceResult(NightLightApplyStatus.Skipped, false, false, [skipReason]),
                Applied: false,
                Message: skipReason);
        }

        var theme = _themeService.Apply(profile.ThemeMode);
        var nightLight = _nightLightService.Apply(profile.ClampedWarmth, config.EnableFallbackFilter);

        return new AutomationApplySummary(profile.Kind, theme, nightLight, Applied: true, Message: "Profile applied.");
    }

    private static string? GetSkipReason(AutomationConfig config, DateTimeOffset now, AutomationTrigger trigger)
    {
        if (trigger == AutomationTrigger.Manual)
        {
            return null;
        }

        if (!ScheduleEvaluator.ShouldApplyAutomatically(config, now))
        {
            return $"Automation is disabled for {now.LocalDateTime.DayOfWeek}.";
        }

        if (trigger == AutomationTrigger.Logon && !config.ApplyOnLogon)
        {
            return "Apply on logon is disabled.";
        }

        if (trigger == AutomationTrigger.Logon && !config.CatchUpMissedSwitch)
        {
            return "Catch-up on missed switches is disabled.";
        }

        return null;
    }
}

public sealed record AutomationApplySummary(
    ProfileKind ProfileKind,
    ApplyResult ThemeResult,
    NightLightServiceResult NightLightResult,
    bool Applied = true,
    string Message = "Profile applied.");

public enum AutomationTrigger
{
    Automatic,
    Manual,
    Logon
}
