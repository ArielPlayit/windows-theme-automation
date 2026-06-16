using ThemeAutomation.Core.Configuration;

namespace ThemeAutomation.Core.Scheduling;

public static class ScheduleEvaluator
{
    public static ThemeProfile SelectProfile(AutomationConfig config, TimeOnly currentTime)
    {
        var inDayWindow = IsInWindow(currentTime, config.DayStart, config.NightStart);
        return inDayWindow ? config.DayProfile : config.NightProfile;
    }

    public static bool ShouldApplyAutomatically(AutomationConfig config, DateTimeOffset now) =>
        config.IsAutomationEnabledOn(now.LocalDateTime.DayOfWeek);

    private static bool IsInWindow(TimeOnly currentTime, TimeOnly start, TimeOnly end)
    {
        if (start == end)
        {
            return true;
        }

        if (start < end)
        {
            return currentTime >= start && currentTime < end;
        }

        return currentTime >= start || currentTime < end;
    }
}
