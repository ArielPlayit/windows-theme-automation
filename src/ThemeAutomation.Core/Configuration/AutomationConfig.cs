namespace ThemeAutomation.Core.Configuration;

public sealed record AutomationConfig
{
    public const int MinLogonDelaySeconds = 0;
    public const int MaxLogonDelaySeconds = 600;

    public AutomationConfig()
    {
    }

    public AutomationConfig(
        TimeOnly dayStart,
        TimeOnly nightStart,
        ThemeProfile dayProfile,
        ThemeProfile nightProfile,
        bool enableFallbackFilter,
        string logLevel)
    {
        DayStart = dayStart;
        NightStart = nightStart;
        DayProfile = dayProfile;
        NightProfile = nightProfile;
        EnableFallbackFilter = enableFallbackFilter;
        LogLevel = logLevel;
    }

    public TimeOnly DayStart { get; init; } = new(7, 0);

    public TimeOnly NightStart { get; init; } = new(19, 0);

    public ThemeProfile DayProfile { get; init; } = new(ProfileKind.Day, ThemeMode.Light, 20);

    public ThemeProfile NightProfile { get; init; } = new(ProfileKind.Night, ThemeMode.Dark, 50);

    public bool EnableFallbackFilter { get; init; } = true;

    public string LogLevel { get; init; } = "Information";

    public AutomationDays ActiveDays { get; init; } = AutomationDays.All;

    public bool ApplyOnLogon { get; init; } = true;

    public int LogonDelaySeconds { get; init; } = 30;

    public bool CatchUpMissedSwitch { get; init; } = true;

    public int NormalizedLogonDelaySeconds => Math.Clamp(LogonDelaySeconds, MinLogonDelaySeconds, MaxLogonDelaySeconds);

    public static AutomationConfig CreateDefault() => new();

    public bool IsAutomationEnabledOn(DayOfWeek dayOfWeek) => ActiveDays.HasFlag(ToAutomationDay(dayOfWeek));

    public AutomationConfig Normalize() => this with
    {
        ActiveDays = ActiveDays & AutomationDays.All,
        LogonDelaySeconds = NormalizedLogonDelaySeconds,
        LogLevel = string.IsNullOrWhiteSpace(LogLevel) ? "Information" : LogLevel
    };

    private static AutomationDays ToAutomationDay(DayOfWeek dayOfWeek) => dayOfWeek switch
    {
        DayOfWeek.Sunday => AutomationDays.Sunday,
        DayOfWeek.Monday => AutomationDays.Monday,
        DayOfWeek.Tuesday => AutomationDays.Tuesday,
        DayOfWeek.Wednesday => AutomationDays.Wednesday,
        DayOfWeek.Thursday => AutomationDays.Thursday,
        DayOfWeek.Friday => AutomationDays.Friday,
        DayOfWeek.Saturday => AutomationDays.Saturday,
        _ => AutomationDays.None
    };
}

public sealed record ThemeProfile(ProfileKind Kind, ThemeMode ThemeMode, int NightLightWarmth)
{
    public int ClampedWarmth => Math.Clamp(NightLightWarmth, 0, 100);
}

[Flags]
public enum AutomationDays
{
    None = 0,
    Sunday = 1 << 0,
    Monday = 1 << 1,
    Tuesday = 1 << 2,
    Wednesday = 1 << 3,
    Thursday = 1 << 4,
    Friday = 1 << 5,
    Saturday = 1 << 6,
    All = Sunday | Monday | Tuesday | Wednesday | Thursday | Friday | Saturday
}

public enum ProfileKind
{
    Day,
    Night
}

public enum ThemeMode
{
    Light,
    Dark
}
