using ThemeAutomation.Core.Configuration;
using ThemeAutomation.Core.NightLight;
using ThemeAutomation.Core.Scheduling;
using ThemeAutomation.Core;
using ThemeAutomation.Core.Services;
using ThemeAutomation.Core.Windows;

var tests = new List<(string Name, Action Body)>
{
    ("06:59 uses night profile", () =>
    {
        var config = AutomationConfig.CreateDefault();
        var profile = ScheduleEvaluator.SelectProfile(config, new TimeOnly(6, 59));
        Assert.Equal(ProfileKind.Night, profile.Kind);
    }),
    ("07:00 uses day profile", () =>
    {
        var config = AutomationConfig.CreateDefault();
        var profile = ScheduleEvaluator.SelectProfile(config, new TimeOnly(7, 0));
        Assert.Equal(ProfileKind.Day, profile.Kind);
    }),
    ("18:59 uses day profile", () =>
    {
        var config = AutomationConfig.CreateDefault();
        var profile = ScheduleEvaluator.SelectProfile(config, new TimeOnly(18, 59));
        Assert.Equal(ProfileKind.Day, profile.Kind);
    }),
    ("19:00 uses night profile", () =>
    {
        var config = AutomationConfig.CreateDefault();
        var profile = ScheduleEvaluator.SelectProfile(config, new TimeOnly(19, 0));
        Assert.Equal(ProfileKind.Night, profile.Kind);
    }),
    ("Default automation config enables every day with safe logon behavior", () =>
    {
        var config = AutomationConfig.CreateDefault();

        Assert.Equal(AutomationDays.All, config.ActiveDays);
        Assert.True(config.ApplyOnLogon);
        Assert.True(config.CatchUpMissedSwitch);
        Assert.Equal(30, config.NormalizedLogonDelaySeconds);
        Assert.True(config.IsAutomationEnabledOn(DayOfWeek.Monday));
        Assert.True(config.IsAutomationEnabledOn(DayOfWeek.Sunday));
    }),
    ("Old config JSON loads Automation Plus defaults", () =>
    {
        var tempPath = Path.Combine(Path.GetTempPath(), "ThemeAutomationTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempPath);
        File.WriteAllText(Path.Combine(tempPath, "config.json"), """
        {
          "DayStart": "07:00:00",
          "NightStart": "19:00:00",
          "DayProfile": {
            "Kind": "Day",
            "ThemeMode": "Light",
            "NightLightWarmth": 20
          },
          "NightProfile": {
            "Kind": "Night",
            "ThemeMode": "Dark",
            "NightLightWarmth": 50
          },
          "EnableFallbackFilter": true,
          "LogLevel": "Information"
        }
        """);

        var config = new JsonConfigurationStore(tempPath).Load();

        Assert.Equal(AutomationDays.All, config.ActiveDays);
        Assert.True(config.ApplyOnLogon);
        Assert.True(config.CatchUpMissedSwitch);
        Assert.Equal(30, config.NormalizedLogonDelaySeconds);
    }),
    ("Logon delay is normalized to a safe range", () =>
    {
        var highDelay = AutomationConfig.CreateDefault() with { LogonDelaySeconds = 9999 };
        var lowDelay = AutomationConfig.CreateDefault() with { LogonDelaySeconds = -5 };

        Assert.Equal(600, highDelay.NormalizedLogonDelaySeconds);
        Assert.Equal(0, lowDelay.NormalizedLogonDelaySeconds);
    }),
    ("Automatic apply skips inactive days", () =>
    {
        var config = AutomationConfig.CreateDefault() with { ActiveDays = AutomationDays.Tuesday };
        var theme = new RecordingThemeService();
        var nightLight = new RecordingNightLightService();
        var runner = new AutomationRunner(new InMemoryConfigurationStore(config), theme, nightLight);

        var summary = runner.Apply(new DateTimeOffset(2026, 6, 15, 8, 0, 0, TimeSpan.Zero), AutomationTrigger.Automatic);

        Assert.False(summary.Applied);
        Assert.Equal(0, theme.ApplyCount);
        Assert.Equal(0, nightLight.ApplyCount);
    }),
    ("Automatic apply runs on active days", () =>
    {
        var config = AutomationConfig.CreateDefault() with { ActiveDays = AutomationDays.Monday };
        var theme = new RecordingThemeService();
        var nightLight = new RecordingNightLightService();
        var runner = new AutomationRunner(new InMemoryConfigurationStore(config), theme, nightLight);

        var summary = runner.Apply(new DateTimeOffset(2026, 6, 15, 8, 0, 0, TimeSpan.Zero), AutomationTrigger.Automatic);

        Assert.True(summary.Applied);
        Assert.Equal(1, theme.ApplyCount);
        Assert.Equal(1, nightLight.ApplyCount);
    }),
    ("Manual apply runs even when current day is inactive", () =>
    {
        var config = AutomationConfig.CreateDefault() with { ActiveDays = AutomationDays.Tuesday };
        var theme = new RecordingThemeService();
        var nightLight = new RecordingNightLightService();
        var runner = new AutomationRunner(new InMemoryConfigurationStore(config), theme, nightLight);

        var summary = runner.Apply(new DateTimeOffset(2026, 6, 15, 8, 0, 0, TimeSpan.Zero), AutomationTrigger.Manual);

        Assert.True(summary.Applied);
        Assert.Equal(1, theme.ApplyCount);
        Assert.Equal(1, nightLight.ApplyCount);
    }),
    ("Scheduler install arguments respect logon settings and delay", () =>
    {
        var config = AutomationConfig.CreateDefault() with { ApplyOnLogon = true, LogonDelaySeconds = 95 };

        var arguments = WindowsSchedulerService.BuildInstallArguments("C:\\Tools\\themeauto.exe", config);

        Assert.True(arguments.Any(argument => argument.Contains("/SC DAILY /ST 07:00", StringComparison.Ordinal)));
        Assert.True(arguments.Any(argument => argument.Contains("/SC DAILY /ST 19:00", StringComparison.Ordinal)));
        Assert.True(arguments.Any(argument => argument.Contains("/SC ONLOGON /DELAY 0001:35", StringComparison.Ordinal)));
    }),
    ("Scheduler skips logon task when disabled", () =>
    {
        var config = AutomationConfig.CreateDefault() with { ApplyOnLogon = false };

        var arguments = WindowsSchedulerService.BuildInstallArguments("C:\\Tools\\themeauto.exe", config);

        Assert.False(arguments.Any(argument => argument.Contains("/SC ONLOGON", StringComparison.Ordinal)));
    }),
    ("Night Light blob update changes every CF 28 marker and preserves other bytes", () =>
    {
        byte[] data =
        [
            0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08,
            0x09, 0x0A, 0x10, 0x20, 0x30, 0x40, 0x50, 0x60,
            0xCF, 0x28, 0x80, 0x55, 0x03, 0x04,
            0xCF, 0x28, 0x80, 0x55, 0x05
        ];

        var result = NightLightBlobEditor.TrySetTemperature(data, 50);

        Assert.True(result.Updated);
        Assert.Equal(2, result.MarkersUpdated);
        Assert.True(result.TimestampBumped);
        Assert.Equal((byte)0x11, data[10]);
        Assert.Equal((byte)0xAA, data[18]);
        Assert.Equal((byte)0x3B, data[19]);
        Assert.Equal((byte)0xAA, data[24]);
        Assert.Equal((byte)0x3B, data[25]);
        Assert.Equal((byte)0x01, data[0]);
        Assert.Equal((byte)0x05, data[26]);
    }),
    ("Night Light blob update reports when timestamp cannot be bumped", () =>
    {
        byte[] data =
        [
            0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08,
            0x09, 0x0A, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x60,
            0xCF, 0x28, 0x80, 0x55
        ];

        var result = NightLightBlobEditor.TrySetTemperature(data, 20);

        Assert.True(result.Updated);
        Assert.False(result.TimestampBumped);
        Assert.Contains("timestamp", result.Message);
    }),
    ("Night Light timestamp bump does not overwrite temperature bytes", () =>
    {
        byte[] data =
        [
            0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08,
            0xCF, 0x28, 0x80, 0x55, 0x10, 0x20, 0x30, 0x40
        ];

        var result = NightLightBlobEditor.TrySetTemperature(data, 50);

        Assert.True(result.Updated);
        Assert.True(result.TimestampBumped);
        Assert.Equal((byte)0xAA, data[10]);
        Assert.Equal((byte)0x3B, data[11]);
        Assert.Equal((byte)0x11, data[12]);
    }),
    ("Night Light blob update reports missing marker", () =>
    {
        byte[] data = [0x01, 0x02, 0x03, 0x04];

        var result = NightLightBlobEditor.TrySetTemperature(data, 50);

        Assert.False(result.Updated);
        Assert.Equal("No CF 28 temperature marker was found.", result.Message);
    }),
    ("Fallback decision skips gamma when native succeeds", () =>
    {
        var decision = FallbackDecision.Decide(
            nativeRegistryUpdated: true,
            nativeVisuallyRefreshed: true,
            fallbackEnabled: true);

        Assert.False(decision.ApplyFallback);
        Assert.Equal(NightLightApplyStatus.NativeApplied, decision.Status);
    }),
    ("Fallback decision applies gamma when native is stale and fallback is enabled", () =>
    {
        var decision = FallbackDecision.Decide(
            nativeRegistryUpdated: true,
            nativeVisuallyRefreshed: false,
            fallbackEnabled: true);

        Assert.True(decision.ApplyFallback);
        Assert.Equal(NightLightApplyStatus.FallbackApplied, decision.Status);
        Assert.Equal(NativeNightLightRefreshStatus.NativeRegistryOnly, decision.NativeRefreshStatus);
    }),
    ("Fallback decision reports registry-only when native is stale and fallback is disabled", () =>
    {
        var decision = FallbackDecision.Decide(
            nativeRegistryUpdated: true,
            nativeVisuallyRefreshed: false,
            fallbackEnabled: false);

        Assert.False(decision.ApplyFallback);
        Assert.Equal(NightLightApplyStatus.NativeRegistryOnly, decision.Status);
        Assert.Equal(NativeNightLightRefreshStatus.NativeRegistryOnly, decision.NativeRefreshStatus);
    }),
    ("Fallback decision reports degraded when registry update fails and fallback is disabled", () =>
    {
        var decision = FallbackDecision.Decide(
            nativeRegistryUpdated: false,
            nativeVisuallyRefreshed: false,
            fallbackEnabled: false);

        Assert.False(decision.ApplyFallback);
        Assert.Equal(NightLightApplyStatus.Degraded, decision.Status);
    }),
    ("WPF shell contains modern Figma-approved sections", () =>
    {
        var xaml = File.ReadAllText(FindRepoFile("src/ThemeAutomation.App/MainWindow.xaml"));

        Assert.Contains("Sidebar", xaml);
        Assert.Contains("ProfileCard", xaml);
        Assert.Contains("Night Light diagnostics", xaml);
        Assert.Contains("Install_Click", xaml);
        Assert.Contains("&#x2600;", xaml);
        Assert.Contains("CrescentIcon", xaml);
        Assert.Contains("AppliedOverlay", xaml);
        Assert.Contains("CloseAppliedOverlay_Click", xaml);
    }),
    ("WPF sidebar exposes clickable views with integrated feedback", () =>
    {
        var xaml = File.ReadAllText(FindRepoFile("src/ThemeAutomation.App/MainWindow.xaml"));
        var codeBehind = File.ReadAllText(FindRepoFile("src/ThemeAutomation.App/MainWindow.xaml.cs"));

        Assert.Contains("SidebarNavButton", xaml);
        Assert.Contains("ProfilesNav_Click", xaml);
        Assert.Contains("DiagnosticsNav_Click", xaml);
        Assert.Contains("SchedulerNav_Click", xaml);
        Assert.Contains("ProfilesView", xaml);
        Assert.Contains("DiagnosticsView", xaml);
        Assert.Contains("SchedulerView", xaml);
        Assert.Contains("FeedbackOverlay", xaml);
        Assert.Contains("ProfilesViewVisibility", codeBehind);
        Assert.Contains("DiagnosticsViewVisibility", codeBehind);
        Assert.Contains("SchedulerViewVisibility", codeBehind);
        Assert.Contains("ProfilesNav_Click", codeBehind);
        Assert.Contains("DiagnosticsNav_Click", codeBehind);
        Assert.Contains("SchedulerNav_Click", codeBehind);
        Assert.DoesNotContain("MessageBox.Show", codeBehind);
    }),
    ("WPF scheduler exposes Automation Plus controls", () =>
    {
        var xaml = File.ReadAllText(FindRepoFile("src/ThemeAutomation.App/MainWindow.xaml"));
        var codeBehind = File.ReadAllText(FindRepoFile("src/ThemeAutomation.App/MainWindow.xaml.cs"));
        var cli = File.ReadAllText(FindRepoFile("src/ThemeAutomation.Cli/Program.cs"));

        Assert.Contains("Active days", xaml);
        Assert.Contains("MondayActive", xaml);
        Assert.Contains("SundayActive", xaml);
        Assert.Contains("ApplyOnLogon", xaml);
        Assert.Contains("LogonDelaySecondsText", xaml);
        Assert.Contains("CatchUpMissedSwitch", xaml);
        Assert.Contains("ActiveDaysSummary", codeBehind);
        Assert.Contains("LogonDelaySecondsText", codeBehind);
        Assert.Contains("Automation days:", cli);
        Assert.Contains("Catch-up missed switch:", cli);
        Assert.Contains("Night Light native refresh:", cli);
        Assert.Contains("Night Light services:", cli);
    }),
    ("WPF scheduler uses styled scroll spacing", () =>
    {
        var xaml = File.ReadAllText(FindRepoFile("src/ThemeAutomation.App/MainWindow.xaml"));

        Assert.Contains("ModernScrollViewer", xaml);
        Assert.Contains("ModernVerticalScrollBar", xaml);
        Assert.Contains("SchedulerScrollContent", xaml);
        Assert.Contains("Margin=\"0,0,18,0\"", xaml);
        Assert.Contains("Width=\"8\"", xaml);
        Assert.Contains("CornerRadius=\"4\"", xaml);
    }),
    ("WPF styles avoid invalid Resources setter", () =>
    {
        var xaml = File.ReadAllText(FindRepoFile("src/ThemeAutomation.App/MainWindow.xaml"));

        Assert.DoesNotContain("Setter Property=\"Resources\"", xaml);
    })
};

var failures = 0;
foreach (var test in tests)
{
    try
    {
        test.Body();
        Console.WriteLine($"PASS {test.Name}");
    }
    catch (Exception ex)
    {
        failures++;
        Console.WriteLine($"FAIL {test.Name}");
        Console.WriteLine($"     {ex.Message}");
    }
}

if (failures > 0)
{
    Console.WriteLine($"{failures} test(s) failed.");
    return 1;
}

Console.WriteLine($"{tests.Count} tests passed.");
return 0;

static string FindRepoFile(string relativePath)
{
    var current = new DirectoryInfo(AppContext.BaseDirectory);
    while (current is not null)
    {
        var candidate = Path.Combine(current.FullName, relativePath);
        if (File.Exists(candidate))
        {
            return candidate;
        }

        current = current.Parent;
    }

    throw new FileNotFoundException($"Could not find {relativePath} from {AppContext.BaseDirectory}.");
}

static class Assert
{
    public static void Equal<T>(T expected, T actual)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual))
        {
            throw new InvalidOperationException($"Expected {expected}, got {actual}.");
        }
    }

    public static void True(bool value)
    {
        if (!value)
        {
            throw new InvalidOperationException("Expected true, got false.");
        }
    }

    public static void False(bool value)
    {
        if (value)
        {
            throw new InvalidOperationException("Expected false, got true.");
        }
    }

    public static void Contains(string expected, string actual)
    {
        if (!actual.Contains(expected, StringComparison.Ordinal))
        {
            throw new InvalidOperationException($"Expected text to contain '{expected}'.");
        }
    }

    public static void DoesNotContain(string unexpected, string actual)
    {
        if (actual.Contains(unexpected, StringComparison.Ordinal))
        {
            throw new InvalidOperationException($"Expected text not to contain '{unexpected}'.");
        }
    }
}

sealed class InMemoryConfigurationStore : IConfigurationStore
{
    private readonly AutomationConfig _config;

    public InMemoryConfigurationStore(AutomationConfig config)
    {
        _config = config;
    }

    public string ConfigPath => "memory";

    public AutomationConfig Load() => _config;

    public void Save(AutomationConfig config)
    {
    }
}

sealed class RecordingThemeService : IThemeService
{
    public int ApplyCount { get; private set; }

    public ApplyResult Apply(ThemeMode mode)
    {
        ApplyCount++;
        return new ApplyResult(true, $"Theme {mode} applied.");
    }
}

sealed class RecordingNightLightService : INightLightService
{
    public int ApplyCount { get; private set; }

    public NightLightServiceResult Apply(int percentage, bool fallbackEnabled)
    {
        ApplyCount++;
        return new NightLightServiceResult(NightLightApplyStatus.NativeApplied, true, false, []);
    }

    public NightLightDiagnostics Diagnose() => new(true, 1, 1, []);
}
