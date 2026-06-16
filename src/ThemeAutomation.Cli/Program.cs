using System.Diagnostics;
using ThemeAutomation.Core;
using ThemeAutomation.Core.Configuration;
using ThemeAutomation.Core.Diagnostics;
using ThemeAutomation.Core.NightLight;
using ThemeAutomation.Core.Scheduling;
using ThemeAutomation.Core.Services;
using ThemeAutomation.Core.Windows;

var command = args.Length == 0 ? "status" : args[0].ToLowerInvariant();
var services = CliServices.Create();

return command switch
{
    "apply" => Apply(services, args),
    "install" => Install(services),
    "uninstall" => Uninstall(services),
    "status" => Status(services),
    "diagnose" => Diagnose(services),
    "help" or "-h" or "--help" => Help(),
    _ => Unknown(command)
};

static int Apply(CliServices services, string[] args)
{
    var trigger = ParseTrigger(args);
    var summary = services.Runner.Apply(DateTimeOffset.Now, trigger);
    services.Logger.Info($"Applied profile {summary.ProfileKind}; theme success={summary.ThemeResult.Success}; night light={summary.NightLightResult.Status}.");
    Console.WriteLine($"Trigger: {trigger}");
    Console.WriteLine($"Applied: {(summary.Applied ? "yes" : "no")}");
    Console.WriteLine($"Profile: {summary.ProfileKind}");
    Console.WriteLine($"Theme: {summary.ThemeResult.Message}");
    Console.WriteLine($"Night Light: {summary.NightLightResult.Status}");
    foreach (var message in summary.NightLightResult.Messages)
    {
        Console.WriteLine($"- {message}");
    }

    return summary.ThemeResult.Success && summary.NightLightResult.Status != NightLightApplyStatus.Degraded ? 0 : 1;
}

static int Install(CliServices services)
{
    var config = services.ConfigurationStore.Load();
    var executablePath = ResolveExecutablePath();
    var result = services.SchedulerService.Install(executablePath, config);
    services.Logger.Info($"Install requested for executable {executablePath}; success={result.Success}.");
    foreach (var message in result.Messages)
    {
        Console.WriteLine(message);
    }

    Console.WriteLine(result.Success
        ? $"Installed scheduled tasks for {config.DayStart:HH\\:mm}, {config.NightStart:HH\\:mm}{(config.ApplyOnLogon ? ", and logon" : string.Empty)}."
        : "Could not install one or more scheduled tasks.");

    return result.Success ? 0 : 1;
}

static int Uninstall(CliServices services)
{
    var result = services.SchedulerService.Uninstall();
    services.Logger.Info($"Uninstall requested; success={result.Success}.");
    foreach (var message in result.Messages)
    {
        Console.WriteLine(message);
    }

    Console.WriteLine("Removed Windows Theme Automation scheduled tasks.");
    return result.Success ? 0 : 1;
}

static int Status(CliServices services)
{
    var config = services.ConfigurationStore.Load();
    var profile = ScheduleEvaluator.SelectProfile(config, TimeOnly.FromDateTime(DateTime.Now));
    var diagnostics = services.NightLightService.Diagnose();
    var schedulerStatus = services.SchedulerService.GetStatus();

    Console.WriteLine("Windows Theme Automation");
    Console.WriteLine($"Config: {services.ConfigurationStore.ConfigPath}");
    Console.WriteLine($"Current profile: {profile.Kind}");
    Console.WriteLine($"Theme: {profile.ThemeMode}");
    Console.WriteLine($"Night Light warmth: {profile.ClampedWarmth}%");
    Console.WriteLine($"Fallback filter: {(config.EnableFallbackFilter ? "enabled" : "disabled")}");
    Console.WriteLine($"Automation days: {FormatAutomationDays(config.ActiveDays)}");
    Console.WriteLine($"Apply on logon: {(config.ApplyOnLogon ? "enabled" : "disabled")}");
    Console.WriteLine($"Logon delay seconds: {config.NormalizedLogonDelaySeconds}");
    Console.WriteLine($"Catch-up missed switch: {(config.CatchUpMissedSwitch ? "enabled" : "disabled")}");
    Console.WriteLine($"Night Light settings blobs: {diagnostics.SettingsBlobCount}");
    Console.WriteLine($"Night Light state blobs: {diagnostics.StateBlobCount}");
    Console.WriteLine($"Night Light native refresh: {diagnostics.NativeRefreshStatus}");
    Console.WriteLine($"Night Light services: {FormatDependencyServices(diagnostics.Services)}");
    Console.WriteLine("Scheduled tasks:");
    foreach (var task in schedulerStatus.Tasks)
    {
        Console.WriteLine($"- {task.Name}: {(task.Exists ? task.State ?? "exists" : "missing")}");
        if (task.Exists)
        {
            Console.WriteLine($"  Next: {task.NextRunTime?.ToString() ?? "n/a"}");
            Console.WriteLine($"  Action: {task.Action ?? "n/a"}");
        }
    }

    return 0;
}

static int Diagnose(CliServices services)
{
    var diagnostics = services.NightLightService.Diagnose();
    Console.WriteLine($"Default settings initialized: {diagnostics.HasDefaultSettings}");
    Console.WriteLine($"Windows version: {diagnostics.WindowsVersion}");
    Console.WriteLine($"Settings blobs: {diagnostics.SettingsBlobCount}");
    Console.WriteLine($"State blobs: {diagnostics.StateBlobCount}");
    Console.WriteLine($"Native refresh: {diagnostics.NativeRefreshStatus}");
    Console.WriteLine("Dependency services:");
    foreach (var service in diagnostics.Services)
    {
        Console.WriteLine($"- {service.Name}: {service.State} - {service.Message}");
    }

    Console.WriteLine("Registry paths:");
    foreach (var path in diagnostics.RegistryPaths)
    {
        Console.WriteLine($"- HKCU\\{path}");
    }

    return diagnostics.SettingsBlobCount > 0 ? 0 : 1;
}

static int Help()
{
    Console.WriteLine("""
    themeauto apply      Apply the current day/night profile for automation
    themeauto install    Create scheduled tasks
    themeauto uninstall  Remove scheduled tasks
    themeauto status     Show current config, tasks, and Night Light diagnostics
    themeauto diagnose   Show Night Light registry diagnostics
    """);
    return 0;
}

static int Unknown(string command)
{
    Console.Error.WriteLine($"Unknown command: {command}");
    Help();
    return 2;
}

static string ResolveExecutablePath()
{
    var appHost = Path.Combine(AppContext.BaseDirectory, "themeauto.exe");
    if (File.Exists(appHost))
    {
        return appHost;
    }

    return Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName ?? "themeauto.exe";
}

static AutomationTrigger ParseTrigger(string[] args)
{
    for (var index = 1; index < args.Length; index++)
    {
        if (args[index].Equals("--manual", StringComparison.OrdinalIgnoreCase))
        {
            return AutomationTrigger.Manual;
        }

        if (args[index].Equals("--trigger", StringComparison.OrdinalIgnoreCase) &&
            index + 1 < args.Length &&
            Enum.TryParse<AutomationTrigger>(args[index + 1], ignoreCase: true, out var trigger))
        {
            return trigger;
        }
    }

    return AutomationTrigger.Automatic;
}

static string FormatAutomationDays(AutomationDays activeDays)
{
    if (activeDays == AutomationDays.All)
    {
        return "Every day";
    }

    if (activeDays == AutomationDays.None)
    {
        return "No days";
    }

    var names = new List<string>();
    foreach (var day in Enum.GetValues<AutomationDays>())
    {
        if (day is AutomationDays.None or AutomationDays.All)
        {
            continue;
        }

        if (activeDays.HasFlag(day))
        {
            names.Add(day.ToString());
        }
    }

    return string.Join(", ", names);
}

static string FormatDependencyServices(IReadOnlyList<NightLightDependencyServiceStatus> services)
{
    if (services.Count == 0)
    {
        return "not checked";
    }

    var unhealthy = services.Where(service => !service.IsHealthy).Select(service => $"{service.Name} {service.State}").ToList();
    return unhealthy.Count == 0
        ? "healthy"
        : string.Join(", ", unhealthy);
}

internal sealed class CliServices
{
    private CliServices(
        IConfigurationStore configurationStore,
        INightLightService nightLightService,
        ISchedulerService schedulerService,
        AutomationRunner runner,
        FileLogger logger)
    {
        ConfigurationStore = configurationStore;
        NightLightService = nightLightService;
        SchedulerService = schedulerService;
        Runner = runner;
        Logger = logger;
    }

    public IConfigurationStore ConfigurationStore { get; }

    public INightLightService NightLightService { get; }

    public ISchedulerService SchedulerService { get; }

    public AutomationRunner Runner { get; }

    public FileLogger Logger { get; }

    public static CliServices Create()
    {
        var configurationStore = new JsonConfigurationStore();
        var fallbackService = new GammaFallbackService();
        var nightLightService = new RegistryNightLightService(fallbackService);
        var themeService = new WindowsThemeService();
        var schedulerService = new WindowsSchedulerService();
        var runner = new AutomationRunner(configurationStore, themeService, nightLightService);
        var logger = new FileLogger();

        return new CliServices(configurationStore, nightLightService, schedulerService, runner, logger);
    }
}
