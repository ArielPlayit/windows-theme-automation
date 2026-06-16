using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using ThemeAutomation.Core;
using ThemeAutomation.Core.Configuration;
using ThemeAutomation.Core.Services;
using ThemeAutomation.Core.Windows;

namespace ThemeAutomation.App;

public partial class MainWindow : Window, INotifyPropertyChanged
{
    private readonly JsonConfigurationStore _configurationStore = new();
    private readonly RegistryNightLightService _nightLightService;
    private readonly WindowsSchedulerService _schedulerService = new();
    private readonly AutomationRunner _runner;
    private AutomationConfig _config;
    private AppView _activeView = AppView.Profiles;
    private string _dayStartText;
    private string _nightStartText;
    private double _dayWarmth;
    private double _nightWarmth;
    private bool _enableFallbackFilter;
    private bool _mondayActive;
    private bool _tuesdayActive;
    private bool _wednesdayActive;
    private bool _thursdayActive;
    private bool _fridayActive;
    private bool _saturdayActive;
    private bool _sundayActive;
    private bool _applyOnLogon;
    private string _logonDelaySecondsText = string.Empty;
    private bool _catchUpMissedSwitch;
    private string _diagnosticsSummary = string.Empty;
    private string _diagnosticsStatus = "Unknown";
    private string _schedulerSummary = string.Empty;
    private string _schedulerDetails = string.Empty;
    private Visibility _appliedOverlayVisibility = Visibility.Collapsed;
    private string _feedbackTitle = string.Empty;
    private string _feedbackMessage = string.Empty;
    private string _feedbackBadge = string.Empty;
    private string _feedbackDetails = string.Empty;

    public MainWindow()
    {
        InitializeComponent();

        var fallbackService = new GammaFallbackService();
        _nightLightService = new RegistryNightLightService(fallbackService);
        _runner = new AutomationRunner(
            _configurationStore,
            new WindowsThemeService(),
            _nightLightService);

        _config = _configurationStore.Load();
        _dayStartText = _config.DayStart.ToString("HH:mm");
        _nightStartText = _config.NightStart.ToString("HH:mm");
        _dayWarmth = _config.DayProfile.NightLightWarmth;
        _nightWarmth = _config.NightProfile.NightLightWarmth;
        _enableFallbackFilter = _config.EnableFallbackFilter;
        SetActiveDayFields(_config.ActiveDays);
        _applyOnLogon = _config.ApplyOnLogon;
        _logonDelaySecondsText = _config.NormalizedLogonDelaySeconds.ToString();
        _catchUpMissedSwitch = _config.CatchUpMissedSwitch;
        RefreshDiagnostics();
        RefreshSchedulerStatus();

        DataContext = this;
        SetActiveView(AppView.Profiles);
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public string ConfigPath => _configurationStore.ConfigPath;

    public string LogPath
    {
        get
        {
            var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return Path.Combine(localAppData, "WindowsThemeAuto", "logs");
        }
    }

    public Visibility ProfilesViewVisibility => _activeView == AppView.Profiles ? Visibility.Visible : Visibility.Collapsed;

    public Visibility DiagnosticsViewVisibility => _activeView == AppView.Diagnostics ? Visibility.Visible : Visibility.Collapsed;

    public Visibility SchedulerViewVisibility => _activeView == AppView.Scheduler ? Visibility.Visible : Visibility.Collapsed;

    public string DayStartText
    {
        get => _dayStartText;
        set => SetField(ref _dayStartText, value);
    }

    public string NightStartText
    {
        get => _nightStartText;
        set => SetField(ref _nightStartText, value);
    }

    public double DayWarmth
    {
        get => _dayWarmth;
        set
        {
            if (SetField(ref _dayWarmth, value))
            {
                OnPropertyChanged(nameof(DayWarmthText));
            }
        }
    }

    public double NightWarmth
    {
        get => _nightWarmth;
        set
        {
            if (SetField(ref _nightWarmth, value))
            {
                OnPropertyChanged(nameof(NightWarmthText));
            }
        }
    }

    public string DayWarmthText => $"{DayWarmth:0}%";

    public string NightWarmthText => $"{NightWarmth:0}%";

    public bool EnableFallbackFilter
    {
        get => _enableFallbackFilter;
        set => SetField(ref _enableFallbackFilter, value);
    }

    public bool MondayActive
    {
        get => _mondayActive;
        set => SetActiveDayField(ref _mondayActive, value);
    }

    public bool TuesdayActive
    {
        get => _tuesdayActive;
        set => SetActiveDayField(ref _tuesdayActive, value);
    }

    public bool WednesdayActive
    {
        get => _wednesdayActive;
        set => SetActiveDayField(ref _wednesdayActive, value);
    }

    public bool ThursdayActive
    {
        get => _thursdayActive;
        set => SetActiveDayField(ref _thursdayActive, value);
    }

    public bool FridayActive
    {
        get => _fridayActive;
        set => SetActiveDayField(ref _fridayActive, value);
    }

    public bool SaturdayActive
    {
        get => _saturdayActive;
        set => SetActiveDayField(ref _saturdayActive, value);
    }

    public bool SundayActive
    {
        get => _sundayActive;
        set => SetActiveDayField(ref _sundayActive, value);
    }

    public string ActiveDaysSummary
    {
        get
        {
            var activeDays = BuildActiveDays();
            if (activeDays == AutomationDays.All)
            {
                return "Every day";
            }

            if (activeDays == AutomationDays.None)
            {
                return "No active days";
            }

            var names = new List<string>();
            AddActiveDayName(names, activeDays, AutomationDays.Monday, "Mon");
            AddActiveDayName(names, activeDays, AutomationDays.Tuesday, "Tue");
            AddActiveDayName(names, activeDays, AutomationDays.Wednesday, "Wed");
            AddActiveDayName(names, activeDays, AutomationDays.Thursday, "Thu");
            AddActiveDayName(names, activeDays, AutomationDays.Friday, "Fri");
            AddActiveDayName(names, activeDays, AutomationDays.Saturday, "Sat");
            AddActiveDayName(names, activeDays, AutomationDays.Sunday, "Sun");
            return string.Join(", ", names);
        }
    }

    public bool ApplyOnLogon
    {
        get => _applyOnLogon;
        set => SetField(ref _applyOnLogon, value);
    }

    public string LogonDelaySecondsText
    {
        get => _logonDelaySecondsText;
        set => SetField(ref _logonDelaySecondsText, value);
    }

    public bool CatchUpMissedSwitch
    {
        get => _catchUpMissedSwitch;
        set => SetField(ref _catchUpMissedSwitch, value);
    }

    public string DiagnosticsSummary
    {
        get => _diagnosticsSummary;
        private set => SetField(ref _diagnosticsSummary, value);
    }

    public string DiagnosticsStatus
    {
        get => _diagnosticsStatus;
        private set => SetField(ref _diagnosticsStatus, value);
    }

    public string SchedulerSummary
    {
        get => _schedulerSummary;
        private set => SetField(ref _schedulerSummary, value);
    }

    public string SchedulerDetails
    {
        get => _schedulerDetails;
        private set => SetField(ref _schedulerDetails, value);
    }

    public Visibility AppliedOverlayVisibility
    {
        get => _appliedOverlayVisibility;
        private set => SetField(ref _appliedOverlayVisibility, value);
    }

    public string FeedbackTitle
    {
        get => _feedbackTitle;
        private set => SetField(ref _feedbackTitle, value);
    }

    public string FeedbackMessage
    {
        get => _feedbackMessage;
        private set => SetField(ref _feedbackMessage, value);
    }

    public string FeedbackBadge
    {
        get => _feedbackBadge;
        private set => SetField(ref _feedbackBadge, value);
    }

    public string FeedbackDetails
    {
        get => _feedbackDetails;
        private set => SetField(ref _feedbackDetails, value);
    }

    private void ProfilesNav_Click(object sender, RoutedEventArgs e) => SetActiveView(AppView.Profiles);

    private void DiagnosticsNav_Click(object sender, RoutedEventArgs e)
    {
        RefreshDiagnostics();
        SetActiveView(AppView.Diagnostics);
    }

    private void SchedulerNav_Click(object sender, RoutedEventArgs e)
    {
        RefreshSchedulerStatus();
        SetActiveView(AppView.Scheduler);
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        if (SaveConfig())
        {
            ShowFeedback(
                "Configuration saved",
                "Your day and night profile settings are ready for the next apply.",
                "Saved",
                $"Config: {ConfigPath}");
        }
    }

    private void Apply_Click(object sender, RoutedEventArgs e)
    {
        if (!SaveConfig())
        {
            return;
        }

        var summary = _runner.Apply(DateTimeOffset.Now, AutomationTrigger.Manual);
        RefreshDiagnostics();
        ShowFeedback(
            "Settings applied",
            $"{summary.ProfileKind} profile is now active.",
            "Profile active",
            $"{summary.ThemeResult.Message}\nNight Light: {summary.NightLightResult.Status}");
    }

    private void CloseAppliedOverlay_Click(object sender, RoutedEventArgs e)
    {
        AppliedOverlayVisibility = Visibility.Collapsed;
    }

    private void Status_Click(object sender, RoutedEventArgs e)
    {
        RefreshDiagnostics();
        RefreshSchedulerStatus();
        SetActiveView(AppView.Diagnostics);
        ShowFeedback(
            "Status refreshed",
            "Diagnostics and scheduled task state are up to date.",
            DiagnosticsStatus,
            $"{DiagnosticsSummary}\n\nAutomation days: {ActiveDaysSummary}\n{SchedulerDetails}");
    }

    private void RefreshDiagnostics_Click(object sender, RoutedEventArgs e)
    {
        RefreshDiagnostics();
        ShowFeedback(
            "Diagnostics refreshed",
            DiagnosticsSummary,
            DiagnosticsStatus,
            $"Config: {ConfigPath}\nLogs: {LogPath}");
    }

    private void RefreshScheduler_Click(object sender, RoutedEventArgs e)
    {
        RefreshSchedulerStatus();
        ShowFeedback(
            "Scheduler refreshed",
            SchedulerSummary,
            "Task status",
            SchedulerDetails);
    }

    private void Install_Click(object sender, RoutedEventArgs e)
    {
        if (!SaveConfig())
        {
            return;
        }

        var executablePath = ResolveThemeAutoExecutable();
        if (string.IsNullOrWhiteSpace(executablePath))
        {
            SetActiveView(AppView.Scheduler);
            ShowFeedback(
                "Install needs the CLI",
                "themeauto.exe was not found. Publish the CLI first, then try Install again.",
                "Install paused",
                "Expected location: %LOCALAPPDATA%\\WindowsThemeAuto\\themeauto.exe");
            return;
        }

        var result = _schedulerService.Install(executablePath, _config);
        RefreshSchedulerStatus();
        SetActiveView(AppView.Scheduler);
        ShowFeedback(
            result.Success ? "Scheduled tasks installed" : "Install finished with warnings",
            result.Success
                ? "Day, night, and logon tasks are registered."
                : "Windows returned messages while creating the tasks.",
            result.Success ? "Installed" : "Review needed",
            string.Join(Environment.NewLine, result.Messages));
    }

    private void Uninstall_Click(object sender, RoutedEventArgs e)
    {
        var result = _schedulerService.Uninstall();
        RefreshSchedulerStatus();
        SetActiveView(AppView.Scheduler);
        ShowFeedback(
            "Scheduled tasks removed",
            "The app removed the task registrations it manages.",
            result.Success ? "Uninstalled" : "Review needed",
            string.Join(Environment.NewLine, result.Messages));
    }

    private bool SaveConfig()
    {
        if (!TryBuildConfig(out var config, out var error))
        {
            ShowFeedback(
                "Invalid settings",
                error,
                "Check time format",
                "Use 24-hour HH:mm values, for example 07:00 and 19:00.");
            return false;
        }

        _configurationStore.Save(config);
        _config = config;
        RefreshDiagnostics();
        return true;
    }

    private bool TryBuildConfig(out AutomationConfig config, out string error)
    {
        config = _config;
        error = string.Empty;

        if (!TimeOnly.TryParse(DayStartText, out var dayStart))
        {
            error = "Day start time must use HH:mm format.";
            return false;
        }

        if (!TimeOnly.TryParse(NightStartText, out var nightStart))
        {
            error = "Night start time must use HH:mm format.";
            return false;
        }

        if (!int.TryParse(LogonDelaySecondsText, out var logonDelaySeconds))
        {
            error = "Logon delay must be a whole number of seconds.";
            return false;
        }

        config = new AutomationConfig(
            dayStart,
            nightStart,
            new ThemeProfile(ProfileKind.Day, ThemeMode.Light, (int)Math.Round(DayWarmth)),
            new ThemeProfile(ProfileKind.Night, ThemeMode.Dark, (int)Math.Round(NightWarmth)),
            EnableFallbackFilter,
            _config.LogLevel)
        {
            ActiveDays = BuildActiveDays(),
            ApplyOnLogon = ApplyOnLogon,
            LogonDelaySeconds = logonDelaySeconds,
            CatchUpMissedSwitch = CatchUpMissedSwitch
        }.Normalize();

        LogonDelaySecondsText = config.NormalizedLogonDelaySeconds.ToString();
        return true;
    }

    private void RefreshDiagnostics()
    {
        var diagnostics = _nightLightService.Diagnose();
        var serviceIssues = diagnostics.Services.Count(service => !service.IsHealthy);
        DiagnosticsSummary = $"{diagnostics.SettingsBlobCount} settings blobs and {diagnostics.StateBlobCount} state blobs detected, including per-device settings. Services: {(serviceIssues == 0 ? "healthy" : $"{serviceIssues} issue(s)")}.";
        DiagnosticsStatus = diagnostics.SettingsBlobCount > 0
            ? diagnostics.NativeRefreshStatus == ThemeAutomation.Core.NightLight.NativeNightLightRefreshStatus.NativeRegistryOnly
                ? "Refresh stale"
                : "Healthy"
            : "Setup needed";
    }

    private void RefreshSchedulerStatus()
    {
        var schedulerStatus = _schedulerService.GetStatus();
        var installedCount = schedulerStatus.Tasks.Count(task => task.Exists);
        SchedulerSummary = $"{installedCount} of {schedulerStatus.Tasks.Count} managed tasks are installed. Active days: {ActiveDaysSummary}.";
        SchedulerDetails = string.Join(Environment.NewLine, schedulerStatus.Tasks.Select(FormatTaskStatus));
    }

    private void SetActiveView(AppView activeView)
    {
        if (_activeView != activeView)
        {
            _activeView = activeView;
            OnPropertyChanged(nameof(ProfilesViewVisibility));
            OnPropertyChanged(nameof(DiagnosticsViewVisibility));
            OnPropertyChanged(nameof(SchedulerViewVisibility));
        }

        UpdateNavigationStyles();
    }

    private void UpdateNavigationStyles()
    {
        SetNavigationButtonState(ProfilesNavButton, _activeView == AppView.Profiles);
        SetNavigationButtonState(DiagnosticsNavButton, _activeView == AppView.Diagnostics);
        SetNavigationButtonState(SchedulerNavButton, _activeView == AppView.Scheduler);
    }

    private void ShowFeedback(string title, string message, string badge, string details)
    {
        FeedbackTitle = title;
        FeedbackMessage = message;
        FeedbackBadge = badge;
        FeedbackDetails = string.IsNullOrWhiteSpace(details) ? "No extra details." : details;
        AppliedOverlayVisibility = Visibility.Visible;
    }

    private static string? ResolveThemeAutoExecutable()
    {
        var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var publishedCli = Path.Combine(localAppData, "WindowsThemeAuto", "themeauto.exe");
        if (File.Exists(publishedCli))
        {
            return publishedCli;
        }

        var nearbyCli = Path.Combine(AppContext.BaseDirectory, "themeauto.exe");
        return File.Exists(nearbyCli) ? nearbyCli : null;
    }

    private AutomationDays BuildActiveDays()
    {
        var days = AutomationDays.None;
        if (MondayActive)
        {
            days |= AutomationDays.Monday;
        }

        if (TuesdayActive)
        {
            days |= AutomationDays.Tuesday;
        }

        if (WednesdayActive)
        {
            days |= AutomationDays.Wednesday;
        }

        if (ThursdayActive)
        {
            days |= AutomationDays.Thursday;
        }

        if (FridayActive)
        {
            days |= AutomationDays.Friday;
        }

        if (SaturdayActive)
        {
            days |= AutomationDays.Saturday;
        }

        if (SundayActive)
        {
            days |= AutomationDays.Sunday;
        }

        return days;
    }

    private void SetActiveDayFields(AutomationDays activeDays)
    {
        _mondayActive = activeDays.HasFlag(AutomationDays.Monday);
        _tuesdayActive = activeDays.HasFlag(AutomationDays.Tuesday);
        _wednesdayActive = activeDays.HasFlag(AutomationDays.Wednesday);
        _thursdayActive = activeDays.HasFlag(AutomationDays.Thursday);
        _fridayActive = activeDays.HasFlag(AutomationDays.Friday);
        _saturdayActive = activeDays.HasFlag(AutomationDays.Saturday);
        _sundayActive = activeDays.HasFlag(AutomationDays.Sunday);
    }

    private bool SetActiveDayField(ref bool field, bool value, [CallerMemberName] string? propertyName = null)
    {
        if (!SetField(ref field, value, propertyName))
        {
            return false;
        }

        OnPropertyChanged(nameof(ActiveDaysSummary));
        RefreshSchedulerStatus();
        return true;
    }

    private static void AddActiveDayName(List<string> names, AutomationDays activeDays, AutomationDays day, string label)
    {
        if (activeDays.HasFlag(day))
        {
            names.Add(label);
        }
    }

    private static string FormatTaskStatus(ScheduledTaskStatus task)
    {
        if (!task.Exists)
        {
            return $"{task.Name}: missing";
        }

        var nextRun = task.NextRunTime.HasValue ? task.NextRunTime.Value.ToString("g") : "not scheduled";
        var state = string.IsNullOrWhiteSpace(task.State) ? "exists" : task.State;
        return $"{task.Name}: {state}, next run {nextRun}";
    }

    private void SetNavigationButtonState(Button button, bool isActive)
    {
        button.Background = (Brush)FindResource(isActive ? "SidebarNavActiveBrush" : "SidebarNavInactiveBrush");
        button.BorderBrush = (Brush)FindResource(isActive ? "SidebarNavActiveBrush" : "SidebarNavInactiveBrush");
        button.Foreground = (Brush)FindResource(isActive ? "SidebarNavActiveTextBrush" : "SidebarNavMutedTextBrush");
        button.FontWeight = isActive ? FontWeights.SemiBold : FontWeights.Medium;
    }

    private bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return false;
        }

        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

    private enum AppView
    {
        Profiles,
        Diagnostics,
        Scheduler
    }
}
