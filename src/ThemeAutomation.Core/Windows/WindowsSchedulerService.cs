using System.Diagnostics;
using System.Globalization;
using System.Text;
using ThemeAutomation.Core.Configuration;
using ThemeAutomation.Core.Services;

namespace ThemeAutomation.Core.Windows;

public sealed class WindowsSchedulerService : ISchedulerService
{
    public const string DayTaskName = "ThemeAutoSwitch_7AM";
    public const string NightTaskName = "ThemeAutoSwitch_7PM";
    public const string StartupTaskName = "ThemeAutoSwitch_Startup";

    private static readonly string[] ManagedTasks =
    [
        DayTaskName,
        NightTaskName,
        StartupTaskName,
        "ThemeAutoSwitch_Hourly",
        "ThemeAutoSwitch_Daily"
    ];

    public SchedulerResult Install(string executablePath, AutomationConfig config)
    {
        var messages = new List<string>();
        config = config.Normalize();
        foreach (var taskName in ManagedTasks)
        {
            RunSchtasks($"/Delete /TN \"{taskName}\" /F", messages, allowFailure: true);
        }

        var success = true;
        foreach (var arguments in BuildInstallArguments(executablePath, config))
        {
            success &= RunSchtasks(arguments, messages);
        }

        return new SchedulerResult(success, messages);
    }

    public static IReadOnlyList<string> BuildInstallArguments(string executablePath, AutomationConfig config)
    {
        config = config.Normalize();
        var dayTime = config.DayStart.ToString("HH:mm", CultureInfo.InvariantCulture);
        var nightTime = config.NightStart.ToString("HH:mm", CultureInfo.InvariantCulture);
        var scheduledAction = Quote(executablePath) + " apply --trigger automatic";
        var logonAction = Quote(executablePath) + " apply --trigger logon";

        var arguments = new List<string>
        {
            $"/Create /TN \"{DayTaskName}\" /SC DAILY /ST {dayTime} /TR \"{scheduledAction}\" /F",
            $"/Create /TN \"{NightTaskName}\" /SC DAILY /ST {nightTime} /TR \"{scheduledAction}\" /F"
        };

        if (config.ApplyOnLogon)
        {
            arguments.Add($"/Create /TN \"{StartupTaskName}\" /SC ONLOGON /DELAY {FormatDelay(config.NormalizedLogonDelaySeconds)} /TR \"{logonAction}\" /F");
        }

        return arguments;
    }

    public SchedulerResult Uninstall()
    {
        var messages = new List<string>();
        foreach (var taskName in ManagedTasks)
        {
            RunSchtasks($"/Delete /TN \"{taskName}\" /F", messages, allowFailure: true);
        }

        return new SchedulerResult(true, messages);
    }

    public SchedulerStatus GetStatus()
    {
        var tasks = new List<ScheduledTaskStatus>();
        foreach (var taskName in ManagedTasks.Take(3))
        {
            var output = CaptureSchtasks($"/Query /TN \"{taskName}\" /V /FO LIST");
            if (output.ExitCode != 0)
            {
                tasks.Add(new ScheduledTaskStatus(taskName, false, null, null, null, null));
                continue;
            }

            tasks.Add(new ScheduledTaskStatus(
                taskName,
                true,
                ExtractValue(output.Output, "Status"),
                TryParseDate(ExtractValue(output.Output, "Last Run Time")),
                TryParseDate(ExtractValue(output.Output, "Next Run Time")),
                ExtractValue(output.Output, "Task To Run")));
        }

        return new SchedulerStatus(tasks);
    }

    private static bool RunSchtasks(string arguments, List<string> messages, bool allowFailure = false)
    {
        var result = CaptureSchtasks(arguments);
        var message = string.IsNullOrWhiteSpace(result.Output) ? result.Error : result.Output;
        if (!string.IsNullOrWhiteSpace(message))
        {
            messages.Add(message.Trim());
        }

        return result.ExitCode == 0 || allowFailure;
    }

    private static (int ExitCode, string Output, string Error) CaptureSchtasks(string arguments)
    {
        using var process = Process.Start(new ProcessStartInfo
        {
            FileName = "schtasks.exe",
            Arguments = arguments,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8
        });

        if (process is null)
        {
            return (1, string.Empty, "Could not start schtasks.exe.");
        }

        var output = process.StandardOutput.ReadToEnd();
        var error = process.StandardError.ReadToEnd();
        process.WaitForExit();
        return (process.ExitCode, output, error);
    }

    private static string Quote(string value) => $"\"{value}\"";

    private static string FormatDelay(int seconds)
    {
        var delay = TimeSpan.FromSeconds(Math.Clamp(
            seconds,
            AutomationConfig.MinLogonDelaySeconds,
            AutomationConfig.MaxLogonDelaySeconds));
        return $"{(int)delay.TotalMinutes:0000}:{delay.Seconds:00}";
    }

    private static string? ExtractValue(string text, string key)
    {
        foreach (var line in text.Split(["\r\n", "\n"], StringSplitOptions.None))
        {
            var separator = line.IndexOf(':');
            if (separator < 0)
            {
                continue;
            }

            if (line[..separator].Trim().Equals(key, StringComparison.OrdinalIgnoreCase))
            {
                return line[(separator + 1)..].Trim();
            }
        }

        return null;
    }

    private static DateTime? TryParseDate(string? value)
    {
        if (string.IsNullOrWhiteSpace(value) ||
            value.Equals("N/A", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        return DateTime.TryParse(value, CultureInfo.CurrentCulture, DateTimeStyles.AssumeLocal, out var parsed)
            ? parsed
            : null;
    }
}
