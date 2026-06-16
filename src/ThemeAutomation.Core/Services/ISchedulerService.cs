using ThemeAutomation.Core.Configuration;

namespace ThemeAutomation.Core.Services;

public interface ISchedulerService
{
    SchedulerResult Install(string executablePath, AutomationConfig config);

    SchedulerResult Uninstall();

    SchedulerStatus GetStatus();
}

public sealed record SchedulerResult(bool Success, IReadOnlyList<string> Messages);

public sealed record SchedulerStatus(IReadOnlyList<ScheduledTaskStatus> Tasks);

public sealed record ScheduledTaskStatus(string Name, bool Exists, string? State, DateTime? LastRunTime, DateTime? NextRunTime, string? Action);
