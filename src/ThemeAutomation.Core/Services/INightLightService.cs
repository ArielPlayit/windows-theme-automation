using ThemeAutomation.Core.NightLight;

namespace ThemeAutomation.Core.Services;

public interface INightLightService
{
    NightLightServiceResult Apply(int percentage, bool fallbackEnabled);

    NightLightDiagnostics Diagnose();
}

public sealed record NightLightServiceResult(
    NightLightApplyStatus Status,
    bool NativeUpdated,
    bool FallbackApplied,
    IReadOnlyList<string> Messages,
    NativeNightLightRefreshStatus NativeRefreshStatus = NativeNightLightRefreshStatus.Unknown);

public sealed record NightLightDiagnostics(
    bool HasDefaultSettings,
    int SettingsBlobCount,
    int StateBlobCount,
    IReadOnlyList<string> RegistryPaths,
    string WindowsVersion = "Unknown",
    NativeNightLightRefreshStatus NativeRefreshStatus = NativeNightLightRefreshStatus.Unknown,
    IReadOnlyList<NightLightDependencyServiceStatus>? DependencyServices = null)
{
    public IReadOnlyList<NightLightDependencyServiceStatus> Services => DependencyServices ?? [];
}

public sealed record NightLightDependencyServiceStatus(
    string Name,
    string State,
    bool IsHealthy,
    string Message);
