namespace ThemeAutomation.Core.NightLight;

public static class FallbackDecision
{
    public static FallbackDecisionResult Decide(bool nativeRegistryUpdated, bool nativeVisuallyRefreshed, bool fallbackEnabled)
    {
        if (nativeRegistryUpdated && nativeVisuallyRefreshed)
        {
            return new FallbackDecisionResult(
                false,
                NightLightApplyStatus.NativeApplied,
                NativeNightLightRefreshStatus.NativeApplied);
        }

        if (nativeRegistryUpdated)
        {
            return fallbackEnabled
                ? new FallbackDecisionResult(
                    true,
                    NightLightApplyStatus.FallbackApplied,
                    NativeNightLightRefreshStatus.NativeRegistryOnly)
                : new FallbackDecisionResult(
                    false,
                    NightLightApplyStatus.NativeRegistryOnly,
                    NativeNightLightRefreshStatus.NativeRegistryOnly);
        }

        return fallbackEnabled
            ? new FallbackDecisionResult(
                true,
                NightLightApplyStatus.FallbackApplied,
                NativeNightLightRefreshStatus.NativeFailed)
            : new FallbackDecisionResult(
                false,
                NightLightApplyStatus.Degraded,
                NativeNightLightRefreshStatus.NativeFailed);
    }
}

public sealed record FallbackDecisionResult(
    bool ApplyFallback,
    NightLightApplyStatus Status,
    NativeNightLightRefreshStatus NativeRefreshStatus);

public enum NightLightApplyStatus
{
    Skipped,
    NativeApplied,
    NativeRegistryOnly,
    FallbackApplied,
    Degraded
}

public enum NativeNightLightRefreshStatus
{
    Unknown,
    NativeApplied,
    NativeRegistryOnly,
    NativeFailed
}
