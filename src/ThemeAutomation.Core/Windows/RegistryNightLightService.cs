using Microsoft.Win32;
using ThemeAutomation.Core.NightLight;
using ThemeAutomation.Core.Services;

namespace ThemeAutomation.Core.Windows;

public sealed class RegistryNightLightService : INightLightService
{
    private const string CurrentStore = @"Software\Microsoft\Windows\CurrentVersion\CloudStore\Store\DefaultAccount\Current";
    private readonly IFilterFallbackService _fallbackService;

    public RegistryNightLightService(IFilterFallbackService fallbackService)
    {
        _fallbackService = fallbackService;
    }

    public NightLightServiceResult Apply(int percentage, bool fallbackEnabled)
    {
        var messages = new List<string>();
        var nudgePercentage = GetNudgePercentage(percentage);
        var nudgeUpdated = UpdateSettingsBlobs(nudgePercentage, messages, "refresh nudge");
        if (nudgeUpdated.UpdatedAny)
        {
            WindowsSettingsNotifier.BroadcastNightLightRefresh();
            Thread.Sleep(350);
        }

        var settingsUpdated = UpdateSettingsBlobs(percentage, messages, "target warmth");
        var stateUpdated = EnableStateBlobs(messages);
        var registryVerified = settingsUpdated.UpdatedAny && VerifySettingsBlobs(percentage, messages);
        var nativeVisuallyRefreshed = registryVerified && settingsUpdated.TimestampBumped;

        if (settingsUpdated.UpdatedAny || stateUpdated)
        {
            WindowsSettingsNotifier.BroadcastNightLightRefresh();
        }

        if (registryVerified && !settingsUpdated.TimestampBumped)
        {
            messages.Add("Night Light registry value verified, but CloudStore timestamp did not change; Windows may keep the previous visual filter until Settings refreshes it.");
        }

        var decision = FallbackDecision.Decide(registryVerified, nativeVisuallyRefreshed, fallbackEnabled);
        var fallbackApplied = false;
        if (decision.ApplyFallback)
        {
            fallbackApplied = _fallbackService.Apply(percentage);
            messages.Add(fallbackApplied
                ? "Applied gamma fallback filter."
                : "Native Night Light failed and gamma fallback could not be applied.");
        }

        if (decision.Status == NightLightApplyStatus.NativeRegistryOnly)
        {
            messages.Add("Native Night Light registry update succeeded, but visual refresh could not be confirmed.");
        }

        if (!registryVerified && !decision.ApplyFallback)
        {
            messages.Add("Native Night Light did not verify and fallback is disabled.");
        }

        return new NightLightServiceResult(
            decision.Status,
            registryVerified,
            fallbackApplied,
            messages,
            decision.NativeRefreshStatus);
    }

    public NightLightDiagnostics Diagnose()
    {
        var settings = FindDataKeys("bluelightreduction.settings").ToList();
        var states = FindDataKeys("bluelightreductionstate").ToList();
        return new NightLightDiagnostics(
            settings.Any(path => path.Contains("default$", StringComparison.OrdinalIgnoreCase)),
            settings.Count,
            states.Count,
            settings.Concat(states).ToList(),
            Environment.OSVersion.VersionString,
            NativeNightLightRefreshStatus.Unknown,
            GetDependencyServiceStatuses());
    }

    private static SettingsUpdateResult UpdateSettingsBlobs(int percentage, List<string> messages, string phase)
    {
        var updatedAny = false;
        var timestampBumped = false;
        var markersUpdated = 0;
        var paths = FindDataKeys("bluelightreduction.settings").ToList();
        if (paths.Count == 0)
        {
            messages.Add("Night Light settings are not initialized. Enable Night Light once in Windows Settings.");
            return new SettingsUpdateResult(false, false, 0);
        }

        foreach (var path in paths)
        {
            using var key = Registry.CurrentUser.OpenSubKey(path, writable: true);
            var data = key?.GetValue("Data") as byte[];
            if (key is null || data is null)
            {
                messages.Add($"Could not read Data from {path}.");
                continue;
            }

            var edit = NightLightBlobEditor.TrySetTemperature(data, percentage);
            if (!edit.Updated)
            {
                messages.Add($"{path}: {edit.Message}");
                continue;
            }

            key.SetValue("Data", data, RegistryValueKind.Binary);
            updatedAny = true;
            timestampBumped |= edit.TimestampBumped;
            markersUpdated += edit.MarkersUpdated;
            messages.Add($"{path} ({phase}): {edit.Message}");
        }

        return new SettingsUpdateResult(updatedAny, timestampBumped, markersUpdated);
    }

    private static bool EnableStateBlobs(List<string> messages)
    {
        var updatedAny = false;
        foreach (var path in FindDataKeys("bluelightreductionstate"))
        {
            using var key = Registry.CurrentUser.OpenSubKey(path, writable: true);
            var data = key?.GetValue("Data") as byte[];
            if (key is null || data is null || data.Length < 6)
            {
                continue;
            }

            var enabledOffset = data.Length - 5;
            if (data[enabledOffset] == 1)
            {
                continue;
            }

            data[enabledOffset] = 1;
            key.SetValue("Data", data, RegistryValueKind.Binary);
            updatedAny = true;
            messages.Add($"{path}: enabled Night Light state byte.");
        }

        return updatedAny;
    }

    private static bool VerifySettingsBlobs(int percentage, List<string> messages)
    {
        var expected = NightLightBlobEditor.CalculateTemperatureValue(percentage);
        var expectedLow = (byte)(expected & 0xFF);
        var expectedHigh = (byte)((expected >> 8) & 0xFF);
        var verifiedAny = false;

        foreach (var path in FindDataKeys("bluelightreduction.settings"))
        {
            using var key = Registry.CurrentUser.OpenSubKey(path, writable: false);
            var data = key?.GetValue("Data") as byte[];
            if (data is null)
            {
                continue;
            }

            for (var i = 0; i < data.Length - 3; i++)
            {
                if (data[i] == 0xCF &&
                    data[i + 1] == 0x28 &&
                    data[i + 2] == expectedLow &&
                    data[i + 3] == expectedHigh)
                {
                    verifiedAny = true;
                    break;
                }
            }
        }

        messages.Add(verifiedAny
            ? "Verified Night Light registry temperature value."
            : "Could not verify the Night Light registry temperature value.");

        return verifiedAny;
    }

    private static IEnumerable<string> FindDataKeys(string nameFragment)
    {
        using var current = Registry.CurrentUser.OpenSubKey(CurrentStore, writable: false);
        if (current is null)
        {
            yield break;
        }

        foreach (var parentName in current.GetSubKeyNames())
        {
            if (!parentName.Contains(nameFragment, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            using var parent = current.OpenSubKey(parentName, writable: false);
            if (parent is null)
            {
                continue;
            }

            foreach (var childName in parent.GetSubKeyNames())
            {
                using var child = parent.OpenSubKey(childName, writable: false);
                if (child?.GetValue("Data") is byte[])
                {
                    yield return $@"{CurrentStore}\{parentName}\{childName}";
                }
            }
        }
    }

    private static int GetNudgePercentage(int percentage)
    {
        percentage = Math.Clamp(percentage, 0, 100);
        return percentage < 100 ? percentage + 1 : percentage - 1;
    }

    private static IReadOnlyList<NightLightDependencyServiceStatus> GetDependencyServiceStatuses()
    {
        const string servicesKeyPath = @"SYSTEM\CurrentControlSet\Services";
        using var servicesKey = Registry.LocalMachine.OpenSubKey(servicesKeyPath, writable: false);
        if (servicesKey is null)
        {
            return
            [
                new NightLightDependencyServiceStatus(
                    "Windows services registry",
                    "Unavailable",
                    false,
                    "Could not open HKLM service configuration.")
            ];
        }

        var names = servicesKey.GetSubKeyNames();
        var statuses = new List<NightLightDependencyServiceStatus>
        {
            ReadServiceStatus(servicesKey, "CDPSvc"),
            ReadServiceStatus(servicesKey, "NcbService")
        };

        var cdpUserService = names.FirstOrDefault(name => name.StartsWith("CDPUserSvc", StringComparison.OrdinalIgnoreCase));
        statuses.Add(cdpUserService is null
            ? new NightLightDependencyServiceStatus(
                "CDPUserSvc*",
                "Missing",
                false,
                "Connected Devices Platform User Service was not found.")
            : ReadServiceStatus(servicesKey, cdpUserService));

        return statuses;
    }

    private static NightLightDependencyServiceStatus ReadServiceStatus(RegistryKey servicesKey, string name)
    {
        using var key = servicesKey.OpenSubKey(name, writable: false);
        if (key is null)
        {
            return new NightLightDependencyServiceStatus(name, "Missing", false, "Service was not found.");
        }

        var startValue = key.GetValue("Start");
        var isDisabled = startValue is int start && start == 4;
        var state = isDisabled ? "Disabled" : "Configured";
        var message = isDisabled
            ? "Service is disabled; Windows Night Light may not refresh correctly."
            : "Service is present and not disabled.";

        return new NightLightDependencyServiceStatus(name, state, !isDisabled, message);
    }

    private sealed record SettingsUpdateResult(bool UpdatedAny, bool TimestampBumped, int MarkersUpdated);
}
