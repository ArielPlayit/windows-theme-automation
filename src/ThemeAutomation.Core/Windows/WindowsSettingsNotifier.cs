using System.Runtime.InteropServices;

namespace ThemeAutomation.Core.Windows;

internal static partial class WindowsSettingsNotifier
{
    private const uint HwndBroadcast = 0xFFFF;
    private const uint WmSettingChange = 0x001A;
    private const uint SmtoAbortIfHung = 0x0002;

    public static void Broadcast(string? area)
    {
        _ = SendMessageTimeout(
            new IntPtr(HwndBroadcast),
            WmSettingChange,
            IntPtr.Zero,
            area,
            SmtoAbortIfHung,
            5000,
            out _);
    }

    public static void BroadcastNightLightRefresh()
    {
        Broadcast(null);
        Broadcast("Windows.UI.ColorEffects");
        Broadcast("Display");
        Broadcast("ImmersiveColorSet");
        Broadcast("Windows.Data.BlueLightReduction");
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SendMessageTimeout(
        IntPtr hWnd,
        uint msg,
        IntPtr wParam,
        string? lParam,
        uint flags,
        uint timeout,
        out IntPtr result);
}
