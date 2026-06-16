using System.Runtime.InteropServices;
using ThemeAutomation.Core.Services;

namespace ThemeAutomation.Core.Windows;

public sealed class GammaFallbackService : IFilterFallbackService
{
    public bool Apply(int percentage)
    {
        percentage = Math.Clamp(percentage, 0, 100);
        var ramp = GammaRamp.CreateWarmthRamp(percentage);
        return ApplyRamp(ramp);
    }

    public bool Reset()
    {
        var ramp = GammaRamp.CreateNeutralRamp();
        return ApplyRamp(ramp);
    }

    private static bool ApplyRamp(GammaRamp ramp)
    {
        var hdc = NativeMethods.GetDC(IntPtr.Zero);
        if (hdc == IntPtr.Zero)
        {
            return false;
        }

        try
        {
            return NativeMethods.SetDeviceGammaRamp(hdc, ref ramp);
        }
        finally
        {
            _ = NativeMethods.ReleaseDC(IntPtr.Zero, hdc);
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct GammaRamp
    {
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 256)]
        public ushort[] Red;

        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 256)]
        public ushort[] Green;

        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 256)]
        public ushort[] Blue;

        public static GammaRamp CreateNeutralRamp() => CreateWarmthRamp(0);

        public static GammaRamp CreateWarmthRamp(int percentage)
        {
            var ramp = new GammaRamp
            {
                Red = new ushort[256],
                Green = new ushort[256],
                Blue = new ushort[256]
            };

            var blueMultiplier = 1.0 - (percentage / 100.0 * 0.70);
            var greenMultiplier = 1.0 - (percentage / 100.0 * 0.30);

            for (var i = 0; i < 256; i++)
            {
                var value = i * 257;
                ramp.Red[i] = ClampToUshort(value);
                ramp.Green[i] = ClampToUshort(value * greenMultiplier);
                ramp.Blue[i] = ClampToUshort(value * blueMultiplier);
            }

            return ramp;
        }

        private static ushort ClampToUshort(double value) =>
            (ushort)Math.Clamp((int)Math.Round(value), 0, ushort.MaxValue);
    }

    private static class NativeMethods
    {
        [DllImport("user32.dll")]
        public static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("gdi32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetDeviceGammaRamp(IntPtr hDC, ref GammaRamp ramp);
    }
}
