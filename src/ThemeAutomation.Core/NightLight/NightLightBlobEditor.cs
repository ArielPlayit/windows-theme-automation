namespace ThemeAutomation.Core.NightLight;

public static class NightLightBlobEditor
{
    public static NightLightBlobEditResult TrySetTemperature(byte[] data, int percentage)
    {
        if (data.Length < 4)
        {
            return NightLightBlobEditResult.NotUpdated("Night Light data is too short.");
        }

        var temperature = CalculateTemperatureValue(percentage);
        var lowByte = (byte)(temperature & 0xFF);
        var highByte = (byte)((temperature >> 8) & 0xFF);
        var markersUpdated = 0;
        var temperatureBytes = new bool[data.Length];

        for (var i = 0; i < data.Length - 3; i++)
        {
            if (data[i] != 0xCF || data[i + 1] != 0x28)
            {
                continue;
            }

            data[i + 2] = lowByte;
            data[i + 3] = highByte;
            temperatureBytes[i + 2] = true;
            temperatureBytes[i + 3] = true;
            markersUpdated++;
        }

        return markersUpdated == 0
            ? NightLightBlobEditResult.NotUpdated("No CF 28 temperature marker was found.")
            : NightLightBlobEditResult.UpdatedMarkers(markersUpdated, temperature, TryBumpCloudStoreTimestamp(data, temperatureBytes));
    }

    public static int CalculateTemperatureValue(int percentage)
    {
        percentage = Math.Clamp(percentage, 0, 100);

        if (percentage <= 20)
        {
            return (int)Math.Round(25600 - (percentage / 20.0 * (25600 - 21888)));
        }

        if (percentage <= 50)
        {
            return (int)Math.Round(21888 - ((percentage - 20) / 30.0 * (21888 - 15274)));
        }

        return (int)Math.Round(15274 - ((percentage - 50) / 50.0 * (15274 - 2560)));
    }

    private static bool TryBumpCloudStoreTimestamp(byte[] data, bool[] protectedOffsets)
    {
        if (data.Length <= 14)
        {
            return false;
        }

        for (var index = 10; index <= 14; index++)
        {
            if (protectedOffsets[index] || data[index] == 0xFF)
            {
                continue;
            }

            data[index]++;
            return true;
        }

        return false;
    }
}

public sealed record NightLightBlobEditResult(
    bool Updated,
    int MarkersUpdated,
    int TemperatureValue,
    string Message,
    bool TimestampBumped)
{
    public static NightLightBlobEditResult UpdatedMarkers(int markersUpdated, int temperatureValue, bool timestampBumped) =>
        new(
            true,
            markersUpdated,
            temperatureValue,
            timestampBumped
                ? $"Updated {markersUpdated} marker(s) and bumped CloudStore timestamp."
                : $"Updated {markersUpdated} marker(s), but CloudStore timestamp could not be bumped.",
            timestampBumped);

    public static NightLightBlobEditResult NotUpdated(string message) =>
        new(false, 0, 0, message, false);
}
