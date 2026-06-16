namespace ThemeAutomation.Core.Diagnostics;

public sealed class FileLogger
{
    public FileLogger(string? baseDirectory = null)
    {
        BaseDirectory = baseDirectory ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "WindowsThemeAuto");
        LogDirectory = Path.Combine(BaseDirectory, "logs");
    }

    public string BaseDirectory { get; }

    public string LogDirectory { get; }

    public void Info(string message) => Write("INFO", message);

    public void Error(string message) => Write("ERROR", message);

    private void Write(string level, string message)
    {
        Directory.CreateDirectory(LogDirectory);
        var logPath = Path.Combine(LogDirectory, $"{DateTime.Now:yyyy-MM-dd}.log");
        File.AppendAllText(logPath, $"{DateTime.Now:O} [{level}] {message}{Environment.NewLine}");
    }
}
