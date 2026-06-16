namespace ThemeAutomation.Core.Services;

public interface IFilterFallbackService
{
    bool Apply(int percentage);

    bool Reset();
}
