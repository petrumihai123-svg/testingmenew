using FlaUI.Core.AutomationElements;

namespace PortableWinFormsRecorder;

public static class Selectors
{
    public static Target FromElement(AutomationElement el, string? windowNameHint = null)
    {
        var p = el.Properties;

        var t = new Target
        {
            AutomationId = p.AutomationId.ValueOrDefault,
            Name = p.Name.ValueOrDefault,
            ClassName = p.ClassName.ValueOrDefault,
            ControlType = p.ControlType.ValueOrDefault.ToString(),
            WindowName = windowNameHint
        };

        if (string.IsNullOrWhiteSpace(t.AutomationId)) t.AutomationId = null;
        if (string.IsNullOrWhiteSpace(t.Name)) t.Name = null;
        if (string.IsNullOrWhiteSpace(t.ClassName)) t.ClassName = null;
        if (string.IsNullOrWhiteSpace(t.ControlType)) t.ControlType = null;
        if (string.IsNullOrWhiteSpace(t.WindowName)) t.WindowName = null;

        return t;
    }

    public static bool IsTooWeak(Target t)
        => string.IsNullOrWhiteSpace(t.AutomationId) &&
           string.IsNullOrWhiteSpace(t.Name) &&
           string.IsNullOrWhiteSpace(t.ClassName);
}