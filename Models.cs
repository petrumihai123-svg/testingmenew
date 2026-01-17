namespace PortableWinFormsRecorder;

public sealed class Script
{
    public AppTarget App { get; set; } = new();
    public List<Step> Steps { get; set; } = new();
}

public sealed class AppTarget
{
    public string ProcessName { get; set; } = "";
}

public sealed class Step
{
    public string Action { get; set; } = "";
    public Target Target { get; set; } = new();
    public string? Value { get; set; }
    public int? DelayMs { get; set; }
    public string? Note { get; set; }
}

public sealed class Target
{
    public string? AutomationId { get; set; }
    public string? Name { get; set; }
    public string? ClassName { get; set; }
    public string? ControlType { get; set; }
    public string? WindowName { get; set; }

    public string? Image { get; set; }
    public int? ClickOffsetX { get; set; }
    public int? ClickOffsetY { get; set; }
}

public static class JsonOpts
{
    public static readonly System.Text.Json.JsonSerializerOptions Default = new()
    {
        PropertyNameCaseInsensitive = true
    };

    public static readonly System.Text.Json.JsonSerializerOptions Indented = new()
    {
        WriteIndented = true
    };
}