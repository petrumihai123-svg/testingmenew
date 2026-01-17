using System.Text;

namespace PortableWinFormsRecorder;

public static class ScriptCsv
{
    // Exports the recorded script steps to a simple CSV for easy viewing/editing.
    // Columns:
    // Action, AutomationId, Name, ClassName, ControlType, Image, Value, DelayMs, Note
    public static void Save(Script script, string path)
    {
        var sb = new StringBuilder();
        sb.AppendLine("Action,AutomationId,Name,ClassName,ControlType,Image,Value,DelayMs,Note");

        foreach (var s in script.Steps)
        {
            sb.Append(Escape(s.Action)).Append(',')
              .Append(Escape(s.Target.AutomationId)).Append(',')
              .Append(Escape(s.Target.Name)).Append(',')
              .Append(Escape(s.Target.ClassName)).Append(',')
              .Append(Escape(s.Target.ControlType)).Append(',')
              .Append(Escape(s.Target.Image)).Append(',')
              .Append(Escape(s.Value)).Append(',')
              .Append(Escape(s.DelayMs?.ToString())).Append(',')
              .Append(Escape(s.Note))
              .AppendLine();
        }

        File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
    }


    public static Script Load(string csvPath, string processName = "")
    {
        var rows = Csv.Load(csvPath);
        var script = new Script
        {
            App = new AppTarget { ProcessName = processName },
            Steps = new List<Step>()
        };

        foreach (var r in rows)
        {
            var step = new Step
            {
                Action = Get(r, "Action"),
                Value = NullIfEmpty(Get(r, "Value")),
                Note = NullIfEmpty(Get(r, "Note")),
                DelayMs = int.TryParse(Get(r, "DelayMs"), out var d) ? d : null,
                Target = new Target
                {
                    AutomationId = NullIfEmpty(Get(r, "AutomationId")),
                    Name = NullIfEmpty(Get(r, "Name")),
                    ClassName = NullIfEmpty(Get(r, "ClassName")),
                    ControlType = NullIfEmpty(Get(r, "ControlType")),
                    Image = NullIfEmpty(Get(r, "Image"))
                }
            };
            script.Steps.Add(step);
        }

        return script;
    }

    private static string Get(Dictionary<string, string> row, string key)
        => row.TryGetValue(key, out var v) ? v : "";

    private static string? NullIfEmpty(string? v)
        => string.IsNullOrWhiteSpace(v) ? null : v;

    private static string Escape(string? v)
    {
        v ??= "";
        v = v.Replace("\"", "\"\"");
        return $"\"{v}\"";
    }
}