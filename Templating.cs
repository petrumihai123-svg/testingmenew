namespace PortableWinFormsRecorder;

public static class Templating
{
    public static string Apply(string input, Dictionary<string, string> data)
    {
        var output = input;

        foreach (var kv in data)
            output = output.Replace("{{" + kv.Key + "}}", kv.Value, StringComparison.OrdinalIgnoreCase);

        output = output.Replace("{{timestamp}}", DateTime.Now.ToString("yyyyMMdd_HHmmss"));
        return output;
    }
}