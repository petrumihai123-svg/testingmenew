namespace PortableWinFormsRecorder;

public sealed class Args
{
    private readonly Dictionary<string, string?> _map = new(StringComparer.OrdinalIgnoreCase);

    public static Args Parse(string[] args)
    {
        var a = new Args();
        for (int i = 0; i < args.Length; i++)
        {
            var k = args[i];
            if (!k.StartsWith("--")) continue;

            string? v = null;
            if (i + 1 < args.Length && !args[i + 1].StartsWith("--"))
            {
                v = args[i + 1];
                i++;
            }
            a._map[k] = v;
        }
        return a;
    }

    public bool Has(string key) => _map.ContainsKey(key);
    public string? Get(string key) => _map.TryGetValue(key, out var v) ? v : null;
    public string Required(string key) => Get(key) ?? throw new Exception($"Missing required arg: {key}");
}