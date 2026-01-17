namespace PortableWinFormsRecorder;

public static class Csv
{
    public static List<Dictionary<string, string>> Load(string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException("CSV not found", path);

        var lines = File.ReadAllLines(path).Where(l => !string.IsNullOrWhiteSpace(l)).ToArray();
        if (lines.Length < 2) return new();

        var headers = Split(lines[0]).Select(h => h.Trim()).ToArray();
        var rows = new List<Dictionary<string, string>>();

        for (int i = 1; i < lines.Length; i++)
        {
            var cols = Split(lines[i]);
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (int c = 0; c < headers.Length; c++)
                dict[headers[c]] = c < cols.Count ? cols[c] : "";
            rows.Add(dict);
        }

        return rows;
    }

    private static List<string> Split(string line)
    {
        var res = new List<string>();
        var cur = new System.Text.StringBuilder();
        bool inQuotes = false;

        for (int i = 0; i < line.Length; i++)
        {
            char ch = line[i];

            if (ch == '"')
            {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                {
                    cur.Append('"');
                    i++;
                }
                else inQuotes = !inQuotes;
                continue;
            }

            if (ch == ',' && !inQuotes)
            {
                res.Add(cur.ToString());
                cur.Clear();
                continue;
            }

            cur.Append(ch);
        }
        res.Add(cur.ToString());
        return res;
    }
}