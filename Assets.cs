namespace PortableWinFormsRecorder;

public static class Assets
{
    public static string ResolveAssetsDir(string scriptPath)
    {
        var full = Path.GetFullPath(scriptPath);
        var dir = Path.GetDirectoryName(full) ?? ".";
        var name = Path.GetFileNameWithoutExtension(full);
        return Path.Combine(dir, name + "_assets");
    }

    public static string Combine(string assetsDir, string? relative)
    {
        if (string.IsNullOrWhiteSpace(relative)) return "";
        return Path.Combine(assetsDir, relative);
    }
}