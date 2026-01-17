namespace PortableWinFormsRecorder;

public static class Cli
{
    public static int Dispatch(string[] args)
    {
        if (args.Length == 0 || args[0] is "-h" or "--help")
        {
            PrintHelp();
            return 0;
        }

        var cmd = args[0].ToLowerInvariant();
        return cmd switch
        {
            "record" => Record(args.Skip(1).ToArray()),
            "run"    => Run(args.Skip(1).ToArray()),
            "print"  => Print(args.Skip(1).ToArray()),
            _        => Unknown(cmd)
        };
    }

    private static int Unknown(string cmd)
    {
        Console.Error.WriteLine($"Unknown command: {cmd}");
        PrintHelp();
        return 2;
    }

    public static void PrintHelp()
    {
        Console.WriteLine("""
PortableWinFormsRecorder v1 - GUI + CLI recorder/runner for WinForms (UIA + image fallback)

USAGE:
  PortableWinFormsRecorder record [--process MyApp] --out script.json [--var-mode] [--capture-images]
  PortableWinFormsRecorder run --script script.json --data data.csv
  PortableWinFormsRecorder print --script script.json

RECORDING:
  - Desktop mode works without a process. If you set --process, start that app first.
  - Stop recording with Ctrl+Shift+S (or in GUI).
  - --capture-images captures element screenshots for image fallback.

ASSERTIONS (in script.json):
  - AssertExists
  - AssertTextEquals
  - AssertTextContains

PUBLISH PORTABLE:
  dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true /p:IncludeNativeLibrariesForSelfExtract=true
""");
    }

    public static int Record(string[] args)
    {
        var opts = Args.Parse(args);
        var processName = opts.Get("--process") ?? "";
        var outPath = opts.Get("--out") ?? "script.json";
        var varMode = opts.Has("--var-mode");
        var captureImages = opts.Has("--capture-images");

        var script = new Script
        {
            App = new AppTarget { ProcessName = processName },
            Steps = new List<Step>()
        };

        var assetsDir = Assets.ResolveAssetsDir(outPath);
        Directory.CreateDirectory(assetsDir);

        using var recorder = new Recorder(processName, script, new RecorderOptions
        {
            VarMode = varMode,
            CaptureImages = captureImages,
            AssetsDir = assetsDir
        });

        recorder.Log += Console.WriteLine;
        recorder.Start();

        Console.WriteLine("Recording... Stop with Ctrl+Shift+S");
        recorder.WaitUntilStopped();

        File.WriteAllText(outPath, System.Text.Json.JsonSerializer.Serialize(script, JsonOpts.Indented));
        Console.WriteLine($"Saved: {Path.GetFullPath(outPath)}");
        return 0;
    }

    public static int Run(string[] args)
    {
        var opts = Args.Parse(args);
        var scriptPath = opts.Get("--script") ?? "script.json";
        var dataPath = opts.Get("--data") ?? "data.csv";

        var script = System.Text.Json.JsonSerializer.Deserialize<Script>(File.ReadAllText(scriptPath), JsonOpts.Default)
                     ?? throw new Exception("Failed to parse script");

        var rows = Csv.Load(dataPath);
        if (rows.Count == 0) throw new Exception("No data rows found in CSV");

        var assetsDir = Assets.ResolveAssetsDir(scriptPath);

        int runNo = 0;
        foreach (var row in rows)
        {
            runNo++;
            Console.WriteLine($"Run #{runNo}: " + string.Join(", ", row.Select(kv => $"{kv.Key}={kv.Value}")));
            Runner.RunOnce(script, row, new RunnerOptions { AssetsDir = assetsDir, Log = Console.WriteLine });
        }

        Console.WriteLine("Done.");
        return 0;
    }

    public static int Print(string[] args)
    {
        var opts = Args.Parse(args);
        var scriptPath = opts.Get("--script") ?? "script.json";
        var script = System.Text.Json.JsonSerializer.Deserialize<Script>(File.ReadAllText(scriptPath), JsonOpts.Default)
                     ?? throw new Exception("Failed to parse script");

        Console.WriteLine(System.Text.Json.JsonSerializer.Serialize(script, JsonOpts.Indented));
        return 0;
    }
}