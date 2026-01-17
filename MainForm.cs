using System.Threading;
using System.Text.Json;
using System.IO;

namespace PortableWinFormsRecorder;

public sealed class MainForm : Form
{
    private static string GetDefaultFlowDir()
    {
        // Prefer a local "win-x64" folder (publish output) if present, otherwise use app startup folder.
        // This makes it easy to keep flows next to the portable EXE after publishing.
        var baseDir = Application.StartupPath;

        // Common publish layouts
        var candidates = new[]
        {
            Path.Combine(baseDir, "win-x64"),
            Path.Combine(baseDir, "publish", "win-x64"),
            Path.Combine(baseDir, "..", "publish", "win-x64"),
            Path.Combine(baseDir, "..", "..", "publish", "win-x64")
        };

        foreach (var c in candidates)
        {
            try
            {
                var full = Path.GetFullPath(c);
                if (Directory.Exists(full))
                    return full;
            }
            catch { /* ignore */ }
        }

        return baseDir;
    }


    private readonly System.Windows.Forms.NumericUpDown _numRetries = new() { Minimum = 0, Maximum = 10, Value = 0, Width = 55 };
    private readonly System.Windows.Forms.NumericUpDown _numRetryDelay = new() { Minimum = 0, Maximum = 10000, Value = 250, Increment = 50, Width = 70 };
    private readonly System.Windows.Forms.CheckBox _chkHighlight = new() { Text = "Highlight current step", AutoSize = true, Checked = true };
    private readonly System.Windows.Forms.CheckBox _chkRunAllRows = new() { Text = "Run all CSV rows", AutoSize = true, Checked = true };

    private readonly System.Windows.Forms.TextBox _txtProcess = new() { Width = 220 };
    private readonly System.Windows.Forms.TextBox _txtScript = new() { Width = 360 };
    private readonly System.Windows.Forms.TextBox _txtData = new() { Width = 360 };
    private readonly System.Windows.Forms.TextBox _txtVars = new() { Width = 300 };
    private readonly System.Windows.Forms.Button _btnBrowseScript = new() { Text = "..." };
    private readonly System.Windows.Forms.Button _btnBrowseData = new() { Text = "..." };
    private readonly System.Windows.Forms.Button _btnBrowseVars = new() { Text = "..." };

    private readonly System.Windows.Forms.Button _btnRecord = new() { Text = "Record" };
    private readonly System.Windows.Forms.Button _btnStop = new() { Text = "Stop", Enabled = false };
    private readonly System.Windows.Forms.Button _btnVideoFlow = new() { Text = "Video → Flow", Width = 100 };
        private readonly System.Windows.Forms.Button _btnPause = new() { Text = "Pause", Enabled = false };
private readonly System.Windows.Forms.Button _btnRun = new() { Text = "Run" };

    private readonly System.Windows.Forms.CheckBox _chkCaptureImages = new() { Text = "Capture images (fallback)", Checked = true };
    private readonly System.Windows.Forms.CheckBox _chkInspector = new() { Text = "Inspector overlay", Checked = false };

    private readonly System.Windows.Forms.TextBox _log = new()
    {
        Multiline = true,
        ScrollBars = ScrollBars.Both,
        ReadOnly = true,
        Dock = DockStyle.Fill,
        Font = new System.Drawing.Font("Consolas", 9f)
    };

    private readonly System.Windows.Forms.PropertyGrid _inspectorGrid = new() { Dock = DockStyle.Fill };
    private readonly StepFlowView _flowView = new() { Dock = DockStyle.Fill };
    private readonly InspectorController _inspector = new();

    private readonly FlowEditorView _flowEditor = new() { Dock = DockStyle.Fill };

    private Recorder? _recorder;
    private Script? _recordingScript;
    private Thread? _recordThread;
    private string? _currentAssetsDir;

    public MainForm()
    {
        Text = "PortableWinFormsRecorder v1";
        Width = 1050;
        Height = 680;

        var tabs = new TabControl { Dock = DockStyle.Fill };
        tabs.TabPages.Add(BuildRunTab());
        tabs.TabPages.Add(BuildInspectorTab());
        tabs.TabPages.Add(BuildFlowTab());
        Controls.Add(tabs);

        _btnBrowseScript.Click += (_, __) => BrowseFileSave(_txtScript, "Script JSON (*.json)|*.json|Flow Steps CSV (*.csv)|*.csv|All files (*.*)|*.*");
        _btnBrowseData.Click += (_, __) => BrowseFileOpen(_txtData, "Flow Steps CSV (*.csv)|*.csv|All files (*.*)|*.*");
        _btnBrowseVars.Click += (_, __) => BrowseFileOpen(_txtVars, "Data CSV (*.csv)|*.csv|All files (*.*)|*.*");

        _btnRecord.Click += (_, __) => StartRecording();
        _btnStop.Click += (_, __) => StopRecording();
        _btnPause.Click += (_, __) => TogglePauseRecording();
        _btnRun.Click += (_, __) => RunScript();
        _btnVideoFlow.Click += (_, __) => ImportFromVideo();

        _chkInspector.CheckedChanged += (_, __) => _inspector.Enabled = _chkInspector.Checked;
        _inspector.ElementChanged += OnInspectElement;
        _flowEditor.Log += AppendLog;
        _flowEditor.RunFromSelected += (idx) => { try { RunScript(startIndex: idx); } catch (Exception ex) { AppendLog("[RUN] ERROR: " + ex); } };

        // sensible defaults
        _txtData.Text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "flow.csv");
        _txtVars.Text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "data.csv");
    }

    private TabPage BuildRunTab()
    {
        var page = new TabPage("Record / Run");

        var top = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 140,
            Padding = new Padding(10),
            AutoSize = false,
            WrapContents = true
        };

        top.Controls.Add(new Label { Text = "Process (optional):", AutoSize = true, Padding = new Padding(0, 8, 0, 0) });
        top.Controls.Add(_txtProcess);


        top.Controls.Add(new Label { Text = "Flow CSV:", AutoSize = true, Padding = new Padding(10, 8, 0, 0) });
        top.Controls.Add(_txtData);
        top.Controls.Add(_btnBrowseData);

        top.Controls.Add(new Label { Text = "Data CSV (optional):", AutoSize = true, Padding = new Padding(10, 8, 0, 0) });
        top.Controls.Add(_txtVars);
        top.Controls.Add(_btnBrowseVars);
        top.Controls.Add(_chkRunAllRows);

        top.Controls.Add(_chkCaptureImages);
        top.Controls.Add(_chkInspector);
        top.Controls.Add(_btnRecord);
        top.Controls.Add(_btnStop);
        top.Controls.Add(_btnPause);

        top.Controls.Add(new Label { Text = "Stop hotkey:", AutoSize = true, Padding = new Padding(10, 8, 0, 0) });
        top.Controls.Add(_chkHotkeyCtrl);
        top.Controls.Add(_chkHotkeyShift);
        top.Controls.Add(_chkHotkeyAlt);
        top.Controls.Add(_txtHotkeyKey);

        top.Controls.Add(_btnRun);
        top.Controls.Add(_btnVideoFlow);

        top.Controls.Add(new Label { Text = "Retry:", AutoSize = true, Padding = new Padding(10, 8, 0, 0) });
        top.Controls.Add(_numRetries);
        top.Controls.Add(new Label { Text = "Delay(ms):", AutoSize = true, Padding = new Padding(10, 8, 0, 0) });
        top.Controls.Add(_numRetryDelay);
        top.Controls.Add(_chkHighlight);

        page.Controls.Add(_log);
        page.Controls.Add(top);
        return page;
    }

    private TabPage BuildInspectorTab()
    {
        var page = new TabPage("Inspector");
        var split = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Vertical, SplitterDistance = 650 };
        split.Panel1.Controls.Add(_inspectorGrid);

        var help = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            Dock = DockStyle.Fill,
            Text =
@"Turn on 'Inspector overlay' on the Record/Run tab.
Hover controls to see UIA properties.

Tips:
- Prefer AutomationId + ControlType when available.
- For weak/custom controls, record with 'Capture images' to use image fallback."
        };
        split.Panel2.Controls.Add(help);

        page.Controls.Add(split);
        return page;
    }
    private TabPage BuildFlowTab()
    {
        var page = new TabPage("Flow");
        page.Controls.Add(_flowEditor);
        return page;
    }

    private void OnInspectElement(FlaUI.Core.AutomationElements.AutomationElement? el)
    {
        if (InvokeRequired) { BeginInvoke(() => OnInspectElement(el)); return; }
        if (el == null) { _inspectorGrid.SelectedObject = null; return; }

        try
        {
            var p = el.Properties;
            _inspectorGrid.SelectedObject = new
            {
                Name = p.Name.ValueOrDefault,
                AutomationId = p.AutomationId.ValueOrDefault,
                ClassName = p.ClassName.ValueOrDefault,
                ControlType = p.ControlType.ValueOrDefault.ToString(),
                FrameworkId = p.FrameworkId.ValueOrDefault,
                ProcessId = p.ProcessId.ValueOrDefault,
                IsEnabled = p.IsEnabled.ValueOrDefault,
                BoundingRect = el.BoundingRectangle.ToString()
            };
        }
        catch { }
    }

    private void BrowseFileOpen(TextBox target, string filter)
    {
        using var dlg = new OpenFileDialog { Filter = filter };
        if (dlg.ShowDialog(this) == DialogResult.OK)
            target.Text = dlg.FileName;
    }

    private string? PromptSaveCsv(string? initialDir, string defaultName)
    {
        using var dlg = new SaveFileDialog
        {
            Filter = "CSV data (*.csv)|*.csv|All files (*.*)|*.*",
            FileName = defaultName,
            OverwritePrompt = true
        };

        if (!string.IsNullOrWhiteSpace(initialDir) && Directory.Exists(initialDir))
            dlg.InitialDirectory = initialDir;

        return dlg.ShowDialog(this) == DialogResult.OK ? dlg.FileName : null;
    }


    private void BrowseFileSave(TextBox target, string filter)
    {
        using var dlg = new SaveFileDialog { Filter = filter, FileName = Path.GetFileName(target.Text) };
        if (dlg.ShowDialog(this) == DialogResult.OK)
            target.Text = dlg.FileName;
    }

    private void AppendLog(string s)
    {
        if (InvokeRequired) { BeginInvoke(() => AppendLog(s)); return; }
        _log.AppendText(s + Environment.NewLine);
    }

    private void StartRecording()
    {
        
        try
        {
            // Minimize immediately so the app doesn't interfere with the recording.
            WindowState = FormWindowState.Minimized;
            // Some systems ignore minimize requests while starting hooks; hide as fallback.
            if (WindowState != FormWindowState.Minimized)
                
            // Wait a moment for the Record button mouse-up to complete before hiding the window,
            // otherwise the mouse-up can "click through" onto the underlying app.
            try
            {
                var sw = System.Diagnostics.Stopwatch.StartNew();
                while (Control.MouseButtons != MouseButtons.None && sw.ElapsedMilliseconds < 500)
                    Thread.Sleep(10);
                Thread.Sleep(80);
            }
            catch { /* ignore */ } // WaitForMouseReleaseBeforeHide
Hide();
        }
        catch { }

if (_recorder != null) return;

        try
        {
            _tray.Visible = true;
            _recOverlay?.Close();
            _recOverlay = new RecordingIndicatorOverlay();
            _recOverlay.Show();
        }
        catch { }

        var process = _txtProcess.Text.Trim();

        var scriptPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "recording.json");
_recordingScript = new Script { App = new AppTarget { ProcessName = process }, Steps = new List<Step>() };
        

        var assetsDir = Assets.ResolveAssetsDir(scriptPath);
        Directory.CreateDirectory(assetsDir);
        _currentAssetsDir = assetsDir;

        _recorder = new Recorder(process, _recordingScript, new RecorderOptions
        {
            VarMode = false,
            CaptureImages = _chkCaptureImages.Checked,
            AssetsDir = assetsDir,
            StopHotkeyCtrl = _chkHotkeyCtrl.Checked,
            StopHotkeyShift = _chkHotkeyShift.Checked,
            StopHotkeyAlt = _chkHotkeyAlt.Checked,
            StopHotkeyKey = string.IsNullOrWhiteSpace(_txtHotkeyKey.Text) ? 'S' : _txtHotkeyKey.Text.Trim()[0]
        });

        
        try
        {
            _btnPause.Enabled = true;
            _btnPause.Text = "Pause";
            _recorder.PausedChanged += paused =>
            {
                try
                {
                    BeginInvoke(() =>
                    {
                        _btnPause.Text = paused ? "Resume" : "Pause";
                        _recOverlay?.SetText(paused ? "PAUSED" : "RECORDING");
                    });
                }
                catch { }
            };
        }
        catch { }

_recorder.Log += AppendLog;
        _flowEditor.BeginRecording(assetsDir);
        _flowEditor.SetCsvPath(null);

        _recorder.StepRecorded += (step, imgPath) =>
        {
            try
            {
                BeginInvoke(() => _flowEditor.AppendStep(step));
            }
            catch { }
        };_btnRecord.Enabled = false;
        _btnStop.Enabled = true;
        _btnRun.Enabled = false;

        _recordThread = new Thread(() =>
        {
            try
            {
                _recorder.Start();
                AppendLog("[GUI] Recording... Stop: Ctrl+Shift+S | Pause/Resume: Ctrl+Shift+P");
                _recorder.WaitUntilStopped();

                File.WriteAllText(scriptPath, JsonSerializer.Serialize(_recordingScript, JsonOpts.Indented));
                // Popout save-as CSV on stop (default to the same folder as Flow images)
                string? csvPath = null;
                var assetsDirForFlow = _currentAssetsDir;
                Invoke(() => { csvPath = PromptSaveCsv(assetsDirForFlow, "flow.csv"); });

                if (!string.IsNullOrWhiteSpace(csvPath))
                {
                    ScriptCsv.Save(_recordingScript!, csvPath!);
                    AppendLog($"[GUI] Saved CSV: {csvPath}");
                    try { BeginInvoke(() => _flowEditor.SetCsvPath(csvPath!)); } catch { }

                    // Ensure flow images are in the same folder as the CSV
                    try
                    {
                        var csvFolder = Path.GetDirectoryName(Path.GetFullPath(csvPath!)) ?? "";
                        if (!string.IsNullOrWhiteSpace(assetsDirForFlow) && Directory.Exists(assetsDirForFlow))
                        {
                            var assetsFull = Path.GetFullPath(assetsDirForFlow);
                            if (!string.Equals(csvFolder, assetsFull, StringComparison.OrdinalIgnoreCase))
                            {
                                foreach (var img in Directory.GetFiles(assetsFull, "*.png"))
                                {
                                    var dest = Path.Combine(csvFolder, Path.GetFileName(img));
                                    if (!File.Exists(dest))
                                        File.Copy(img, dest, overwrite: false);
                                }
                                AppendLog($"[GUI] Copied flow images to: {csvFolder}");
                            }
                            else
                            {
                                AppendLog("[GUI] Flow images already in the same folder as CSV.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        AppendLog("[GUI] Flow copy warning: " + ex.Message);
                    }
                }
                else
                {
                    AppendLog("[GUI] CSV save cancelled.");
                }

                AppendLog($"[GUI] Saved script: {scriptPath}");
                AppendLog($"[GUI] Assets: {assetsDir}");
            }
            catch (Exception ex)
            {
                AppendLog("[GUI] ERROR: " + ex);
            }
            finally
            {
                try { _recorder?.Dispose(); } catch { }
                _recorder = null;

                BeginInvoke(() =>
                {
                    // Always hide recording UI when the recorder stops (including stop-hotkey).
                    try { _tray.Visible = false; } catch { }
                    try { _recOverlay?.Close(); _recOverlay = null; } catch { }

                    // Restore window so save dialogs and UI are visible.
                    try
                    {
                        Show();
                        if (WindowState == FormWindowState.Minimized)
                            WindowState = FormWindowState.Normal;
                        Activate();
                    }
                    catch { }

                    _btnRecord.Enabled = true;
                    _btnStop.Enabled = false;
                    _btnPause.Enabled = false;
                    _btnPause.Text = "Pause";
                    _btnRun.Enabled = true;
                    _btnPause.Enabled = false;
                    _btnPause.Text = "Pause";
                });
            }

        });

        _recordThread.IsBackground = true;
        _recordThread.SetApartmentState(ApartmentState.STA);
        _recordThread.Start();
    }

        private void TogglePauseRecording()
    {
        try { _recorder?.TogglePauseFromUi(); } catch { }
    }

    private void StopRecording()
    {
        
        try
        {
            _tray.Visible = false;
            try { _recOverlay?.Close(); _recOverlay = null; } catch { }
            // Restore window so save dialogs and UI are visible.
            Show();
            WindowState = FormWindowState.Normal;
            Activate();
        }
        catch { }

_recorder?.RequestStop();
        AppendLog("[GUI] stop requested");
    }

    private void RunScript(int startIndex = 0)
    {
        var flowCsvPath = _txtData.Text.Trim();
        if (string.IsNullOrWhiteSpace(flowCsvPath) || !File.Exists(flowCsvPath))
        {
            MessageBox.Show(this, "Flow CSV file not found.", "Missing flow", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var varsCsvPath = _txtVars.Text.Trim();
        var procName = _txtProcess.Text.Trim();

        _btnRun.Enabled = false;

        Task.Run(() =>
        {
            try
            {
                // Steps come from Flow CSV; process name is optional (desktop mode if empty)
                var script = ScriptCsv.Load(flowCsvPath, procName);

                // Assets directory: same folder as the flow CSV (so images live next to it)
                var assetsDir = Path.GetDirectoryName(Path.GetFullPath(flowCsvPath)) ?? ".";

                var rows = new List<Dictionary<string, string>>();
                if (!string.IsNullOrWhiteSpace(varsCsvPath) && File.Exists(varsCsvPath))
                {
                    rows = Csv.Load(varsCsvPath);
                    if (!_chkRunAllRows.Checked && rows.Count > 0)
                        rows = new List<Dictionary<string, string>> { rows[0] };
                }

                if (rows.Count == 0)
                    rows.Add(new Dictionary<string, string>()); // no variables; run once

                var opts = new RunnerOptions
                {
                    AssetsDir = assetsDir,
                    StopHotkeyCtrl = _chkHotkeyCtrl.Checked,
                    StopHotkeyShift = _chkHotkeyShift.Checked,
                    StopHotkeyAlt = _chkHotkeyAlt.Checked,
                    StopHotkeyKey = string.IsNullOrWhiteSpace(_txtHotkeyKey.Text) ? 'S' : _txtHotkeyKey.Text.Trim()[0],
                    Log = AppendLog,
                    StartIndex = Math.Max(0, startIndex),
                    RetryCount = (int)_numRetries.Value,
                    RetryDelayMs = (int)_numRetryDelay.Value,
                    OnStepStart = _chkHighlight.Checked ? (i => { try { BeginInvoke(() => _flowEditor.SetRunningIndex(i)); } catch { } }) : null
                };

                for (int r = 0; r < rows.Count; r++)
                {
                    AppendLog(rows.Count > 1 ? $"[RUN] row {r + 1}/{rows.Count}" : "[RUN] starting");
                    Runner.RunOnce(script, rows[r], opts);
                }

                AppendLog("[RUN] done");
            }
            catch (OperationCanceledException)
            {
                AppendLog("[RUN] stopped");
            }
            catch (Exception ex)
            {
                AppendLog("[RUN] ERROR: " + ex);
            }
            finally
            {
                try { BeginInvoke(() => _btnRun.Enabled = true); } catch { }
            }
        });
    }


    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        try { _inspector.Enabled = false; _inspector.Dispose(); } catch { }
        try { _recorder?.RequestStop(); } catch { }
        base.OnFormClosing(e);
    
    }

    private readonly NotifyIcon _tray = new();
    private RecordingIndicatorOverlay? _recOverlay;

    private readonly TextBox _txtHotkeyKey = new() { Text = "S", Width = 30 };
    private readonly CheckBox _chkHotkeyCtrl = new() { Text = "Ctrl", AutoSize = true, Checked = true };
    private readonly CheckBox _chkHotkeyShift = new() { Text = "Shift", AutoSize = true, Checked = true };
    private readonly CheckBox _chkHotkeyAlt = new() { Text = "Alt", AutoSize = true, Checked = false };


    
    private void ImportFromVideo()
    {
        try
        {
            using var wiz = new VideoToFlowWizard();
            if (wiz.ShowDialog(this) != DialogResult.OK) return;

            // Always save a CSV so we can place extracted template images into the correct *_assets folder.
            using var sfd = new SaveFileDialog
            {
                Filter = "CSV files|*.csv|All files|*.*",
                FileName = "video_flow.csv",
                InitialDirectory = GetDefaultFlowDir(),
                Title = "Save draft flow CSV (with extracted images)"
            };
            if (sfd.ShowDialog(this) != DialogResult.OK) return;

            wiz.SaveCsvWithImages(sfd.FileName);

            // Load what we just saved into the flow editor.
            var script = ScriptCsv.Load(sfd.FileName);
            _flowEditor.SetCsvPath(sfd.FileName);
            _flowEditor.ReplaceSteps(script.Steps);

            AppendLog($"[GUI] Video → Flow: saved CSV + images and loaded {script.Steps.Count} steps.");
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Video → Flow", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }


}


    

