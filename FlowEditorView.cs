using System.Drawing;
using System.Linq;

namespace PortableWinFormsRecorder;

public sealed class FlowEditorView : UserControl
{
    private readonly Button _btnImport = new() { Text = "Import CSV + Flow Images" };
    private readonly Button _btnSaveAs = new() { Text = "Save CSV As..." };
    private readonly Button _btnMoveUp = new() { Text = "Move Up" };
    private readonly Button _btnMoveDown = new() { Text = "Move Down" };
    private readonly Button _btnDelete = new() { Text = "Delete Step" };
    private readonly Button _btnRunFromSelected = new() { Text = "Run from selected" };
    private readonly Button _btnPickOcr = new() { Text = "Pick OCR Region" };

    private readonly Label _lblStatus = new() { AutoSize = true };

    private readonly DataGridView _grid = new()
    {
        Dock = DockStyle.Fill,
        AllowUserToAddRows = false,
        AllowUserToDeleteRows = true,
        RowHeadersVisible = false,
        SelectionMode = DataGridViewSelectionMode.FullRowSelect,
        MultiSelect = false,
        AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    };

    private string? _csvPath;
    private string? _imagesDir;
    private Script _script = new();

    private bool _suppressSave;
    private bool _isRecordingMode;
    private bool _thumbsEnabled = true;

    private int _dragRowIndex = -1;
    private readonly System.Windows.Forms.Timer _saveTimer = new() { Interval = 300 };

    public event Action<string>? Log;
    public event Action<int>? RunFromSelected;

    public FlowEditorView()
    {
        Dock = DockStyle.Fill;

        var top = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 38,
            Padding = new Padding(6),
            WrapContents = false
        };

        top.Controls.Add(_btnImport);
        top.Controls.Add(_btnSaveAs);
        top.Controls.Add(_btnMoveUp);
        top.Controls.Add(_btnMoveDown);
        top.Controls.Add(_btnDelete);
        top.Controls.Add(_btnRunFromSelected);
        top.Controls.Add(_btnPickOcr);
        top.Controls.Add(new Label { Text = "   ", AutoSize = true });
        top.Controls.Add(_lblStatus);

        Controls.Add(_grid);
        Controls.Add(top);

        BuildGrid();
        EnableDragDropReorder();

        _btnImport.Click += (_, __) => Import();
        _btnSaveAs.Click += (_, __) => SaveAs();
        _btnMoveUp.Click += (_, __) => MoveSelected(-1);
        _btnMoveDown.Click += (_, __) => MoveSelected(+1);
        _btnDelete.Click += (_, __) => DeleteSelected();
        _btnRunFromSelected.Click += (_, __) =>
        {
            try { RunFromSelected?.Invoke(GetSelectedIndex()); } catch { }
        };

        _btnPickOcr.Click += (_, __) =>
        {
            try { PickOcrRegionForSelected(); } catch (Exception ex) { Log?.Invoke("[FLOW] OCR pick error: " + ex.Message); }
        };

        _grid.CellValueChanged += (_, __) => SaveSoon();
        _grid.RowsRemoved += (_, __) => SaveSoon();
        _grid.UserDeletingRow += (_, __) => { /* save happens in RowsRemoved */ };
        _grid.SelectionChanged += (_, __) => UpdateButtons();
        _grid.DataError += (_, e) => { e.ThrowException = false; };
        _grid.CellEndEdit += (_, e) =>
        {
            try
            {
                if (e.RowIndex < 0) return;
                var colName = _grid.Columns[e.ColumnIndex].Name;
                if (colName == "Image")
                {
                    var row = _grid.Rows[e.RowIndex];
                    var imgFile = row.Cells["Image"].Value?.ToString();
                    row.Cells["Thumb"].Value = LoadThumb(imgFile);
                }
            }
            catch { }
        };
        _grid.CurrentCellDirtyStateChanged += (_, __) =>
        {
            try
            {
                if (_grid.IsCurrentCellDirty) _grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
            catch { }
        };

        _saveTimer.Tick += (_, __) =>
        {
            _saveTimer.Stop();
            try { SaveIfReady(); } catch (Exception ex) { Log?.Invoke("[FLOW] Debounced save error: " + ex); }
        };

        UpdateButtons();
        UpdateStatus();
    }

    
    public void LoadFromRecording(Script script, string assetsDir)
    {
        _isRecordingMode = false;
        _thumbsEnabled = true;
        _imagesDir = assetsDir;
        _script = new Script { App = script.App, Steps = new List<Step>() };

        foreach (var st in script.Steps.ToList())
        {
            _script.Steps.Add(new Step
            {
                Action = st.Action,
                Value = st.Value,
                Note = st.Note,
                DelayMs = st.DelayMs,
                Target = new Target
                {
                    AutomationId = st.Target.AutomationId,
                    Name = st.Target.Name,
                    ClassName = st.Target.ClassName,
                    ControlType = st.Target.ControlType,
                    Image = st.Target.Image
                }
            });
        }

        RebuildRows();
        UpdateStatus();
    }

public void BeginRecording(string assetsDir)
    {
        _csvPath = null;
        _imagesDir = assetsDir;
        _isRecordingMode = true;
        _thumbsEnabled = false;
        _script = new Script { App = new AppTarget { ProcessName = "" }, Steps = new List<Step>() };
        _suppressSave = true;
        try { _grid.Rows.Clear(); }
        finally { _suppressSave = false; }
        UpdateStatus();
    }

    
    public void EndRecording()
    {
        _isRecordingMode = false;
        _thumbsEnabled = true;

        // Refresh thumbnails once (fast enough) now that recording is stopped
        try
        {
            foreach (DataGridViewRow r in _grid.Rows)
            {
                var imgFile = r.Cells["Image"].Value?.ToString();
                r.Cells["Thumb"].Value = LoadThumb(imgFile);
            }
        }
        catch { }

        UpdateButtons();
        UpdateStatus();
    }

public void AppendStep(Step step)
    {
        // Clone to avoid sharing mutable objects across threads
        var s = new Step
        {
            Action = step.Action,
            Value = step.Value,
            Note = step.Note,
            DelayMs = step.DelayMs,
            Target = new Target
            {
                AutomationId = step.Target.AutomationId,
                Name = step.Target.Name,
                ClassName = step.Target.ClassName,
                ControlType = step.Target.ControlType,
                Image = step.Target.Image
            }
        };

        _script.Steps.Add(s);

        var thumb = _thumbsEnabled ? LoadThumb(s.Target.Image) : null;
        _grid.Rows.Add(
            thumb,
            s.Action,
            s.Target.AutomationId ?? "",
            s.Target.Name ?? "",
            s.Target.ClassName ?? "",
            s.Target.ControlType ?? "",
            s.Target.Image ?? "",
            s.Value ?? "",
            s.DelayMs?.ToString() ?? "",
            s.Note ?? ""
        );

        if (_grid.Rows.Count > 0)
            _grid.CurrentCell = _grid.Rows[_grid.Rows.Count - 1].Cells[1];

        UpdateButtons();
        UpdateStatus();
    }
public void ReplaceSteps(List<Step> steps)
    {
        steps ??= new List<Step>();
        _script.Steps.Clear();
        _grid.Rows.Clear();
        foreach (var s in steps)
            AppendStep(s);
    }

public void SetCsvPath(string? csvPath)
    {
        _csvPath = csvPath;
        UpdateStatus();
    }


    private void EnableDragDropReorder()
    {
        _grid.AllowDrop = true;

        _grid.MouseDown += (_, e) =>
        {
            try
            {
                var hit = _grid.HitTest(e.X, e.Y);
                _dragRowIndex = hit.RowIndex;
            }
            catch { _dragRowIndex = -1; }
        };

        _grid.MouseMove += (_, e) =>
        {
            if ((e.Button & MouseButtons.Left) == 0) return;
            if (_dragRowIndex < 0 || _dragRowIndex >= _grid.Rows.Count) return;
            try
            {
                _grid.DoDragDrop(_grid.Rows[_dragRowIndex], DragDropEffects.Move);
            }
            catch { }
        };

        _grid.DragOver += (_, e) =>
        {
            e.Effect = DragDropEffects.Move;
        };

        _grid.DragDrop += (_, e) =>
        {
            try
            {
                var client = _grid.PointToClient(new Point(e.X, e.Y));
                var hit = _grid.HitTest(client.X, client.Y);
                int dropIndex = hit.RowIndex;
                if (dropIndex < 0) dropIndex = _grid.Rows.Count - 1;

                if (_dragRowIndex < 0 || _dragRowIndex >= _grid.Rows.Count) return;
                if (dropIndex == _dragRowIndex) return;

                _suppressSave = true;
                try
                {
                    SwapRows(_dragRowIndex, dropIndex);
                    _grid.CurrentCell = _grid.Rows[dropIndex].Cells[1];
                }
                finally
                {
                    _suppressSave = false;
                }

                _dragRowIndex = -1;
                SaveSoon();
                UpdateButtons();
                UpdateStatus();
            }
            catch { }
        };
    }

    private void BuildGrid()
    {
        _grid.Columns.Clear();

        var colImg = new DataGridViewImageColumn
        {
            Name = "Thumb",
            HeaderText = "Flow",
            Width = 180,
            ImageLayout = DataGridViewImageCellLayout.Zoom
        };
        _grid.Columns.Add(colImg);
        _grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "ImageOnly", HeaderText = "ImageOnly", Width = 70 });

        _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Action", HeaderText = "Action", Width = 110 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "AutomationId", HeaderText = "AutomationId", Width = 120 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Name", HeaderText = "Name", Width = 140 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "ClassName", HeaderText = "ClassName", Width = 120 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "ControlType", HeaderText = "ControlType", Width = 90 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Image", HeaderText = "Image (file)", Width = 140 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Value", HeaderText = "Value", Width = 180 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "DelayMs", HeaderText = "DelayMs", Width = 70 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Note", HeaderText = "Title/Note", Width = 180 });

        _grid.RowTemplate.Height = 78;
    }

    private void RebuildRows()
    {
        _suppressSave = true;
        try
        {
            _grid.Rows.Clear();

            foreach (var s in _script.Steps)
            {
                var thumb = _thumbsEnabled ? LoadThumb(s.Target.Image) : null;
                _grid.Rows.Add(
                    thumb,
                    s.Action,
                    s.Target.AutomationId ?? "",
                    s.Target.Name ?? "",
                    s.Target.ClassName ?? "",
                    s.Target.ControlType ?? "",
                    s.Target.Image ?? "",
                    s.Value ?? "",
                    s.DelayMs?.ToString() ?? "",
                    s.Note ?? ""
                );
            }
        }
        finally
        {
            _suppressSave = false;
        }
        UpdateButtons();
    }

    private Image? LoadThumb(string? imageFile)
    {
        if (!_thumbsEnabled) return null;
        try
        {
            if (string.IsNullOrWhiteSpace(imageFile) || string.IsNullOrWhiteSpace(_imagesDir)) return null;

            var full = Path.Combine(_imagesDir, imageFile);
            if (!File.Exists(full)) return null;

            using var tmp = new Bitmap(full);
            return new Bitmap(tmp);
        }
        catch { return null; }
    }

    private void UpdateButtons()
    {
        var hasSel = _grid.CurrentRow != null && _grid.CurrentRow.Index >= 0;
        if (_isRecordingMode)
        {
            _btnMoveUp.Enabled = false;
            _btnMoveDown.Enabled = false;
            _btnDelete.Enabled = false;
            return;
        }

        _btnMoveUp.Enabled = hasSel && _grid.CurrentRow!.Index > 0;
        _btnMoveDown.Enabled = hasSel && _grid.CurrentRow!.Index >= 0 && _grid.CurrentRow!.Index < _grid.Rows.Count - 1;
        _btnDelete.Enabled = hasSel;
    }

    private void UpdateStatus()
    {
        var csv = string.IsNullOrWhiteSpace(_csvPath) ? "(not set)" : _csvPath;
        var img = string.IsNullOrWhiteSpace(_imagesDir) ? "(not set)" : _imagesDir;
        _lblStatus.Text = $"CSV: {csv}   |   Images: {img}   |   Steps: {_grid.Rows.Count}";
    }

    private void Import()
    {
        using var csvDlg = new OpenFileDialog
        {
            Filter = "CSV data (*.csv)|*.csv|All files (*.*)|*.*",
            Title = "Select CSV file"
        };

        if (csvDlg.ShowDialog(FindForm()) != DialogResult.OK) return;

        using var folderDlg = new FolderBrowserDialog
        {
            Description = "Select folder containing flow PNG images (optional)"
        };

        string? imgDir = null;
        if (folderDlg.ShowDialog(FindForm()) == DialogResult.OK)
            imgDir = folderDlg.SelectedPath;

        try
        {
            _csvPath = csvDlg.FileName;
            _imagesDir = !string.IsNullOrWhiteSpace(imgDir) ? imgDir : Path.GetDirectoryName(_csvPath);
            _isRecordingMode = false;
            _thumbsEnabled = true;

            _script = ScriptCsv.Load(_csvPath);
            RebuildRows();
            UpdateStatus();

            Log?.Invoke($"[FLOW] Imported CSV: {_csvPath}");
        }
        catch (Exception ex)
        {
            MessageBox.Show(FindForm(), ex.Message, "Import failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            Log?.Invoke("[FLOW] Import error: " + ex);
        }
    }

    private void SaveAs()
    {
        using var dlg = new SaveFileDialog
        {
            Filter = "CSV data (*.csv)|*.csv|All files (*.*)|*.*",
            FileName = Path.GetFileName(_csvPath ?? "recording.csv"),
            OverwritePrompt = true
        };

        if (dlg.ShowDialog(FindForm()) != DialogResult.OK) return;

        _csvPath = dlg.FileName;
        SaveIfReady(force: true);
        UpdateStatus();
    }

    private void DeleteSelected()
    {
        if (_grid.CurrentRow == null) return;
        int idx = _grid.CurrentRow.Index;
        if (idx < 0 || idx >= _grid.Rows.Count) return;

        _suppressSave = true;
        try
        {
            _grid.Rows.RemoveAt(idx);
        }
        finally
        {
            _suppressSave = false;
        }

        // Select a safe row after delete
        if (_grid.Rows.Count > 0)
        {
            int newIdx = Math.Min(idx, _grid.Rows.Count - 1);
            _grid.CurrentCell = _grid.Rows[newIdx].Cells[1];
        }

        SaveSoon();
        _saveTimer.Tick += (_, __) =>
        {
            _saveTimer.Stop();
            try { SaveIfReady(); } catch (Exception ex) { Log?.Invoke("[FLOW] Debounced save error: " + ex); }
        };

        UpdateButtons();
        UpdateStatus();
    }

    private void MoveSelected(int delta)
    {
        if (_grid.CurrentRow == null) return;
        int i = _grid.CurrentRow.Index;
        int j = i + delta;
        if (j < 0 || j >= _grid.Rows.Count) return;

        _suppressSave = true;
        try
        {
            SwapRows(i, j);
            _grid.CurrentCell = _grid.Rows[j].Cells[1];
        }
        finally
        {
            _suppressSave = false;
        }

        SaveSoon();
        UpdateButtons();
        UpdateStatus();
    }

    private void SwapRows(int a, int b)
    {
        if (a == b) return;

        // Swap all editable cells including thumbnail + text fields
        foreach (DataGridViewColumn col in _grid.Columns)
        {
            var va = _grid.Rows[a].Cells[col.Index].Value;
            var vb = _grid.Rows[b].Cells[col.Index].Value;
            _grid.Rows[a].Cells[col.Index].Value = vb;
            _grid.Rows[b].Cells[col.Index].Value = va;
        }
    }


    private void SaveSoon()
    {
        if (_suppressSave) return;
        // Debounce saves to avoid UI freezes during edits/moves/deletes.
        _saveTimer.Stop();
        _saveTimer.Start();
    }

    private void SaveIfReady(bool force = false)
    {
        if (_suppressSave) return;
        if (string.IsNullOrWhiteSpace(_csvPath) && !force) return;

        try
        {
            // Rebuild Script from grid, then save CSV.
            var script = new Script { App = _script.App, Steps = new List<Step>() };

            foreach (DataGridViewRow r in _grid.Rows)
            {
                var step = new Step
                {
                    Action = GetStr(r, "Action"),
                    Value = NullIfEmpty(GetStr(r, "Value")),
                    DelayMs = int.TryParse(GetStr(r, "DelayMs"), out var d) ? d : null,
                    Note = NullIfEmpty(GetStr(r, "Note")),
                    Target = new Target
                    {
                        AutomationId = NullIfEmpty(GetStr(r, "AutomationId")),
                        Name = NullIfEmpty(GetStr(r, "Name")),
                        ClassName = NullIfEmpty(GetStr(r, "ClassName")),
                        ControlType = NullIfEmpty(GetStr(r, "ControlType")),
                        Image = NullIfEmpty(GetStr(r, "Image"))
                    }
                };
                script.Steps.Add(step);
            }

            _script = script;

            if (!string.IsNullOrWhiteSpace(_csvPath))
            {
                ScriptCsv.Save(_script, _csvPath!);
                Log?.Invoke($"[FLOW] CSV updated: {_csvPath}");
            }

            UpdateStatus();
        }
        catch (Exception ex)
        {
            Log?.Invoke("[FLOW] Save error: " + ex);
        }
    }

    private static string GetStr(DataGridViewRow row, string colName)
        => row.Cells[colName].Value?.ToString() ?? "";

    private static string? NullIfEmpty(string? v)
        => string.IsNullOrWhiteSpace(v) ? null : v;

    public int GetSelectedIndex()
        => _grid.CurrentRow?.Index ?? 0;

    public void SetRunningIndex(int index)
    {
        try
        {
            if (index < 0 || index >= _grid.Rows.Count) return;
            _grid.ClearSelection();
            _grid.Rows[index].Selected = true;
            _grid.CurrentCell = _grid.Rows[index].Cells[1];
            _grid.FirstDisplayedScrollingRowIndex = Math.Max(0, index - 2);
        }
        catch { }
    }

    private void PickOcrRegionForSelected()
    {
        if (_grid.CurrentRow == null) return;
        var row = _grid.CurrentRow;

        var action = row.Cells["Action"].Value?.ToString() ?? "";
        if (!string.Equals(action, "OcrRead", StringComparison.OrdinalIgnoreCase))
        {
            Log?.Invoke("[FLOW] Select an 'OcrRead' step row first.");
            return;
        }

        using var picker = new RegionPickerOverlay();
        var res = picker.ShowDialog(FindForm());
        if (res != DialogResult.OK) return;

        var r = picker.SelectedRect;
        var coords = $"{r.X},{r.Y},{r.Width},{r.Height}";

        var existing = row.Cells["Value"].Value?.ToString() ?? "";
        if (string.IsNullOrWhiteSpace(existing))
        {
            // Default save var if none provided yet.
            row.Cells["Value"].Value = coords + "|save=OcrText";
        }
        else
        {
            var parts = existing.Split('|', StringSplitOptions.TrimEntries);
            if (parts.Length >= 2)
            {
                // Replace first token (rect / rel: token), keep the rest.
                row.Cells["Value"].Value = coords + "|" + string.Join('|', parts.Skip(1));
            }
            else
            {
                // Treat the existing token as shorthand (e.g. VarName or save=...), append it.
                row.Cells["Value"].Value = coords + "|" + existing.Trim();
            }
        }

        SaveSoon();
        Log?.Invoke($"[FLOW] OCR region set: {coords}");
    }

}