using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using WorksheetWatcher.Models;
using WorksheetWatcher.Services;

namespace WorksheetWatcher.Forms;

/// <summary>
/// The application's main window: pick an open notebook and a worksheet (a page title
/// recurring across students), start watching, and see a scrollable grid of up to
/// <c>AppConfig.MaxStudents</c> live-updating student thumbnails. Click a tile to see that
/// student full screen; hover a tile to peek at a larger preview without leaving the grid.
/// </summary>
public sealed class MainForm : Form
{
    private const string BaseFontFamily = "Trebuchet MS";
    private const float BaseFontSize = 9.75f;

    // The only poll cadences offered in the UI. Measured against a real 23-student
    // notebook, the per-tick hierarchy re-check costs 50-200ms regardless of interval, so
    // even 5s is safe - the real pacing limit is how many pages actually changed that
    // tick, since each one costs a separate ~1-2s Publish/rasterise call.
    private static readonly int[] IntervalChoiceSeconds = { 5, 10, 15, 30 };

    private readonly AppConfig _config;
    private readonly UserSettings _settings;
    private WatcherPollingService? _poller;
    private readonly Dictionary<string, ThumbnailState> _thumbnails = new();
    private readonly Dictionary<string, StudentThumbnailControl> _tiles = new();
    private FullScreenViewForm? _fullScreen;
    private IReadOnlyList<NotebookInfo> _notebooks = Array.Empty<NotebookInfo>();

    private readonly ComboBox _cboNotebook = new()
    {
        DropDownStyle = ComboBoxStyle.DropDownList,
        Width = 360,
        DropDownWidth = 360,
        IntegralHeight = false,
        MaxDropDownItems = 16,
        Margin = new Padding(0, 2, 8, 2)
    };

    private readonly Button _btnRefreshNotebooks = new() { Text = "Refresh" };

    private readonly ComboBox _cboWorksheet = new()
    {
        DropDownStyle = ComboBoxStyle.DropDown,
        Width = 280,
        DropDownWidth = 280,
        IntegralHeight = false,
        MaxDropDownItems = 20,
        Margin = new Padding(0, 2, 8, 2)
    };

    private readonly ComboBox _cboInterval = new()
    {
        DropDownStyle = ComboBoxStyle.DropDownList,
        Width = 70,
        Margin = new Padding(0, 2, 8, 2)
    };

    private readonly Button _btnStart = new() { Text = "▶  Start Watching" };
    private readonly Button _btnStop = new() { Text = "■  Stop", Enabled = false };
    private readonly Button _btnRefreshNow = new() { Text = "Refresh now", Enabled = false };

    private readonly Label _lblStatus = new()
    {
        Text = "Pick a notebook to begin.",
        AutoSize = false,
        Dock = DockStyle.Fill,
        TextAlign = ContentAlignment.MiddleLeft
    };

    private readonly FlowLayoutPanel _grid = new()
    {
        AutoScroll = true,
        Dock = DockStyle.Fill,
        WrapContents = true,
        FlowDirection = FlowDirection.LeftToRight,
        BackColor = Color.WhiteSmoke,
        Padding = new Padding(8)
    };

    // Ticks the "updated Ns ago" captions between poll cycles - no OneNote calls here.
    private readonly System.Windows.Forms.Timer _captionTimer = new() { Interval = 5000 };

    // The zoom overlay's screen position is always centred on this window, never anchored
    // to the hovered tile - a tile scrolled low in the grid would otherwise push a
    // cursor-anchored popup off the bottom of the screen. It just redraws whatever bitmap
    // the tile already has cached, so hovering costs nothing extra (no OneNote calls).
    private readonly Panel _zoomOverlay = new()
    {
        Visible = false,
        BackColor = Color.Black,
        Padding = new Padding(1) // a thin dark frame around the picture
    };
    private readonly PictureBox _zoomPicture = new()
    {
        Dock = DockStyle.Fill,
        SizeMode = PictureBoxSizeMode.Zoom,
        BackColor = Color.White
    };
    private readonly Label _zoomCaption = new()
    {
        Dock = DockStyle.Top,
        Height = 28,
        TextAlign = ContentAlignment.MiddleCenter,
        Font = new Font(BaseFontFamily, 10.5f, FontStyle.Bold),
        BackColor = Color.Black,
        ForeColor = Color.White
    };

    // Tracks which tile is currently hovered so a poll update mid-hover can keep the
    // overlay's image live, and so leaving one child control and entering a sibling of the
    // same tile doesn't flicker the overlay off and on.
    private StudentThumbnailControl? _hoveredTile;
    private readonly System.Windows.Forms.Timer _hoverLeaveTimer = new() { Interval = 120 };

    public MainForm()
    {
        _config = AppConfig.Load(out var configWarning);
        _settings = UserSettings.Load();

        foreach (var seconds in IntervalChoiceSeconds) _cboInterval.Items.Add($"{seconds}s");
        var closestInterval = IntervalChoiceSeconds.OrderBy(s => Math.Abs(s - _config.PollIntervalSeconds)).First();
        _cboInterval.SelectedIndex = Array.IndexOf(IntervalChoiceSeconds, closestInterval);

        AutoScaleMode = AutoScaleMode.Font;
        AutoScaleDimensions = new SizeF(7f, 16f);

        foreach (var b in new[] { _btnRefreshNotebooks, _btnStart, _btnStop, _btnRefreshNow })
            StyleButton(b);

        Text = "Worksheet Watcher - Live Class Notebook Viewer";
        Icon = AppIcon.Value;
        Font = new Font(BaseFontFamily, BaseFontSize);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.Sizable;
        MinimumSize = new Size(760, 520);
        Size = new Size(1200, 800);

        BuildLayout();
        WireEvents();

        if (configWarning is not null)
            _lblStatus.Text = configWarning;

        Shown += async (_, _) => await LoadNotebooksAsync();
        Resize += (_, _) => { if (_zoomOverlay.Visible) PositionZoomOverlay(); };
        FormClosing += (_, _) => { _poller?.Dispose(); _captionTimer.Dispose(); _hoverLeaveTimer.Dispose(); };
    }

    /// <summary>
    /// Common button styling: size to the text plus generous padding, never collapse below
    /// a real minimum height, and scale with the ambient font.
    /// </summary>
    private static void StyleButton(Button b)
    {
        b.AutoSize = true;
        b.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        b.Padding = new Padding(14, 6, 14, 6);
        b.MinimumSize = new Size(88, 30);
        b.Margin = new Padding(0, 2, 8, 2);
        b.UseVisualStyleBackColor = true;
    }

    private void BuildLayout()
    {
        var top = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(8),
            WrapContents = true
        };

        top.Controls.Add(new Label { Text = "Notebook:", AutoSize = true, Margin = new Padding(0, 8, 4, 0) });
        top.Controls.Add(_cboNotebook);
        top.Controls.Add(_btnRefreshNotebooks);
        top.Controls.Add(new Label { Text = "Worksheet:", AutoSize = true, Margin = new Padding(16, 8, 4, 0) });
        top.Controls.Add(_cboWorksheet);
        top.Controls.Add(new Label { Text = "Poll every:", AutoSize = true, Margin = new Padding(16, 8, 4, 0) });
        top.Controls.Add(_cboInterval);
        top.Controls.Add(_btnStart);
        top.Controls.Add(_btnStop);
        top.Controls.Add(_btnRefreshNow);

        var statusPanel = new Panel { Dock = DockStyle.Top, Height = 26, Padding = new Padding(8, 0, 8, 4) };
        statusPanel.Controls.Add(_lblStatus);

        _zoomOverlay.Controls.Add(_zoomPicture);
        _zoomOverlay.Controls.Add(_zoomCaption);

        Controls.Add(_grid);
        Controls.Add(statusPanel);
        Controls.Add(top);
        Controls.Add(_zoomOverlay);
    }

    private void WireEvents()
    {
        _btnRefreshNotebooks.Click += async (_, _) => await LoadNotebooksAsync();
        _cboNotebook.SelectedIndexChanged += async (_, _) => await LoadWorksheetTitlesAsync();
        _btnStart.Click += (_, _) => StartWatching();
        _btnStop.Click += (_, _) => StopWatching();
        _btnRefreshNow.Click += (_, _) => _poller?.RequestImmediateRefresh();
        _captionTimer.Tick += (_, _) => { foreach (var tile in _tiles.Values) tile.RefreshCaption(); };
        _captionTimer.Start();

        // A tile's own child controls each fire enter/leave as the cursor crosses between
        // them, so a leave is not trusted until this timer confirms the cursor is truly
        // outside the tile's bounds - otherwise moving from the picture to the name label
        // would flicker the overlay off and straight back on.
        _hoverLeaveTimer.Tick += (_, _) =>
        {
            _hoverLeaveTimer.Stop();
            if (_hoveredTile is null) return;
            var stillOver = _hoveredTile.RectangleToScreen(_hoveredTile.ClientRectangle).Contains(Cursor.Position);
            if (!stillOver)
            {
                _hoveredTile = null;
                HideZoomOverlay();
            }
        };
    }

    /// <summary>Shows the zoom overlay for <paramref name="tile"/>, centred on this window.</summary>
    private void ShowZoomOverlay(StudentThumbnailControl tile)
    {
        _hoveredTile = tile;
        if (tile.CurrentImage is null) return;

        _zoomCaption.Text = tile.StudentName;
        _zoomPicture.Image = tile.CurrentImage;
        PositionZoomOverlay();
        _zoomOverlay.Visible = true;
        _zoomOverlay.BringToFront();
    }

    private void HideZoomOverlay()
    {
        _zoomOverlay.Visible = false;
        _zoomPicture.Image = null;
    }

    /// <summary>
    /// Sizes and centres the overlay within this window's own client area - deliberately
    /// independent of the hovered tile's position, so it is always fully visible no matter
    /// where in the (possibly scrolled) grid that tile sits.
    /// </summary>
    private void PositionZoomOverlay()
    {
        var w = (int)(ClientSize.Width * 0.7);
        var h = (int)(ClientSize.Height * 0.7);
        _zoomOverlay.Size = new Size(w, h);
        _zoomOverlay.Location = new Point((ClientSize.Width - w) / 2, (ClientSize.Height - h) / 2);
    }

    /// <summary>Lists every notebook currently open in OneNote, off the UI thread.</summary>
    private async Task LoadNotebooksAsync()
    {
        _lblStatus.Text = "Loading notebooks...";
        _cboNotebook.Enabled = false;
        try
        {
            _notebooks = await Task.Run(() =>
            {
                using var hierarchy = new OneNoteHierarchyService(_config);
                return hierarchy.GetOpenNotebooks();
            });

            _cboNotebook.Items.Clear();
            foreach (var nb in _notebooks) _cboNotebook.Items.Add(nb);
            _cboNotebook.DisplayMember = nameof(NotebookInfo.DisplayName);

            var preselect = _settings.LastNotebookId is null
                ? null
                : _notebooks.FirstOrDefault(n => n.Id == _settings.LastNotebookId);
            _cboNotebook.SelectedItem = preselect ?? _notebooks.FirstOrDefault();

            _lblStatus.Text = _notebooks.Count == 0
                ? "No notebooks are open in OneNote. Open your Class Notebook and click Refresh."
                : $"{_notebooks.Count} notebook(s) open.";
        }
        catch (OneNoteUnavailableException ex)
        {
            _lblStatus.Text = ex.Message;
        }
        finally
        {
            _cboNotebook.Enabled = true;
        }
    }

    /// <summary>
    /// Walks the picked notebook once to collect its distinct page titles, so the
    /// worksheet field can offer them - the same discovery WheresTheWork's report does
    /// for its matrix columns, just surfaced as a picker instead.
    /// </summary>
    private async Task LoadWorksheetTitlesAsync()
    {
        if (_cboNotebook.SelectedItem is not NotebookInfo notebook) return;

        var previousText = _cboWorksheet.Text;
        _cboWorksheet.Items.Clear();
        _lblStatus.Text = $"Reading '{notebook.DisplayName}'...";
        _btnStart.Enabled = false;

        try
        {
            var titles = await Task.Run(() =>
            {
                using var hierarchy = new OneNoteHierarchyService(_config);
                var resolver = new WorksheetResolutionService(hierarchy);
                return resolver.GetDistinctPageTitles(notebook.Id);
            });

            foreach (var title in titles) _cboWorksheet.Items.Add(title);

            _cboWorksheet.Text =
                !string.IsNullOrWhiteSpace(previousText) && titles.Contains(previousText, StringComparer.OrdinalIgnoreCase) ? previousText :
                _settings.LastWorksheetTitle is not null && titles.Contains(_settings.LastWorksheetTitle, StringComparer.OrdinalIgnoreCase) ? _settings.LastWorksheetTitle :
                string.Empty;

            _lblStatus.Text = $"{titles.Count} distinct page title(s) found in '{notebook.DisplayName}'.";
        }
        catch (OneNoteContentException ex)
        {
            _lblStatus.Text = $"Could not read that notebook: {ex.Message}";
        }
        finally
        {
            _btnStart.Enabled = true;
        }
    }

    private void StartWatching()
    {
        if (_cboNotebook.SelectedItem is not NotebookInfo notebook) return;
        var worksheetTitle = _cboWorksheet.Text.Trim();
        if (worksheetTitle.Length == 0)
        {
            _lblStatus.Text = "Type or pick a worksheet page title first.";
            return;
        }

        _settings.LastNotebookId = notebook.Id;
        _settings.LastWorksheetTitle = worksheetTitle;
        _settings.Save();

        _config.PollIntervalSeconds = IntervalChoiceSeconds[_cboInterval.SelectedIndex];

        ClearGrid();

        _poller?.Dispose();
        _poller = new WatcherPollingService(_config, OnThumbnailUpdate);
        _poller.Start(notebook.Id, worksheetTitle);

        _btnStart.Enabled = false;
        _btnStop.Enabled = true;
        _btnRefreshNow.Enabled = true;
        _cboNotebook.Enabled = false;
        _cboWorksheet.Enabled = false;
        _cboInterval.Enabled = false;
        _lblStatus.Text = $"Watching '{worksheetTitle}' in '{notebook.DisplayName}'...";
    }

    private void StopWatching()
    {
        _poller?.Stop();
        _btnStart.Enabled = true;
        _btnStop.Enabled = false;
        _btnRefreshNow.Enabled = false;
        _cboNotebook.Enabled = true;
        _cboWorksheet.Enabled = true;
        _cboInterval.Enabled = true;
        _lblStatus.Text = "Stopped.";
    }

    /// <summary>
    /// Called from the polling background thread - marshals to the UI thread before
    /// touching any control or the thumbnail-state dictionaries.
    /// </summary>
    private void OnThumbnailUpdate(ThumbnailUpdate update)
    {
        if (IsDisposed) return;
        try
        {
            if (InvokeRequired) { BeginInvoke(new Action(() => ApplyUpdate(update))); return; }
            ApplyUpdate(update);
        }
        catch (ObjectDisposedException)
        {
            // window closed mid-update - ignore
        }
    }

    private void ApplyUpdate(ThumbnailUpdate update)
    {
        if (update.StudentId.Length == 0)
        {
            _lblStatus.Text = update.StatusMessage ?? "OneNote error.";
            return;
        }

        if (!_thumbnails.TryGetValue(update.StudentId, out var state))
        {
            state = new ThumbnailState { StudentId = update.StudentId, StudentName = update.StudentName };
            _thumbnails[update.StudentId] = state;
        }

        if (update.Image is not null)
        {
            state.SetImage(update.Image);
            state.LastRendered = DateTime.Now;
        }
        state.LastModified = update.LastModified;
        state.Status = update.Status;
        state.StatusMessage = update.StatusMessage;

        if (!_tiles.TryGetValue(update.StudentId, out var tile))
        {
            tile = new StudentThumbnailControl(update.StudentId, update.StudentName)
            {
                Width = _config.TileDisplayWidth + 4,
                Height = _config.TileDisplayHeight + 48
            };
            tile.TileActivated += (_, _) => OpenFullScreen(update.StudentId);
            tile.TileHoverEnter += (_, _) => { _hoverLeaveTimer.Stop(); ShowZoomOverlay(tile); };
            tile.TileHoverLeave += (_, _) => { _hoverLeaveTimer.Stop(); _hoverLeaveTimer.Start(); };
            _tiles[update.StudentId] = tile;
            _grid.Controls.Add(tile);
        }

        tile.SetImage(state.Image);
        tile.SetStatus(state.Status, state.StatusMessage, state.LastRendered);

        if (_fullScreen is { IsDisposed: false } && _fullScreen.StudentId == update.StudentId)
            _fullScreen.UpdateImage(state.Image);

        // Keep a live overlay in sync if this is the tile currently being peeked at.
        if (ReferenceEquals(_hoveredTile, tile))
            _zoomPicture.Image = state.Image;
    }

    private void OpenFullScreen(string studentId)
    {
        if (!_thumbnails.TryGetValue(studentId, out var state)) return;

        _hoverLeaveTimer.Stop();
        _hoveredTile = null;
        HideZoomOverlay();

        if (_fullScreen is { IsDisposed: false })
            _fullScreen.Close();

        _fullScreen = new FullScreenViewForm(studentId, state.StudentName, state.Image);
        _fullScreen.FormClosed += (_, _) => _fullScreen = null;
        _fullScreen.Show(this);
    }

    private void ClearGrid()
    {
        _hoverLeaveTimer.Stop();
        _hoveredTile = null;
        HideZoomOverlay();

        foreach (var tile in _tiles.Values) tile.Dispose();
        _tiles.Clear();
        foreach (var state in _thumbnails.Values) state.Dispose();
        _thumbnails.Clear();
        _grid.Controls.Clear();
    }
}
