using System.Drawing;
using System.Windows.Forms;
using WorksheetWatcher.Services;

namespace WorksheetWatcher.Forms;

/// <summary>
/// One student's page, full window. Non-modal, so the main grid keeps polling and updating
/// behind it; <see cref="MainForm"/> hands this window each fresh export via
/// <see cref="SetContent"/> for as long as it stays open.
///
/// It renders straight from the page's PDF at the size of the window, so it is sharp at any
/// screen size rather than a stretched thumbnail. Two views, switched with the button at the
/// bottom: the close-up of the most recent change (the default, same spot as the tile), or
/// every sheet of the page stacked and scrollable, with the changed area boxed in red.
/// </summary>
public sealed class FullScreenViewForm : Form
{
    private readonly string _studentName;

    private readonly Label _nameLabel = new()
    {
        Dock = DockStyle.Top,
        Height = 44,
        TextAlign = ContentAlignment.MiddleCenter,
        Font = new Font("Trebuchet MS", 16f, FontStyle.Bold)
    };

    private readonly PictureBox _closeUp = new()
    {
        Dock = DockStyle.Fill,
        SizeMode = PictureBoxSizeMode.Zoom,
        BackColor = Color.White
    };

    private readonly Panel _pages = new()
    {
        Dock = DockStyle.Fill,
        AutoScroll = true,
        BackColor = Color.FromArgb(0x60, 0x64, 0x6B),
        Visible = false
    };

    private readonly Button _btnToggle = new()
    {
        Text = "Show whole page",
        Dock = DockStyle.Left,
        Width = 220,
        Font = new Font("Trebuchet MS", 10f)
    };

    // Re-rendering on every pixel of a drag-resize would be wasteful; wait for it to settle.
    private readonly System.Windows.Forms.Timer _resizeTimer = new() { Interval = 250 };

    private byte[]? _pdf;
    private PageFocus? _focus;
    private bool _wholePage;
    private int _generation;

    /// <summary>The student this window is showing, so the host knows whether to forward an update.</summary>
    public string StudentId { get; }

    public FullScreenViewForm(string studentId, string studentName)
    {
        StudentId = studentId;
        _studentName = studentName;

        Text = $"{studentName} - Worksheet Watcher";
        Icon = AppIcon.Value;
        StartPosition = FormStartPosition.CenterParent;
        WindowState = FormWindowState.Maximized;
        MinimumSize = new Size(640, 480);
        Font = new Font("Trebuchet MS", 9.75f);
        KeyPreview = true;

        _nameLabel.Text = studentName;

        var btnClose = new Button { Text = "Close", Dock = DockStyle.Fill, Font = new Font("Trebuchet MS", 10f) };
        btnClose.Click += (_, _) => Close();
        _btnToggle.Click += (_, _) => ToggleView();

        var bottom = new Panel { Dock = DockStyle.Bottom, Height = 40 };
        bottom.Controls.Add(btnClose);
        bottom.Controls.Add(_btnToggle);

        Controls.Add(_closeUp);
        Controls.Add(_pages);
        Controls.Add(bottom);
        Controls.Add(_nameLabel);

        _resizeTimer.Tick += (_, _) => { _resizeTimer.Stop(); Render(); };
        SizeChanged += (_, _) => { _resizeTimer.Stop(); _resizeTimer.Start(); };
        Shown += (_, _) => Render();
        KeyDown += (_, e) => { if (e.KeyCode == Keys.Escape) Close(); };
        FormClosed += (_, _) => ReleaseImages();
    }

    /// <summary>
    /// Shows a (new) export of the page: <paramref name="pdf"/> is the whole PDF and
    /// <paramref name="focus"/> where the most recent change is. Called when the window
    /// opens and again whenever this student's page changes while it stays open.
    /// </summary>
    public void SetContent(byte[]? pdf, PageFocus? focus)
    {
        _pdf = pdf;
        _focus = focus;
        if (IsHandleCreated) Render();
    }

    private void ToggleView()
    {
        _wholePage = !_wholePage;
        _btnToggle.Text = _wholePage ? "Show close-up of change" : "Show whole page";
        _closeUp.Visible = !_wholePage;
        _pages.Visible = _wholePage;
        Render();
    }

    /// <summary>Fire-and-forget by design: everything inside is guarded, nothing here may throw.</summary>
    private async void Render()
    {
        var generation = ++_generation;
        if (_pdf is null) return;

        try
        {
            using var pdf = await OpenedPdf.OpenAsync(_pdf);
            if (_wholePage) await RenderWholePageAsync(pdf, generation);
            else await RenderCloseUpAsync(pdf, generation);
        }
        catch (ObjectDisposedException)
        {
            // window closed mid-render
        }
        catch (Exception ex)
        {
            if (generation == _generation && !IsDisposed)
                _nameLabel.Text = $"{_studentName} - could not render: {ex.Message}";
        }
    }

    private async Task RenderCloseUpAsync(OpenedPdf pdf, int generation)
    {
        var pageIndex = Math.Min(_focus?.PageIndex ?? 0, pdf.PageCount - 1);
        var window = _focus?.Window ?? new RectangleF(PointF.Empty, pdf.PageSize(pageIndex));

        var area = _closeUp.ClientSize;
        if (area.Width < 50 || area.Height < 50) return;

        var scale = Math.Min(area.Width / window.Width, area.Height / window.Height);
        var width = (int)Math.Min(4000, window.Width * scale);
        var height = (int)Math.Min(4000, window.Height * scale);

        var bitmap = await pdf.RenderAsync(pageIndex, window, width, height);
        if (generation != _generation || IsDisposed)
        {
            bitmap.Dispose();
            return;
        }

        var previous = _closeUp.Image;
        _closeUp.Image = bitmap;
        previous?.Dispose();
    }

    private async Task RenderWholePageAsync(OpenedPdf pdf, int generation)
    {
        // Capped: on a very wide monitor a sheet as wide as the window would be enormous and
        // need endless scrolling. Sheets are centred instead.
        var available = Math.Max(300, _pages.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - 24);
        var width = Math.Min(available, 1500);
        var bitmaps = new List<Bitmap>();

        for (var i = 0; i < pdf.PageCount; i++)
        {
            var size = pdf.PageSize(i);
            var height = (int)Math.Round(width * size.Height / size.Width);
            var bitmap = await pdf.RenderAsync(i, null, width, height);

            if (_focus is { } focus && focus.PageIndex == i)
            {
                var scale = width / size.Width;
                using var graphics = Graphics.FromImage(bitmap);
                using var pen = new Pen(Color.FromArgb(0xE5, 0x39, 0x35), 4);
                graphics.DrawRectangle(pen, focus.Window.X * scale, focus.Window.Y * scale, focus.Window.Width * scale, focus.Window.Height * scale);
            }

            bitmaps.Add(bitmap);

            if (generation != _generation || IsDisposed)
            {
                foreach (var b in bitmaps) b.Dispose();
                return;
            }
        }

        // First time round, jump to the change; on later refreshes leave the teacher where
        // they had scrolled to rather than yanking the view away mid-read.
        var firstTime = _pages.Controls.Count == 0;
        var savedScroll = -_pages.AutoScrollPosition.Y;

        ClearPages();

        var y = 12;
        var focusTop = 0;
        for (var i = 0; i < bitmaps.Count; i++)
        {
            if (_focus is { } focus && focus.PageIndex == i)
                focusTop = y + (int)(focus.Window.Y * width / pdf.PageSize(i).Width);

            _pages.Controls.Add(new PictureBox
            {
                Image = bitmaps[i],
                Size = bitmaps[i].Size,
                Location = new Point(12 + (available - width) / 2, y)
            });
            y += bitmaps[i].Height + 16;
        }

        _pages.AutoScrollMinSize = new Size(0, y);
        _pages.AutoScrollPosition = new Point(0, firstTime ? Math.Max(0, focusTop - 60) : savedScroll);
    }

    private void ClearPages()
    {
        foreach (Control c in _pages.Controls)
        {
            (c as PictureBox)?.Image?.Dispose();
            c.Dispose();
        }
        _pages.Controls.Clear();
    }

    private void ReleaseImages()
    {
        _resizeTimer.Dispose();
        _closeUp.Image?.Dispose();
        _closeUp.Image = null;
        ClearPages();
    }
}
