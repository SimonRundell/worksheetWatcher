using System.Drawing;
using System.Windows.Forms;
using WorksheetWatcher.Models;

namespace WorksheetWatcher.Forms;

/// <summary>
/// One student's tile in the watcher grid: a rasterised page preview, the student's
/// name, a "last updated" caption, and a status-coloured stripe (green = just updated,
/// grey = no page yet, amber = section locked, red = error). Raises
/// <see cref="TileActivated"/> - clicking anywhere on the tile, or its magnifying-glass
/// button in the bottom-right corner - so the host opens the full-screen view. There is
/// deliberately no hover-triggered preview: an earlier version popped one up automatically
/// on hover and it proved obtrusive when scanning across a full grid of tiles.
/// </summary>
public sealed class StudentThumbnailControl : UserControl
{
    private readonly PictureBox _picture = new()
    {
        Dock = DockStyle.Fill,
        SizeMode = PictureBoxSizeMode.Zoom,
        BackColor = Color.White
    };

    private readonly Label _nameLabel = new()
    {
        Dock = DockStyle.Top,
        Height = 22,
        TextAlign = ContentAlignment.MiddleLeft,
        AutoEllipsis = true,
        Font = new Font("Trebuchet MS", 9.75f, FontStyle.Bold),
        Padding = new Padding(4, 0, 0, 0)
    };

    private readonly Label _statusLabel = new()
    {
        Dock = DockStyle.Bottom,
        Height = 18,
        TextAlign = ContentAlignment.MiddleLeft,
        Font = new Font("Trebuchet MS", 8f),
        ForeColor = Color.DimGray,
        Padding = new Padding(4, 0, 0, 0)
    };

    private readonly Panel _statusStripe = new() { Dock = DockStyle.Top, Height = 4, BackColor = Color.Gainsboro };

    // An explicit, always-visible affordance for "open this student full screen", sitting
    // over the bottom-right corner of the preview - deliberate and discoverable, unlike a
    // hover popup that appears uninvited.
    private readonly Button _zoomButton = new()
    {
        Text = "\U0001F50D",
        Width = 34,
        Height = 30,
        FlatStyle = FlatStyle.Flat,
        BackColor = Color.White,
        Font = new Font("Segoe UI Emoji", 12f),
        Cursor = Cursors.Hand,
        TabStop = false
    };

    private PollStatus _status = PollStatus.NotStarted;
    private string? _statusMessage;
    private DateTime? _lastRendered;

    /// <summary>The student's section-group hierarchy ID.</summary>
    public string StudentId { get; }

    /// <summary>The student's name, shown at the top of the tile.</summary>
    public string StudentName { get; }

    /// <summary>The child controls that fill the tile's whole clickable area.</summary>
    private Control[] HitAreaControls => new Control[] { this, _picture, _nameLabel, _statusLabel, _statusStripe, _zoomButton };

    /// <summary>Raised when the tile - or its zoom button - is clicked; the host opens the full-screen view.</summary>
    public event EventHandler? TileActivated;

    public StudentThumbnailControl(string studentId, string studentName)
    {
        StudentId = studentId;
        StudentName = studentName;

        BorderStyle = BorderStyle.FixedSingle;
        Margin = new Padding(6);
        _nameLabel.Text = studentName;
        _zoomButton.FlatAppearance.BorderColor = Color.Gray;

        Controls.Add(_picture);
        Controls.Add(_statusLabel);
        Controls.Add(_nameLabel);
        Controls.Add(_statusStripe);
        Controls.Add(_zoomButton); // added last so it paints on top of the picture

        foreach (var c in HitAreaControls)
            c.Click += (_, _) => TileActivated?.Invoke(this, EventArgs.Empty);

        Resize += (_, _) => PositionZoomButton();
        PositionZoomButton();
    }

    /// <summary>
    /// Keeps the zoom button pinned just above the status caption, at the tile's
    /// bottom-right - recalculated on resize rather than relying on anchoring, since the
    /// tile's size is 0 at construction time and only set afterwards by the host.
    /// </summary>
    private void PositionZoomButton()
    {
        _zoomButton.Location = new Point(
            ClientSize.Width - _zoomButton.Width - 6,
            ClientSize.Height - _statusLabel.Height - _zoomButton.Height - 6);
    }

    /// <summary>Sets the tile's rendered preview. The control does not take ownership of disposing it.</summary>
    public void SetImage(Image? image) => _picture.Image = image;

    /// <summary>Records the latest poll outcome and refreshes the stripe colour and caption.</summary>
    public void SetStatus(PollStatus status, string? message, DateTime? lastRendered)
    {
        _status = status;
        _statusMessage = message;
        _lastRendered = lastRendered;
        RefreshCaption();
    }

    /// <summary>
    /// Redraws just the "updated Ns ago" caption from the last recorded status, without a
    /// new poll result - called on a short UI timer so the age keeps counting up between
    /// poll cycles.
    /// </summary>
    public void RefreshCaption()
    {
        _statusStripe.BackColor = _status switch
        {
            PollStatus.Ok => Color.MediumSeaGreen,
            PollStatus.NotStarted => Color.Gainsboro,
            PollStatus.Locked => Color.Goldenrod,
            PollStatus.Error => Color.IndianRed,
            _ => Color.Gainsboro
        };

        _statusLabel.Text = _status switch
        {
            PollStatus.Ok when _lastRendered is not null => $"Updated {Describe(DateTime.Now - _lastRendered.Value)}",
            PollStatus.Ok => "Updated",
            PollStatus.NotStarted => "No page yet",
            PollStatus.Locked => "Section locked",
            PollStatus.Error => _statusMessage ?? "Error",
            _ => string.Empty
        };
    }

    private static string Describe(TimeSpan age)
    {
        if (age.TotalSeconds < 60) return "just now";
        if (age.TotalMinutes < 60) return $"{(int)age.TotalMinutes}m ago";
        return $"{(int)age.TotalHours}h ago";
    }
}
