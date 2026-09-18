using System.Drawing;
using System.Windows.Forms;
using WorksheetWatcher.Models;

namespace WorksheetWatcher.Forms;

/// <summary>
/// One student's tile in the watcher grid: a rasterised page preview, the student's
/// name, a "last updated" caption, and a status-coloured stripe (green = just updated,
/// grey = no page yet, amber = section locked, red = error). Raises
/// <see cref="TileActivated"/> on click so the host can open the full-screen view, and
/// <see cref="TileHoverEnter"/> / <see cref="TileHoverLeave"/> so the host can show a
/// larger "peek" preview in a fixed spot on screen - deliberately not positioned near the
/// tile itself, since a tile scrolled low in the grid would push a cursor-anchored popup
/// off the bottom of the screen.
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

    private PollStatus _status = PollStatus.NotStarted;
    private string? _statusMessage;
    private DateTime? _lastRendered;

    /// <summary>The student's section-group hierarchy ID.</summary>
    public string StudentId { get; }

    /// <summary>The student's name, shown at the top of the tile.</summary>
    public string StudentName { get; }

    /// <summary>The child controls that fill the tile's whole clickable/hoverable area.</summary>
    private Control[] HitAreaControls => new Control[] { this, _picture, _nameLabel, _statusLabel, _statusStripe };

    /// <summary>Raised when the tile is clicked - the host opens the full-screen view.</summary>
    public event EventHandler? TileActivated;

    /// <summary>Raised when the cursor enters any part of the tile - the host shows the zoom overlay.</summary>
    public event EventHandler? TileHoverEnter;

    /// <summary>
    /// Raised when the cursor leaves any part of the tile. Moving between two of the
    /// tile's own child controls (picture -&gt; label, say) fires this and then
    /// <see cref="TileHoverEnter"/> again in the same tick; the host debounces that rather
    /// than this control trying to.
    /// </summary>
    public event EventHandler? TileHoverLeave;

    /// <summary>The tile's current rendered preview, for the host's zoom overlay to draw from.</summary>
    public Image? CurrentImage => _picture.Image;

    public StudentThumbnailControl(string studentId, string studentName)
    {
        StudentId = studentId;
        StudentName = studentName;

        BorderStyle = BorderStyle.FixedSingle;
        Margin = new Padding(6);
        _nameLabel.Text = studentName;

        Controls.Add(_picture);
        Controls.Add(_statusLabel);
        Controls.Add(_nameLabel);
        Controls.Add(_statusStripe);

        foreach (var c in HitAreaControls)
        {
            c.Click += (_, _) => TileActivated?.Invoke(this, EventArgs.Empty);
            c.MouseEnter += (_, _) => TileHoverEnter?.Invoke(this, EventArgs.Empty);
            c.MouseLeave += (_, _) => TileHoverLeave?.Invoke(this, EventArgs.Empty);
        }
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
