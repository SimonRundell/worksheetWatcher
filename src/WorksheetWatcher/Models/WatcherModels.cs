using System.Drawing;

namespace WorksheetWatcher.Models;

/// <summary>
/// One student's resolved copy of the watched worksheet: which page it is and when it
/// was last modified, or the fact that the student has no matching page yet.
/// </summary>
/// <param name="StudentId">The student's section-group hierarchy ID.</param>
/// <param name="StudentName">The student's name (section-group name).</param>
/// <param name="SectionName">The section the page was found in, or null when there is no page.</param>
/// <param name="PageId">The page's hierarchy ID, or null when the student has not created it yet.</param>
/// <param name="PageTitle">The matched page's title, or null.</param>
/// <param name="LastModified">The page's last-modified timestamp, or null.</param>
public sealed record StudentPageTarget(
    string StudentId,
    string StudentName,
    string? SectionName,
    string? PageId,
    string? PageTitle,
    DateTime? LastModified)
{
    /// <summary>True when this student has created the worksheet page.</summary>
    public bool HasPage => PageId is not null;

    /// <summary>A target for a student who has not yet created the worksheet page.</summary>
    public static StudentPageTarget NotStarted(string studentId, string studentName) =>
        new(studentId, studentName, null, null, null, null);
}

/// <summary>The state of a single thumbnail tile's most recent poll.</summary>
public enum PollStatus
{
    /// <summary>The student has no page matching the worksheet title yet.</summary>
    NotStarted,

    /// <summary>Rendered successfully at least once.</summary>
    Ok,

    /// <summary>The page exists but OneNote refused to read it (e.g. a locked section).</summary>
    Locked,

    /// <summary>OneNote itself could not be reached.</summary>
    Error
}

/// <summary>
/// Live, mutable view-state for one student's tile: the current rendered image plus poll
/// bookkeeping. Owned by the UI thread; <see cref="SetImage"/> disposes the previous
/// bitmap so a long watch session does not leak GDI handles.
/// </summary>
public sealed class ThumbnailState : IDisposable
{
    /// <summary>The student's section-group hierarchy ID.</summary>
    public required string StudentId { get; init; }

    /// <summary>The student's name, shown under the tile.</summary>
    public required string StudentName { get; init; }

    /// <summary>The most recently resolved page target for this student.</summary>
    public StudentPageTarget? Target { get; set; }

    /// <summary>The current rendered page image, or null before the first render / when there is no page.</summary>
    public Bitmap? Image { get; private set; }

    /// <summary>The page's <c>lastModifiedTime</c> as at the last successful render.</summary>
    public DateTime? LastModified { get; set; }

    /// <summary>When this tile's image was last (re)rendered, for the "updated Ns ago" caption.</summary>
    public DateTime? LastRendered { get; set; }

    /// <summary>The outcome of the most recent poll.</summary>
    public PollStatus Status { get; set; } = PollStatus.NotStarted;

    /// <summary>A short human-readable reason for a non-<see cref="PollStatus.Ok"/> status.</summary>
    public string? StatusMessage { get; set; }

    /// <summary>Replaces the tile's image, disposing whatever bitmap it held before.</summary>
    public void SetImage(Bitmap? bitmap)
    {
        if (ReferenceEquals(Image, bitmap)) return;
        Image?.Dispose();
        Image = bitmap;
    }

    /// <summary>Releases the current image.</summary>
    public void Dispose() => Image?.Dispose();
}
