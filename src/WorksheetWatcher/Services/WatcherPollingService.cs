using System.Drawing;
using System.Threading;
using WorksheetWatcher.Models;

namespace WorksheetWatcher.Services;

/// <summary>One student's tile as refreshed by a poll cycle.</summary>
/// <param name="StudentId">The student this update is for.</param>
/// <param name="StudentName">The student's name.</param>
/// <param name="FocusImage">
/// The close-up of the latest change, or null when nothing new was rendered - e.g. a
/// status-only update such as "no page yet" or an error. The caller owns disposing the
/// previous image once it has stored this one.
/// </param>
/// <param name="PdfBytes">
/// The page's full PDF, so the full-screen view can render any part of any sheet itself,
/// sharply, at whatever size it needs. Null whenever <paramref name="FocusImage"/> is.
/// </param>
/// <param name="Focus">Which sheet, and where on it, <paramref name="FocusImage"/> shows.</param>
/// <param name="LastModified">The page's last-modified time as of this render, if any.</param>
/// <param name="Status">The outcome of this poll for this student.</param>
/// <param name="StatusMessage">A short reason for a non-<see cref="PollStatus.Ok"/> status.</param>
/// <param name="ChangeDetected">
/// True when the work visibly changed since the last look (or a page just appeared that was
/// missing before) - the signal the activity bar flashes on. False for the baseline first
/// render, which is not a change a student made.
/// </param>
public sealed record ThumbnailUpdate(
    string StudentId,
    string StudentName,
    Bitmap? FocusImage,
    byte[]? PdfBytes,
    PageFocus? Focus,
    DateTime? LastModified,
    PollStatus Status,
    string? StatusMessage,
    bool ChangeDetected);

/// <summary>
/// Owns a single OneNote COM client for the life of a watch session and polls one
/// notebook/worksheet on a dedicated background (STA) thread, invoking the update
/// callback given to the constructor for each student whose tile actually needs to change.
///
/// Every tick first asks OneNote to sync the notebook from the cloud - without this, a
/// student's edits sit on their own device (or in the cloud) and never reach the local
/// hierarchy XML this app reads, so <c>lastModifiedTime</c> simply never changes no matter
/// how often it is polled. Only after that does it re-resolve the worksheet (cheap: one
/// hierarchy XML call) and compare each student's page <c>lastModifiedTime</c> against
/// what was last seen. Only a page whose time moved pays for the heavier work: a PDF
/// export, a fingerprint comparison against the previous export to find where it changed
/// (see <see cref="PageChangeAnalyzer"/>), and a sharp close-up render of that spot. A
/// modified time that moved without anything visibly changing (OneNote touches pages when
/// it syncs) raises no update at all.
///
/// All OneNote calls happen sequentially on this one thread - the callback is invoked from
/// that same background thread, so callers updating WinForms controls must marshal back to
/// the UI thread themselves (<c>Control.Invoke</c>), the same discipline WheresTheWork's
/// report run uses.
/// </summary>
public sealed class WatcherPollingService : IDisposable
{
    /// <summary>What is remembered about each student between ticks.</summary>
    private sealed class StudentTrack
    {
        public DocumentSignature? Signature;
        public bool SeenWithoutPage;
    }

    private readonly AppConfig _config;
    private readonly Action<ThumbnailUpdate> _onUpdate;
    private readonly Dictionary<string, DateTime?> _lastSeenModified = new();
    private readonly Dictionary<string, StudentTrack> _tracks = new();

    private CancellationTokenSource? _cts;
    private Thread? _thread;
    private volatile bool _refreshRequested;

    public WatcherPollingService(AppConfig config, Action<ThumbnailUpdate> onUpdate)
    {
        _config = config;
        _onUpdate = onUpdate;
    }

    /// <summary>True while a watch session is running.</summary>
    public bool IsRunning => _thread is { IsAlive: true };

    /// <summary>
    /// Starts watching <paramref name="worksheetTitle"/> in <paramref name="notebookId"/>.
    /// Stops any session already running first. Runs until <see cref="Stop"/> or dispose.
    /// </summary>
    public void Start(string notebookId, string worksheetTitle)
    {
        Stop();

        _cts = new CancellationTokenSource();
        var token = _cts.Token;
        _lastSeenModified.Clear();
        _tracks.Clear();

        _thread = new Thread(() => Run(notebookId, worksheetTitle, token))
        {
            IsBackground = true,
            Name = "WorksheetWatcher-Poll"
        };
        // OneNote's automation object behaves as an STA COM server; give the polling
        // thread its own apartment rather than sharing the thread pool's MTA threads.
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
    }

    /// <summary>Stops the current watch session, if any, and waits for the thread to exit.</summary>
    public void Stop()
    {
        _cts?.Cancel();
        _thread?.Join(TimeSpan.FromSeconds(5));
        _cts?.Dispose();
        _cts = null;
        _thread = null;
    }

    /// <summary>Wakes the poll loop immediately instead of waiting out the current interval.</summary>
    public void RequestImmediateRefresh() => _refreshRequested = true;

    private void Run(string notebookId, string worksheetTitle, CancellationToken token)
    {
        using var com = new OneNoteComClient();
        try
        {
            com.Connect();
        }
        catch (OneNoteUnavailableException ex)
        {
            _onUpdate(new ThumbnailUpdate(string.Empty, string.Empty, null, null, null, null, PollStatus.Error, ex.Message, false));
            return;
        }

        var hierarchy = new OneNoteHierarchyService(com, _config);
        var resolver = new WorksheetResolutionService(hierarchy);
        var exporter = new PageExportService(com);

        while (!token.IsCancellationRequested)
        {
            try
            {
                // Best-effort: pulls fresh content from the cloud into OneNote's local
                // cache before we read anything. Without this, edits a student makes never
                // show up here no matter the poll interval - see the class remarks.
                com.SyncNode(notebookId);

                var targets = resolver.ResolveWorksheet(notebookId, worksheetTitle);
                foreach (var target in targets)
                {
                    token.ThrowIfCancellationRequested();
                    PollOne(target, exporter);
                }
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch (OneNoteContentException)
            {
                // A hierarchy-level failure (e.g. the notebook closed mid-session) - just
                // retry on the next tick rather than tearing down the whole watch session.
            }

            _refreshRequested = false;
            Wait(TimeSpan.FromSeconds(Math.Max(5, _config.PollIntervalSeconds)), token);
        }
    }

    private void PollOne(StudentPageTarget target, PageExportService exporter)
    {
        if (!_tracks.TryGetValue(target.StudentId, out var track))
            _tracks[target.StudentId] = track = new StudentTrack();

        if (!target.HasPage)
        {
            track.SeenWithoutPage = true;
            _onUpdate(new ThumbnailUpdate(target.StudentId, target.StudentName, null, null, null, null, PollStatus.NotStarted, "No page yet", false));
            return;
        }

        _lastSeenModified.TryGetValue(target.StudentId, out var previous);
        if (previous is not null && previous == target.LastModified)
            return; // unchanged since the last look - nothing to do

        try
        {
            var pdfBytes = exporter.ExportPdf(target.PageId!);
            using var pdf = OpenedPdf.Open(pdfBytes);
            if (pdf.PageCount == 0)
                throw new InvalidOperationException("OneNote exported an empty PDF.");

            var signature = PageChangeAnalyzer.Sign(pdf);
            var aspect = (double)_config.TileDisplayWidth / _config.TileDisplayHeight;

            var firstLook = track.Signature is null;
            var change = firstLook
                ? null
                : PageChangeAnalyzer.FindChange(track.Signature!, signature, _config.FocusWindowFraction, aspect);

            if (!firstLook && change is null)
            {
                // The modified time moved but nothing visible did - OneNote touches pages
                // when it syncs. Not worth redrawing, and certainly not worth a flash.
                track.Signature = signature;
                _lastSeenModified[target.StudentId] = target.LastModified;
                return;
            }

            var focus = change ?? PageChangeAnalyzer.InitialFocus(signature, _config.FocusWindowFraction, aspect);

            // Twice the on-screen size, so the tile stays crisp on high-DPI screens and the
            // full-screen view has detail to scale from.
            var closeUp = pdf.Render(focus.PageIndex, focus.Window, _config.TileDisplayWidth * 2, _config.TileDisplayHeight * 2);

            var flash = change is not null || (firstLook && track.SeenWithoutPage);

            // Only now is this look "banked": had the render above failed, the previous
            // fingerprint and modified time stay put so the next tick tries again.
            track.Signature = signature;
            track.SeenWithoutPage = false;
            _lastSeenModified[target.StudentId] = target.LastModified;
            _onUpdate(new ThumbnailUpdate(target.StudentId, target.StudentName, closeUp, pdfBytes, focus, target.LastModified, PollStatus.Ok, null, flash));
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (OneNoteContentException ex)
        {
            _onUpdate(new ThumbnailUpdate(target.StudentId, target.StudentName, null, null, null, target.LastModified, PollStatus.Locked, ex.Message, false));
        }
        catch (Exception ex)
        {
            // Anything else (a PDF that will not open, a GDI+ failure) must never escape:
            // an unhandled exception on this background thread would take the whole app down.
            _onUpdate(new ThumbnailUpdate(target.StudentId, target.StudentName, null, null, null, target.LastModified, PollStatus.Error, $"Could not render: {ex.Message}", false));
        }
    }

    /// <summary>Sleeps in short steps so cancellation and manual refresh requests are responsive.</summary>
    private void Wait(TimeSpan delay, CancellationToken token)
    {
        var elapsed = TimeSpan.Zero;
        var step = TimeSpan.FromMilliseconds(250);
        while (elapsed < delay && !token.IsCancellationRequested && !_refreshRequested)
        {
            Thread.Sleep(step);
            elapsed += step;
        }
    }

    public void Dispose() => Stop();
}
