using System.Runtime.InteropServices;
using System.Threading;
using Interop = Microsoft.Office.Interop.OneNote;
using HierarchyScope = Microsoft.Office.Interop.OneNote.HierarchyScope;
using PageInfo = Microsoft.Office.Interop.OneNote.PageInfo;
using PublishFormat = Microsoft.Office.Interop.OneNote.PublishFormat;

namespace WorksheetWatcher.Services;

/// <summary>
/// Wrapper around the classic desktop OneNote automation object
/// (<c>Microsoft.Office.Interop.OneNote.Application</c>).
///
/// The interop is a pre-generated, embedded assembly (see <c>libs/</c> and the project
/// file), not late binding - see the sibling WheresTheWork project's README for why: on
/// Click-to-Run OneNote the type library is registered only under Win32 and OneNote's own
/// <c>IDispatch::GetTypeInfo</c> / <c>GetIDsOfNames</c> return errors, so <c>dynamic</c>
/// and <c>Type.InvokeMember</c> both fail with <c>0x8002801D</c> / <c>E_FAIL</c>. A
/// generated interop dispatches straight down the interface vtable and never touches the
/// type library at run time.
/// </summary>
public sealed class OneNoteComClient : IDisposable
{
    // HRESULTs OneNote/DCOM raise transiently while OneNote is busy, still starting, or
    // servicing another call - including a second activation request arriving while one is
    // already in flight, which this app can trigger (the notebook picker, the worksheet
    // lookup, and the watch session each activate their own Application COM object, from
    // different threads, and can overlap). Worth a short retry rather than a hard failure.
    private const int RpcServerCallRetryLater = unchecked((int)0x8001010A); // RPC_E_SERVERCALL_RETRYLATER
    private const int RpcCallRejected = unchecked((int)0x80010001);         // RPC_E_CALL_REJECTED
    private const int RpcServerUnavailable = unchecked((int)0x800706BA);    // RPC_S_SERVER_UNAVAILABLE
    private const int RpcCallFailedDidNotExecute = unchecked((int)0x800706BF); // RPC_S_CALL_FAILED_DNE

    private static bool IsTransient(int hresult) =>
        hresult is RpcServerCallRetryLater or RpcCallRejected or RpcServerUnavailable or RpcCallFailedDidNotExecute;

    private Interop.Application? _app;

    /// <summary>True once <see cref="Connect"/> has succeeded.</summary>
    public bool IsConnected => _app is not null;

    /// <summary>
    /// Instantiates the OneNote automation object, retrying transient COM activation
    /// failures a few times first. Throws <see cref="OneNoteUnavailableException"/> with a
    /// user-facing message if OneNote is not installed, or still cannot be reached after
    /// retrying.
    /// </summary>
    public void Connect()
    {
        if (_app is not null) return;

        const int maxAttempts = 4;

        for (var attempt = 1; ; attempt++)
        {
            try
            {
                _app = new Interop.Application();
                return;
            }
            catch (COMException com) when (IsTransient(com.HResult) && attempt < maxAttempts)
            {
                Thread.Sleep(250 * attempt);
            }
            catch (Exception ex) when (ex is COMException or InvalidCastException or MemberAccessException or TypeInitializationException or FileNotFoundException)
            {
                throw new OneNoteUnavailableException(
                    "OneNote could not be reached. Make sure the classic desktop version of " +
                    "OneNote is installed and open with your Class Notebook loaded, then try again.", ex);
            }
        }
    }

    /// <summary>Returns the notebook-scope hierarchy XML: every currently open notebook.</summary>
    public string GetNotebooksXml() => GetHierarchy(null, HierarchyScope.hsNotebooks);

    /// <summary>
    /// Returns the full page-scope hierarchy XML rooted at <paramref name="notebookId"/>:
    /// section groups, sections and pages in one call.
    /// </summary>
    public string GetNotebookTreeXml(string notebookId) => GetHierarchy(notebookId, HierarchyScope.hsPages);

    /// <summary>
    /// Returns a single page's content XML, including embedded image binary data (and OCR
    /// text where OneNote exposes it) according to <paramref name="pageInfoFlag"/>.
    /// </summary>
    /// <param name="pageId">Page hierarchy ID.</param>
    /// <param name="pageInfoFlag">
    /// One of the <c>piXxx</c> names (<c>piBasic</c>, <c>piBinaryData</c>, <c>piAll</c>, ...).
    /// Unknown values fall back to <c>piBinaryData</c>.
    /// </param>
    public string GetPageContentXml(string pageId, string pageInfoFlag)
    {
        var flag = ParsePageInfo(pageInfoFlag);
        EnsureConnected();

        return Invoke("GetPageContent", () =>
        {
            _app!.GetPageContent(pageId, out var xml, flag);
            return xml ?? string.Empty;
        });
    }

    /// <summary>Maps a config string to the <see cref="PageInfo"/> enum, defaulting to <c>piBinaryData</c>.</summary>
    public static PageInfo ParsePageInfo(string? name) =>
        Enum.TryParse<PageInfo>(name, ignoreCase: true, out var value) ? value : PageInfo.piBinaryData;

    /// <summary>
    /// Exports a single page to <paramref name="targetFilePath"/> in the given
    /// <paramref name="format"/> - <see cref="PublishFormat.pfEMF"/> (the default) writes
    /// an Enhanced Metafile: a vector snapshot of exactly what OneNote renders for that
    /// page, ink included, which is what the thumbnail rasteriser draws from. Unlike
    /// <see cref="GetPageContentXml"/>, <c>Publish</c> needs a real file path - it cannot
    /// write to a stream - so the caller owns cleaning up the file afterwards.
    /// </summary>
    public void PublishPageToFile(string pageId, string targetFilePath, PublishFormat format = PublishFormat.pfEMF)
    {
        EnsureConnected();
        InvokeAction("Publish", () => _app!.Publish(pageId, targetFilePath, format, string.Empty));
    }

    /// <summary>
    /// Asks OneNote to sync a hierarchy node (a whole notebook when given its ID). Best
    /// effort - failures are swallowed, since a poll can still run against whatever is
    /// already local.
    /// </summary>
    public void SyncNode(string hierarchyId)
    {
        if (_app is null) return;
        try
        {
            _app.SyncHierarchy(hierarchyId);
        }
        catch (COMException)
        {
            // best effort
        }
    }

    private string GetHierarchy(string? startNodeId, HierarchyScope scope)
    {
        EnsureConnected();

        return Invoke("GetHierarchy", () =>
        {
            _app!.GetHierarchy(startNodeId, scope, out var xml);
            return xml ?? string.Empty;
        });
    }

    /// <summary>
    /// Runs a single COM call that returns a value. Retries the "OneNote is busy" HRESULTs
    /// a few times and otherwise wraps the raw <see cref="COMException"/> in an
    /// <see cref="OneNoteContentException"/> so callers can isolate per-page failures
    /// (locked sections, deleted pages) without a crash.
    /// </summary>
    private static T Invoke<T>(string method, Func<T> call)
    {
        const int maxAttempts = 4;

        for (var attempt = 1; ; attempt++)
        {
            try
            {
                return call();
            }
            catch (COMException com)
            {
                if (IsTransient(com.HResult) && attempt < maxAttempts)
                {
                    Thread.Sleep(250 * attempt);
                    continue;
                }

                throw new OneNoteContentException($"OneNote rejected the '{method}' call: {com.Message}", com);
            }
        }
    }

    /// <summary>Same retry/wrap behaviour as <see cref="Invoke{T}"/>, for a void-returning call.</summary>
    private static void InvokeAction(string method, Action call) =>
        Invoke(method, () => { call(); return true; });

    private void EnsureConnected()
    {
        if (_app is null)
            throw new InvalidOperationException("Connect() must be called before using the OneNote client.");
    }

    /// <summary>Releases the underlying COM object.</summary>
    public void Dispose()
    {
        if (_app is not null)
        {
            try { Marshal.FinalReleaseComObject(_app); } catch { /* best effort */ }
            _app = null;
        }
        GC.SuppressFinalize(this);
    }
}

/// <summary>Raised when OneNote cannot be reached at all (not running / not installed).</summary>
public sealed class OneNoteUnavailableException : Exception
{
    public OneNoteUnavailableException(string message) : base(message) { }
    public OneNoteUnavailableException(string message, Exception inner) : base(message, inner) { }
}

/// <summary>
/// Raised when a specific hierarchy, page or publish call fails (for example a
/// password-protected section). Callers processing many pages should catch this per page
/// and continue.
/// </summary>
public sealed class OneNoteContentException : Exception
{
    public OneNoteContentException(string message, Exception inner) : base(message, inner) { }
}
