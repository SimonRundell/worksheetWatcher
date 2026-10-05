using System.Drawing;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Data.Pdf;
using Windows.Storage.Streams;

namespace WorksheetWatcher.Services;

/// <summary>
/// A PDF held in memory and rendered with Windows' own built-in PDF engine
/// (<c>Windows.Data.Pdf</c>), so nothing extra ships with the app.
///
/// Rendering is the reason PDF is used rather than OneNote's single-page EMF export: the
/// EMF only ever contains the first A4 sheet of a page, silently dropping everything below
/// it, whereas the PDF holds every sheet. And because a PDF is vector, any rectangle of a
/// sheet can be rendered at whatever pixel size is wanted - which is how the tile close-ups
/// stay sharp without rendering whole pages at huge sizes.
///
/// Positions and sizes are in the PDF's own device-independent units (A4 is about
/// 794 x 1124). A <see cref="Windows.Data.Pdf.PdfDocument"/> allows one open page at a time,
/// so calls on one instance must not overlap; use a separate instance per thread.
/// </summary>
public sealed class OpenedPdf : IDisposable
{
    private readonly InMemoryRandomAccessStream _stream;
    private readonly PdfDocument _document;

    private OpenedPdf(InMemoryRandomAccessStream stream, PdfDocument document)
    {
        _stream = stream;
        _document = document;
    }

    /// <summary>Number of sheets in the PDF.</summary>
    public int PageCount => (int)_document.PageCount;

    /// <summary>Opens a PDF from its bytes.</summary>
    public static async Task<OpenedPdf> OpenAsync(byte[] pdf)
    {
        var stream = new InMemoryRandomAccessStream();
        await stream.WriteAsync(pdf.AsBuffer());
        stream.Seek(0);
        var document = await PdfDocument.LoadFromStreamAsync(stream);
        return new OpenedPdf(stream, document);
    }

    /// <summary>Synchronous <see cref="OpenAsync"/>, for the background polling thread (never the UI thread).</summary>
    public static OpenedPdf Open(byte[] pdf) => OpenAsync(pdf).GetAwaiter().GetResult();

    /// <summary>The size of a sheet in the PDF's device-independent units.</summary>
    public SizeF PageSize(int pageIndex)
    {
        using var page = _document.GetPage((uint)pageIndex);
        return new SizeF((float)page.Size.Width, (float)page.Size.Height);
    }

    /// <summary>
    /// Renders a sheet, or just <paramref name="source"/> of it, to a bitmap of exactly
    /// <paramref name="width"/> x <paramref name="height"/> pixels. Keep the destination's
    /// shape the same as the source rectangle's or the result is stretched.
    /// </summary>
    public async Task<Bitmap> RenderAsync(int pageIndex, RectangleF? source, int width, int height)
    {
        using var page = _document.GetPage((uint)pageIndex);

        var options = new PdfPageRenderOptions
        {
            DestinationWidth = (uint)Math.Max(1, width),
            DestinationHeight = (uint)Math.Max(1, height),
            BackgroundColor = Windows.UI.Color.FromArgb(255, 255, 255, 255)
        };
        if (source is { } r)
            options.SourceRect = new Windows.Foundation.Rect(r.X, r.Y, r.Width, r.Height);

        using var rendered = new InMemoryRandomAccessStream();
        await page.RenderToStreamAsync(rendered, options);

        using var decoded = new Bitmap(rendered.GetInputStreamAt(0).AsStreamForRead());
        return new Bitmap(decoded); // a copy, so it no longer depends on the stream
    }

    /// <summary>Synchronous <see cref="RenderAsync"/>, for the background polling thread (never the UI thread).</summary>
    public Bitmap Render(int pageIndex, RectangleF? source, int width, int height) =>
        RenderAsync(pageIndex, source, width, height).GetAwaiter().GetResult();

    /// <summary>Releases the in-memory stream.</summary>
    public void Dispose() => _stream.Dispose();
}
