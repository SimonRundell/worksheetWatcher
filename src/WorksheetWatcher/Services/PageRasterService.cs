using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using PublishFormat = Microsoft.Office.Interop.OneNote.PublishFormat;

namespace WorksheetWatcher.Services;

/// <summary>
/// Rasterises a single OneNote page into a fixed-size <see cref="Bitmap"/> by asking
/// OneNote to <c>Publish</c> it as an Enhanced Metafile (a vector snapshot, ink included)
/// and drawing that into a bitmap sized for the thumbnail grid.
///
/// This is a per-page export, cheaper than WheresTheWork's whole-notebook content walk,
/// and is only called for pages the poller has determined actually changed - see
/// <see cref="WatcherPollingService"/>.
/// </summary>
public sealed class PageRasterService
{
    private readonly OneNoteComClient _com;

    public PageRasterService(OneNoteComClient com) => _com = com;

    /// <summary>
    /// Publishes <paramref name="pageId"/> to a temp EMF file, loads it, and draws it
    /// aspect-fit (letterboxed on white) into a new <paramref name="width"/> x
    /// <paramref name="height"/> bitmap. The temp file is deleted before returning.
    /// Throws <see cref="OneNoteContentException"/> if OneNote refuses the page.
    /// </summary>
    public Bitmap RenderPage(string pageId, int width, int height)
    {
        var tempPath = Path.Combine(Path.GetTempPath(), $"WorksheetWatcher-{Guid.NewGuid():N}.emf");
        try
        {
            _com.PublishPageToFile(pageId, tempPath, PublishFormat.pfEMF);

            using var metafile = new Metafile(tempPath);
            return DrawToBitmap(metafile, width, height);
        }
        finally
        {
            try { if (File.Exists(tempPath)) File.Delete(tempPath); }
            catch { /* best effort - a stray temp file is not worth failing the render over */ }
        }
    }

    private static Bitmap DrawToBitmap(Metafile metafile, int width, int height)
    {
        var bitmap = new Bitmap(width, height);
        using var g = Graphics.FromImage(bitmap);
        g.Clear(Color.White);
        g.SmoothingMode = SmoothingMode.HighQuality;
        g.InterpolationMode = InterpolationMode.HighQualityBicubic;

        var srcSize = metafile.Size;
        if (srcSize.Width <= 0 || srcSize.Height <= 0)
            return bitmap;

        // Aspect-fit, letterboxed on the white background - the metafile is vector, so
        // this scales cleanly at any thumbnail size.
        var scale = Math.Min((double)width / srcSize.Width, (double)height / srcSize.Height);
        var drawWidth = (int)Math.Round(srcSize.Width * scale);
        var drawHeight = (int)Math.Round(srcSize.Height * scale);
        var x = (width - drawWidth) / 2;
        var y = (height - drawHeight) / 2;

        g.DrawImage(metafile, new Rectangle(x, y, drawWidth, drawHeight));
        return bitmap;
    }
}
