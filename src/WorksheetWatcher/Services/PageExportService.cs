using System.IO;
using PublishFormat = Microsoft.Office.Interop.OneNote.PublishFormat;

namespace WorksheetWatcher.Services;

/// <summary>
/// Asks OneNote to export a single page as a PDF and hands back the bytes.
///
/// PDF rather than EMF because OneNote's single-page EMF export contains only the first
/// A4 sheet of the page - anything a student writes below that is silently cut off - while
/// the PDF export contains every sheet. Measured on a real class notebook, a PDF export
/// takes roughly 0.1 - 0.7s and is a few hundred KB at most.
/// </summary>
public sealed class PageExportService
{
    private readonly OneNoteComClient _com;

    public PageExportService(OneNoteComClient com) => _com = com;

    /// <summary>
    /// Publishes <paramref name="pageId"/> to a temporary PDF, reads it into memory, and
    /// deletes the file. Throws <see cref="OneNoteContentException"/> if OneNote refuses
    /// the page (for example a locked section).
    /// </summary>
    public byte[] ExportPdf(string pageId)
    {
        var tempPath = Path.Combine(Path.GetTempPath(), $"WorksheetWatcher-{Guid.NewGuid():N}.pdf");
        try
        {
            _com.PublishPageToFile(pageId, tempPath, PublishFormat.pfPDF);
            return File.ReadAllBytes(tempPath);
        }
        finally
        {
            try { if (File.Exists(tempPath)) File.Delete(tempPath); }
            catch { /* best effort - a stray temp file is not worth failing the export over */ }
        }
    }
}
