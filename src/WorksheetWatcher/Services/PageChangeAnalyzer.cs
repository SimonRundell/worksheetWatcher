using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace WorksheetWatcher.Services;

/// <summary>
/// A compact greyscale fingerprint of one sheet: the sheet rendered at a fixed modest
/// size, then averaged down in small square blocks. Comparing two of these shows where a
/// sheet changed, at a tiny fraction of the memory a full bitmap would cost per student.
/// </summary>
/// <param name="GridWidth">Blocks across.</param>
/// <param name="GridHeight">Blocks down.</param>
/// <param name="Gray">Average brightness (0-255) of each block, row by row.</param>
/// <param name="PageSize">The sheet's size in the PDF's device-independent units.</param>
public sealed record PageSignature(int GridWidth, int GridHeight, byte[] Gray, SizeF PageSize);

/// <summary>Fingerprints for every sheet of a page, in order.</summary>
public sealed record DocumentSignature(IReadOnlyList<PageSignature> Pages);

/// <summary>
/// The part of a page a tile should show: which sheet, and the window on it, in the PDF's
/// device-independent units. The window always has the tile's shape.
/// </summary>
public sealed record PageFocus(int PageIndex, RectangleF Window);

/// <summary>
/// Works out where on a student's page the latest change happened, so the tile can show a
/// sharp close-up of it instead of a whole shrunken page.
///
/// How: each export is fingerprinted (see <see cref="PageSignature"/>); the new fingerprint
/// is compared with the previous one; the first run of changed rows becomes the focus.
/// "First" rather than "biggest" is deliberate. Typing at the end of the work changes only
/// the new text, so first is simply right. Typing in the middle pushes everything below it
/// down, which changes all of that too - the first changed row is the real edit and the rest
/// is just displaced text. The run is capped at a fraction of the window height for the
/// same reason, so displaced text cannot stretch the focus over the whole page.
/// </summary>
public static class PageChangeAnalyzer
{
    /// <summary>Width, in pixels, each sheet is rendered at for fingerprinting.</summary>
    public const int AnalysisWidth = 1400;

    private const int Block = 4;               // fingerprint block size, in pixels
    private const int DiffThreshold = 24;      // brightness difference that counts as changed
    private const int MinChangedBlocks = 2;    // fewer than this on a sheet is treated as noise
    private const int RunGapBlocks = 10;       // unchanged rows tolerated inside one run (about a line)
    private const double RunCapOfWindow = 0.6; // longest run, as a fraction of the window height
    private const double FooterBand = 0.05;    // bottom of every sheet is a fixed "Page N" footer
    private const float Padding = 12;          // breathing room under the last content, in DIPs

    /// <summary>Renders and fingerprints every sheet of <paramref name="pdf"/>.</summary>
    public static DocumentSignature Sign(OpenedPdf pdf)
    {
        var pages = new List<PageSignature>();
        for (var i = 0; i < pdf.PageCount; i++)
        {
            var size = pdf.PageSize(i);
            var height = (int)Math.Round(AnalysisWidth * size.Height / size.Width);
            using var bitmap = pdf.Render(i, null, AnalysisWidth, height);
            pages.Add(SignBitmap(bitmap, size));
        }
        return new DocumentSignature(pages);
    }

    /// <summary>Fingerprints one already-rendered sheet.</summary>
    public static PageSignature SignBitmap(Bitmap bitmap, SizeF pageSize)
    {
        var data = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height),
            ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);

        byte[] pixels;
        int stride;
        try
        {
            stride = data.Stride;
            pixels = new byte[Math.Abs(stride) * bitmap.Height];
            Marshal.Copy(data.Scan0, pixels, 0, pixels.Length);
        }
        finally
        {
            bitmap.UnlockBits(data);
        }

        var gridWidth = bitmap.Width / Block;
        var gridHeight = bitmap.Height / Block;
        var gray = new byte[gridWidth * gridHeight];

        for (var by = 0; by < gridHeight; by++)
        {
            for (var bx = 0; bx < gridWidth; bx++)
            {
                var sum = 0;
                for (var y = 0; y < Block; y++)
                {
                    var i = (by * Block + y) * stride + bx * Block * 4;
                    for (var x = 0; x < Block; x++, i += 4)
                        sum += (pixels[i + 2] * 77 + pixels[i + 1] * 151 + pixels[i] * 28) >> 8;
                }
                gray[by * gridWidth + bx] = (byte)(sum / (Block * Block));
            }
        }

        return new PageSignature(gridWidth, gridHeight, gray, pageSize);
    }

    /// <summary>
    /// Compares two fingerprints and returns where the first visible change is, or null if
    /// nothing visibly changed. <paramref name="aspect"/> is the tile's width / height.
    /// </summary>
    public static PageFocus? FindChange(DocumentSignature previous, DocumentSignature current, double windowFraction, double aspect)
    {
        for (var p = 0; p < current.Pages.Count; p++)
        {
            var now = current.Pages[p];
            var before = p < previous.Pages.Count ? previous.Pages[p] : null;

            var changed = ChangedBlocks(before, now, out var count);
            if (count < MinChangedBlocks) continue;

            return BuildFocus(p, now, changed, windowFraction, aspect);
        }

        return null;
    }

    /// <summary>
    /// Where to look before any change has been seen: the end of the work, since that is
    /// where a student is most likely to be up to. The fixed page footer is ignored.
    /// </summary>
    public static PageFocus InitialFocus(DocumentSignature signature, double windowFraction, double aspect)
    {
        for (var p = signature.Pages.Count - 1; p >= 0; p--)
        {
            var page = signature.Pages[p];
            if (ContentBounds(page) is not { } bounds) continue;

            var sx = page.PageSize.Width / page.GridWidth;
            var sy = page.PageSize.Height / page.GridHeight;
            var windowWidth = (float)(windowFraction * page.PageSize.Width);
            var windowHeight = windowWidth / (float)aspect;

            // Start at the left edge of the work so lines read from their beginning; with no
            // change to centre on, cutting the start of every line off would just be confusing.
            var centreX = bounds.Left * sx - Padding + windowWidth / 2;
            var bottom = (bounds.Bottom + 1) * sy + Padding;
            return new PageFocus(p, FitWindow(page.PageSize, centreX, bottom - windowHeight / 2, 0, 0, windowFraction, aspect));
        }

        var first = signature.Pages.Count > 0 ? signature.Pages[0].PageSize : new SizeF(794, 1124);
        return new PageFocus(0, FitWindow(first, first.Width / 2, 0, 0, 0, windowFraction, aspect));
    }

    private static PageFocus BuildFocus(int pageIndex, PageSignature page, bool[] changed, double windowFraction, double aspect)
    {
        var sx = page.PageSize.Width / page.GridWidth;
        var sy = page.PageSize.Height / page.GridHeight;
        var windowHeight = windowFraction * page.PageSize.Width / aspect;
        var maxRunRows = Math.Max(RunGapBlocks, (int)(RunCapOfWindow * windowHeight / sy));

        var rowHasChange = new bool[page.GridHeight];
        for (var y = 0; y < page.GridHeight; y++)
            for (var x = 0; x < page.GridWidth; x++)
                if (changed[y * page.GridWidth + x]) { rowHasChange[y] = true; break; }

        var top = Array.IndexOf(rowHasChange, true);

        var bottom = top;
        for (var y = top; y < page.GridHeight && y - top < maxRunRows && y - bottom <= RunGapBlocks; y++)
            if (rowHasChange[y]) bottom = y;

        int left = page.GridWidth, right = 0;
        for (var y = top; y <= bottom; y++)
        {
            for (var x = 0; x < page.GridWidth; x++)
            {
                if (!changed[y * page.GridWidth + x]) continue;
                if (x < left) left = x;
                if (x > right) right = x;
            }
        }

        var changeWidth = (right - left + 1) * sx;
        var changeHeight = (bottom - top + 1) * sy;
        var centreX = (left + right + 1) * 0.5f * sx;
        var centreY = (top + bottom + 1) * 0.5f * sy;

        return new PageFocus(pageIndex, FitWindow(page.PageSize, centreX, centreY, changeWidth, changeHeight, windowFraction, aspect));
    }

    /// <summary>
    /// A window of the tile's shape centred on a point, wide enough for the usual fraction
    /// of the page - or for the change itself if that is bigger - and kept inside the sheet.
    /// </summary>
    private static RectangleF FitWindow(SizeF page, float centreX, float centreY, float needWidth, float needHeight, double fraction, double aspect)
    {
        var width = (float)(fraction * page.Width);
        width = Math.Max(width, needWidth * 1.25f);
        width = Math.Max(width, needHeight * 1.25f * (float)aspect);
        width = Math.Min(width, page.Width);

        var height = width / (float)aspect;
        if (height > page.Height)
        {
            height = page.Height;
            width = height * (float)aspect;
        }

        var x = Math.Clamp(centreX - width / 2, 0, page.Width - width);
        var y = Math.Clamp(centreY - height / 2, 0, page.Height - height);
        return new RectangleF(x, y, width, height);
    }

    /// <summary>Marks the blocks whose brightness moved by more than the threshold.</summary>
    private static bool[] ChangedBlocks(PageSignature? before, PageSignature now, out int count)
    {
        var changed = new bool[now.Gray.Length];
        count = 0;

        // A sheet that did not exist before (the page just grew onto it) counts as changed
        // wherever it has content; so does a sheet whose size differs, since the grids
        // would not line up.
        var comparable = before is not null
            && before.GridWidth == now.GridWidth
            && before.GridHeight == now.GridHeight;

        var background = Background(now.Gray);
        for (var i = 0; i < now.Gray.Length; i++)
        {
            var differs = comparable
                ? Math.Abs(now.Gray[i] - before!.Gray[i]) >= DiffThreshold
                : Math.Abs(now.Gray[i] - background) >= DiffThreshold;

            if (!differs) continue;
            changed[i] = true;
            count++;
        }

        return changed;
    }

    /// <summary>Block bounds of everything that is not background, ignoring the footer band; null if blank.</summary>
    private static Rectangle? ContentBounds(PageSignature page)
    {
        var background = Background(page.Gray);
        var lastRow = (int)(page.GridHeight * (1 - FooterBand));

        int left = page.GridWidth, top = page.GridHeight, right = -1, bottom = -1;
        for (var y = 0; y < lastRow; y++)
        {
            for (var x = 0; x < page.GridWidth; x++)
            {
                if (Math.Abs(page.Gray[y * page.GridWidth + x] - background) < DiffThreshold) continue;
                if (x < left) left = x;
                if (x > right) right = x;
                if (y < top) top = y;
                if (y > bottom) bottom = y;
            }
        }

        return right < 0 ? null : Rectangle.FromLTRB(left, top, right, bottom);
    }

    /// <summary>The most common brightness - the page colour.</summary>
    private static byte Background(byte[] gray)
    {
        var histogram = new int[256];
        foreach (var g in gray) histogram[g]++;

        var best = 0;
        for (var i = 1; i < 256; i++)
            if (histogram[i] > histogram[best]) best = i;
        return (byte)best;
    }
}
