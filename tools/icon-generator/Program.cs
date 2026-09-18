using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;

// Generates appicon.ico for "Where's the Work" - a magnifying glass over a marked page.
// Usage: dotnet run -- <output-path>

var outPath = args.Length > 0 ? args[0] : "appicon.ico";
int[] sizes = { 16, 24, 32, 48, 64, 128, 256 };

// Every frame as an uncompressed 32bpp DIB: larger file, but parsed identically by the
// .NET SDK icon embedder, Windows shell, and older GDI+ - no PNG-in-ICO edge cases.
var entries = sizes.Select(s =>
{
    using var bmp = Render(s);
    return (size: s, data: ToDib(bmp), png: false);
}).ToArray();

WriteIco(outPath, entries);
Console.WriteLine($"Wrote {outPath} with sizes: {string.Join(", ", sizes)}");

// ---------------------------------------------------------------------------

static Bitmap Render(int size)
{
    var bmp = new Bitmap(size, size, PixelFormat.Format32bppArgb);
    using var g = Graphics.FromImage(bmp);
    g.SmoothingMode = SmoothingMode.AntiAlias;
    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
    g.PixelOffsetMode = PixelOffsetMode.HighQuality;
    g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
    g.Clear(Color.Transparent);

    // Work in a 256 x 256 design space, scaled to the target size.
    float scale = size / 256f;
    g.ScaleTransform(scale, scale);

    // Rounded background with a vertical blue gradient.
    var bg = new RectangleF(6, 6, 244, 244);
    using (var path = RoundedRect(bg, 46))
    using (var brush = new LinearGradientBrush(bg, Color.FromArgb(0x33, 0x8F, 0xD4), Color.FromArgb(0x14, 0x4E, 0x86), LinearGradientMode.Vertical))
    {
        g.FillPath(brush, path);
        using var rim = new Pen(Color.FromArgb(60, 255, 255, 255), 2f);
        g.DrawPath(rim, path);
    }

    // The page, tilted slightly, with a soft shadow.
    var pageRect = new RectangleF(52, 40, 118, 156);
    var state = g.Save();
    g.TranslateTransform(pageRect.Left + pageRect.Width / 2, pageRect.Top + pageRect.Height / 2);
    g.RotateTransform(-7f);
    g.TranslateTransform(-(pageRect.Left + pageRect.Width / 2), -(pageRect.Top + pageRect.Height / 2));

    var shadow = pageRect;
    shadow.Offset(6, 8);
    using (var sp = RoundedRect(shadow, 12))
    using (var sb = new SolidBrush(Color.FromArgb(70, 0, 0, 0)))
        g.FillPath(sb, sp);

    using (var pp = RoundedRect(pageRect, 12))
    {
        g.FillPath(Brushes.White, pp);
        using var edge = new Pen(Color.FromArgb(30, 0, 0, 0), 1.5f);
        g.DrawPath(edge, pp);
    }

    // Ruled lines (dropped at tiny sizes where they only add noise) and a bold green tick.
    bool tiny = size <= 24;
    if (!tiny)
    {
        using var line = new Pen(Color.FromArgb(0xC2, 0xCC, 0xD6), 6f) { StartCap = LineCap.Round, EndCap = LineCap.Round };
        for (int i = 0; i < 4; i++)
        {
            float y = pageRect.Top + 40 + i * 26;
            g.DrawLine(line, pageRect.Left + 18, y, pageRect.Right - 18, y);
        }
    }
    using (var tick = new Pen(Color.FromArgb(0x2E, 0xA0, 0x43), tiny ? 20f : 13f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
    {
        float bx = pageRect.Left + (tiny ? 20 : 26), by = pageRect.Top + (tiny ? 66 : 34);
        float u = tiny ? 22 : 16, r = tiny ? 30 : 44, up = tiny ? 26 : 20;
        g.DrawLines(tick, new[] { new PointF(bx, by), new PointF(bx + u, by + u), new PointF(bx + r, by - up) });
    }
    g.Restore(state);

    // Magnifying glass over the lower-right of the page. A touch bigger and bolder at
    // tiny sizes so it stays the recognisable shape.
    var lens = tiny ? new RectangleF(118, 114, 104, 104) : new RectangleF(120, 116, 92, 92);
    using (var handle = new Pen(Color.FromArgb(0x1B, 0x28, 0x32), tiny ? 34f : 26f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
        g.DrawLine(handle, lens.Right - 18, lens.Bottom - 18, lens.Right + 20, lens.Bottom + 20);
    using (var glass = new SolidBrush(Color.FromArgb(tiny ? 70 : 90, 0x9A, 0xD9, 0xF0)))
        g.FillEllipse(glass, lens);
    using (var ring = new Pen(Color.FromArgb(0x1B, 0x28, 0x32), tiny ? 28f : 16f))
        g.DrawEllipse(ring, lens);
    if (!tiny)
        using (var glint = new Pen(Color.FromArgb(160, 255, 255, 255), 7f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
            g.DrawArc(glint, lens.Left + 14, lens.Top + 14, lens.Width - 28, lens.Height - 28, 160, 70);

    return bmp;
}

static GraphicsPath RoundedRect(RectangleF r, float radius)
{
    float d = radius * 2;
    var p = new GraphicsPath();
    p.AddArc(r.Left, r.Top, d, d, 180, 90);
    p.AddArc(r.Right - d, r.Top, d, d, 270, 90);
    p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
    p.AddArc(r.Left, r.Bottom - d, d, d, 90, 90);
    p.CloseFigure();
    return p;
}

// Uncompressed 32bpp DIB (BITMAPINFOHEADER + BGRA XOR bitmap + 1bpp AND mask), bottom-up.
static byte[] ToDib(Bitmap bmp)
{
    int w = bmp.Width, h = bmp.Height;
    using var ms = new MemoryStream();
    using var bw = new BinaryWriter(ms);

    bw.Write(40);            // biSize
    bw.Write(w);             // biWidth
    bw.Write(h * 2);         // biHeight (XOR + AND)
    bw.Write((short)1);      // biPlanes
    bw.Write((short)32);     // biBitCount
    bw.Write(0);             // biCompression BI_RGB
    bw.Write(0);             // biSizeImage
    bw.Write(0); bw.Write(0); bw.Write(0); bw.Write(0);

    for (int y = h - 1; y >= 0; y--)
        for (int x = 0; x < w; x++)
        {
            var c = bmp.GetPixel(x, y);
            bw.Write(c.B); bw.Write(c.G); bw.Write(c.R); bw.Write(c.A);
        }

    int maskStride = ((w + 31) / 32) * 4;
    var maskRow = new byte[maskStride];
    for (int y = h - 1; y >= 0; y--)
    {
        Array.Clear(maskRow);
        for (int x = 0; x < w; x++)
            if (bmp.GetPixel(x, y).A < 128)
                maskRow[x / 8] |= (byte)(0x80 >> (x % 8)); // 1 = transparent
        bw.Write(maskRow);
    }

    return ms.ToArray();
}

static void WriteIco(string path, (int size, byte[] data, bool png)[] imgs)
{
    using var fs = File.Create(path);
    using var bw = new BinaryWriter(fs);

    bw.Write((short)0);            // reserved
    bw.Write((short)1);            // type: icon
    bw.Write((short)imgs.Length);

    int offset = 6 + imgs.Length * 16;
    foreach (var (size, data, _) in imgs)
    {
        bw.Write((byte)(size >= 256 ? 0 : size)); // width
        bw.Write((byte)(size >= 256 ? 0 : size)); // height
        bw.Write((byte)0);                        // palette
        bw.Write((byte)0);                        // reserved
        bw.Write((short)1);                       // planes
        bw.Write((short)32);                      // bpp
        bw.Write(data.Length);
        bw.Write(offset);
        offset += data.Length;
    }

    foreach (var (_, data, _) in imgs)
        bw.Write(data);
}
