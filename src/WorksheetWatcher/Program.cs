using WorksheetWatcher.Forms;
using WorksheetWatcher.Services;

namespace WorksheetWatcher;

/// <summary>Application entry point.</summary>
internal static class Program
{
    /// <summary>
    /// Standard WinForms bootstrap. STA is required for both WinForms and OneNote COM
    /// automation.
    /// </summary>
    /// <param name="args">
    /// <c>--selftest [notebook]</c> - headless connectivity check: lists open notebooks,
    /// walks one, and lists its students and page count. Takes an optional substring of a
    /// notebook's name or OneNote nickname.
    /// <c>--rastertest &lt;notebook&gt; &lt;pageTitle&gt; [student]</c> - rasterises one
    /// student's copy of a worksheet page to a PNG in the temp folder, to confirm the
    /// Publish/EMF pipeline works before relying on it in the UI.
    /// </param>
    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length > 0 && string.Equals(args[0], "--selftest", StringComparison.OrdinalIgnoreCase))
            return SelfTest(args.Length > 1 ? args[1] : null);

        if (args.Length > 0 && string.Equals(args[0], "--rastertest", StringComparison.OrdinalIgnoreCase))
            return RasterTest(
                args.Length > 1 ? args[1] : null,
                args.Length > 2 ? args[2] : null,
                args.Length > 3 ? args[3] : null);

        ApplicationConfiguration.Initialize();
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
        Application.ThreadException += (_, e) =>
            MessageBox.Show(
                e.Exception.Message,
                "Worksheet Watcher - unexpected error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

        Application.Run(new MainForm());
        return 0;
    }

    /// <summary>
    /// Headless connectivity check, used from the command line and during bring-up. Prints
    /// to stdout and returns 0 on success, 1 on failure.
    /// </summary>
    private static int SelfTest(string? notebookFilter)
    {
        try
        {
            var config = AppConfig.Load(out var warning);
            if (warning is not null) Console.WriteLine($"[config] {warning}");
            Console.WriteLine($"[config] Poll interval: {config.PollIntervalSeconds}s, thumbnail {config.ThumbnailWidth}x{config.ThumbnailHeight}");

            using var hierarchy = new OneNoteHierarchyService(config);
            var notebooks = hierarchy.GetOpenNotebooks();
            Console.WriteLine($"Open notebooks: {notebooks.Count}");
            foreach (var nb in notebooks)
                Console.WriteLine($"  - {nb.DisplayName}   {nb.Id}");

            var target = notebookFilter is null
                ? notebooks.FirstOrDefault()
                : notebooks.FirstOrDefault(n => n.Matches(notebookFilter));

            if (target is null)
            {
                Console.WriteLine("No notebook to walk. Done.");
                return 0;
            }

            Console.WriteLine($"\nWalking: {target.DisplayName}");
            var students = hierarchy.GetStudents(target.Id);
            Console.WriteLine($"Looks like a Class Notebook: {hierarchy.LooksLikeClassNotebook}");
            Console.WriteLine($"Groups (students): {students.Count}");
            foreach (var s in students)
                Console.WriteLine($"  {s.Name}: {s.Sections.Count} section(s), {s.Sections.Sum(sec => sec.Pages.Count)} page(s)");

            var resolver = new Services.WorksheetResolutionService(hierarchy);
            var titles = resolver.GetDistinctPageTitles(target.Id);
            Console.WriteLine($"\nDistinct page titles: {titles.Count}");
            foreach (var t in titles.Take(20)) Console.WriteLine($"  - {t}");

            Console.WriteLine("\nSelf-test OK.");
            return 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"SELF-TEST FAILED: {ex.GetType().Name}: {ex.Message}");
            return 1;
        }
    }

    /// <summary>
    /// Rasterises one student's copy of a worksheet page (via <see cref="Services.PageRasterService"/>,
    /// the same code path <see cref="Services.WatcherPollingService"/> uses) and saves it
    /// as a PNG in the temp folder, so the Publish/EMF pipeline can be eyeballed directly.
    /// </summary>
    private static int RasterTest(string? notebookFilter, string? pageTitleFilter, string? studentFilter)
    {
        if (pageTitleFilter is null)
        {
            Console.WriteLine("Usage: --rastertest <notebook> <pageTitle> [student]");
            return 1;
        }

        try
        {
            var config = AppConfig.Load(out _);
            using var hierarchy = new OneNoteHierarchyService(config);
            var notebooks = hierarchy.GetOpenNotebooks();
            var nb = notebookFilter is null
                ? notebooks.FirstOrDefault()
                : notebooks.FirstOrDefault(n => n.Matches(notebookFilter));
            if (nb is null)
            {
                Console.WriteLine($"No open notebook matches '{notebookFilter}'.");
                return 1;
            }
            Console.WriteLine($"Notebook: {nb.DisplayName}");

            var resolver = new Services.WorksheetResolutionService(hierarchy);
            var targets = resolver.ResolveWorksheet(nb.Id, pageTitleFilter).Where(t => t.HasPage).ToList();
            Console.WriteLine($"Students with a '{pageTitleFilter}' page: {targets.Count}");
            if (targets.Count == 0) return 1;

            var target = studentFilter is null
                ? targets[0]
                : targets.FirstOrDefault(t => t.StudentName.Contains(studentFilter, StringComparison.OrdinalIgnoreCase));
            if (target is null)
            {
                Console.WriteLine($"No matching student for '{studentFilter}'.");
                return 1;
            }
            Console.WriteLine($"Rasterising: {target.StudentName} / {target.SectionName} / {target.PageTitle}");

            var raster = new Services.PageRasterService(hierarchy.Client);
            var sw = System.Diagnostics.Stopwatch.StartNew();
            using var bitmap = raster.RenderPage(target.PageId!, config.ThumbnailWidth * 3, config.ThumbnailHeight * 3);
            sw.Stop();

            var outPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"WorksheetWatcher-rastertest-{DateTime.Now:HHmmss}.png");
            bitmap.Save(outPath, System.Drawing.Imaging.ImageFormat.Png);

            Console.WriteLine($"Rendered {bitmap.Width}x{bitmap.Height} in {sw.ElapsedMilliseconds}ms -> {outPath}");
            Console.WriteLine("Rastertest OK.");
            return 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"RASTERTEST FAILED: {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
            return 1;
        }
    }
}
