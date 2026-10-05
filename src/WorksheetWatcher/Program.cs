using System.Drawing;
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
    /// <c>--focustest &lt;notebook&gt; &lt;pageTitle&gt; [student]</c> - runs the export,
    /// change-detection and close-up pipeline on one student's copy of a worksheet and saves
    /// the results as PNGs in the temp folder, to check it before relying on it in the UI.
    /// <c>--polltimingtest &lt;notebook&gt; &lt;pageTitle&gt; [rounds]</c> - times the per-tick
    /// sync and hierarchy re-check.
    /// </param>
    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length > 0 && string.Equals(args[0], "--selftest", StringComparison.OrdinalIgnoreCase))
            return SelfTest(args.Length > 1 ? args[1] : null);

        if (args.Length > 0 && string.Equals(args[0], "--focustest", StringComparison.OrdinalIgnoreCase))
            return FocusTest(
                args.Length > 1 ? args[1] : null,
                args.Length > 2 ? args[2] : null,
                args.Length > 3 ? args[3] : null);

        if (args.Length > 0 && string.Equals(args[0], "--polltimingtest", StringComparison.OrdinalIgnoreCase))
            return PollTimingTest(
                args.Length > 1 ? args[1] : null,
                args.Length > 2 ? args[2] : null,
                args.Length > 3 && int.TryParse(args[3], out var rounds) ? rounds : 5);

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
            Console.WriteLine($"[config] Poll interval: {config.PollIntervalSeconds}s, tile {config.TileDisplayWidth}x{config.TileDisplayHeight}, focus window {config.FocusWindowFraction:P0} of page width");

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

            var resolver = new WorksheetResolutionService(hierarchy);
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
    /// Runs the real export / fingerprint / close-up pipeline on one student's page and
    /// checks the two properties the activity bar depends on: that exporting the same page
    /// twice reads as "no change" (otherwise every poll would flash), and that a known edit
    /// is found and centred in the close-up.
    /// </summary>
    private static int FocusTest(string? notebookFilter, string? pageTitleFilter, string? studentFilter)
    {
        if (pageTitleFilter is null)
        {
            Console.WriteLine("Usage: --focustest <notebook> <pageTitle> [student]");
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

            var targets = new WorksheetResolutionService(hierarchy).ResolveWorksheet(nb.Id, pageTitleFilter).Where(t => t.HasPage).ToList();
            var target = studentFilter is null
                ? targets.FirstOrDefault()
                : targets.FirstOrDefault(t => t.StudentName.Contains(studentFilter, StringComparison.OrdinalIgnoreCase));
            if (target is null)
            {
                Console.WriteLine("No matching student with that page.");
                return 1;
            }
            Console.WriteLine($"{nb.DisplayName}: {target.StudentName} / {target.PageTitle}");

            var exporter = new PageExportService(hierarchy.Client);
            var aspect = (double)config.TileDisplayWidth / config.TileDisplayHeight;
            var failures = 0;

            var sw = System.Diagnostics.Stopwatch.StartNew();
            var pdfBytes = exporter.ExportPdf(target.PageId!);
            Console.WriteLine($"Export:      {sw.ElapsedMilliseconds}ms, {pdfBytes.Length / 1024}KB");

            using var pdf = OpenedPdf.Open(pdfBytes);
            sw.Restart();
            var first = PageChangeAnalyzer.Sign(pdf);
            Console.WriteLine($"Fingerprint: {sw.ElapsedMilliseconds}ms, {pdf.PageCount} sheet(s)");

            // 1. The same page exported again must read as unchanged.
            using var pdfAgain = OpenedPdf.Open(exporter.ExportPdf(target.PageId!));
            var again = PageChangeAnalyzer.Sign(pdfAgain);
            var spurious = PageChangeAnalyzer.FindChange(first, again, config.FocusWindowFraction, aspect);
            Console.WriteLine($"Re-export unchanged:  {(spurious is null ? "PASS" : $"FAIL - reported a change on sheet {spurious.PageIndex + 1} at {spurious.Window}")}");
            if (spurious is not null) failures++;

            // 2. Draw a known edit onto sheet 1 and check it is found, and sits inside the window.
            var size = pdf.PageSize(0);
            var height = (int)Math.Round(PageChangeAnalyzer.AnalysisWidth * size.Height / size.Width);
            using var edited = pdf.Render(0, null, PageChangeAnalyzer.AnalysisWidth, height);
            var before = PageChangeAnalyzer.SignBitmap(edited, size);
            var edit = new Rectangle(edited.Width * 55 / 100, edited.Height * 40 / 100, 260, 28);
            using (var g = Graphics.FromImage(edited))
                g.FillRectangle(Brushes.Black, edit);
            var after = PageChangeAnalyzer.SignBitmap(edited, size);

            var found = PageChangeAnalyzer.FindChange(
                new DocumentSignature(new[] { before }), new DocumentSignature(new[] { after }),
                config.FocusWindowFraction, aspect);

            var scale = size.Width / PageChangeAnalyzer.AnalysisWidth;
            var editDip = new RectangleF(edit.X * scale, edit.Y * scale, edit.Width * scale, edit.Height * scale);
            var contains = found is not null && found.Window.Contains(editDip);
            Console.WriteLine($"Synthetic edit found: {(contains ? "PASS" : "FAIL")}  edit={editDip}  window={found?.Window}");
            if (!contains) failures++;

            // 3. Render the close-ups so they can be looked at.
            var initial = PageChangeAnalyzer.InitialFocus(first, config.FocusWindowFraction, aspect);
            sw.Restart();
            using var closeUp = pdf.Render(initial.PageIndex, initial.Window, config.TileDisplayWidth * 2, config.TileDisplayHeight * 2);
            Console.WriteLine($"Close-up:    {sw.ElapsedMilliseconds}ms, {closeUp.Width}x{closeUp.Height}, initial focus sheet {initial.PageIndex + 1} {initial.Window}");

            var temp = System.IO.Path.GetTempPath();
            closeUp.Save(System.IO.Path.Combine(temp, "WorksheetWatcher-focus-initial.png"), System.Drawing.Imaging.ImageFormat.Png);
            if (found is not null)
            {
                using var edit1 = pdf.Render(found.PageIndex, found.Window, config.TileDisplayWidth * 2, config.TileDisplayHeight * 2);
                edit1.Save(System.IO.Path.Combine(temp, "WorksheetWatcher-focus-edit.png"), System.Drawing.Imaging.ImageFormat.Png);
            }
            Console.WriteLine($"Saved PNGs to {temp}");

            Console.WriteLine(failures == 0 ? "Focustest OK." : $"Focustest: {failures} FAILED.");
            return failures == 0 ? 0 : 1;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"FOCUSTEST FAILED: {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
            return 1;
        }
    }

    /// <summary>
    /// Times a bare sync plus <see cref="WorksheetResolutionService.ResolveWorksheet"/> call
    /// (what every poll tick pays even when nothing changed, with no export) over several
    /// rounds, to find a realistic floor for the poll interval on a real-sized notebook.
    /// </summary>
    private static int PollTimingTest(string? notebookFilter, string? pageTitleFilter, int rounds)
    {
        if (pageTitleFilter is null)
        {
            Console.WriteLine("Usage: --polltimingtest <notebook> <pageTitle> [rounds]");
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

            var resolver = new WorksheetResolutionService(hierarchy);
            var times = new List<long>();

            for (var i = 0; i < rounds; i++)
            {
                var sw = System.Diagnostics.Stopwatch.StartNew();
                hierarchy.Client.SyncNode(nb.Id); // matches what WatcherPollingService does every tick
                var targets = resolver.ResolveWorksheet(nb.Id, pageTitleFilter);
                sw.Stop();
                times.Add(sw.ElapsedMilliseconds);
                Console.WriteLine($"  round {i + 1}: {sw.ElapsedMilliseconds}ms ({targets.Count(t => t.HasPage)} of {targets.Count} students have the page)");
            }

            Console.WriteLine($"\nmin={times.Min()}ms  avg={times.Average():F0}ms  max={times.Max()}ms");
            Console.WriteLine("Polltimingtest OK.");
            return 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"POLLTIMINGTEST FAILED: {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
            return 1;
        }
    }
}
