# Worksheet Watcher

A live "progress wall" for a OneNote Class Notebook: pick a notebook and a worksheet
(a page title that recurs once inside each student's own area), and watch a scrollable
grid of up to 30 student thumbnails update as they work - without leaving your desk.
A bar of student initials runs across the top and flashes red whenever someone's work
changes, so you can see at a glance who is actively typing and who is coasting; each
tile zooms in on the spot that changed most recently. Click a tile, or its
magnifying-glass button, to see that student full screen. It is read-only: it never
writes back to OneNote.

See [PLAN.md](PLAN.md) for the full scope, architecture, and design decisions.

Sibling project to `D:\WherestheWork`, reusing its proven OneNote COM automation
approach (a pre-generated interop assembly rather than late binding - see that
project's README for why) and general conventions.

## Requirements

- Windows 10 or 11
- Classic desktop OneNote installed and running, with the target Class Notebook open
- .NET 8 Desktop Runtime (or the SDK) to build

## Build and run

```bash
dotnet build WorksheetWatcher.sln -c Release
```

```bash
dotnet run --project src/WorksheetWatcher/WorksheetWatcher.csproj -c Release
```

Or open `WorksheetWatcher.sln` in Visual Studio 2022 and press F5.

A headless connectivity check is built in (no UI):

```bash
dotnet run --project src/WorksheetWatcher/WorksheetWatcher.csproj -c Release -- --selftest "Level 2 GRP B"
```

It lists open notebooks, walks the matched one, lists its students, and lists the
distinct page titles found - useful for confirming a notebook's structure before
picking a worksheet in the UI.

A second check exercises the real export / change-detection / close-up pipeline on one
student's page and saves the close-ups as PNGs in your temp folder:

```bash
dotnet run --project src/WorksheetWatcher/WorksheetWatcher.csproj -c Release -- --focustest "Level 2 GRP B" "Unit 1 Lesson 1" IRYNA
```

It confirms that exporting the same page twice reads as "no change" (so nothing flashes
falsely) and that a known edit is found and sits inside the close-up.

## Publishing a standalone .exe

To hand someone a single file that runs with no .NET install and nothing else needed
alongside it:

```bash
dotnet publish src/WorksheetWatcher/WorksheetWatcher.csproj -p:PublishProfile=FolderProfile
```

(or Visual Studio's **Publish...** using the `FolderProfile` profile). This bundles the
.NET runtime and every dependency into one ~80MB `WorksheetWatcher.exe`, targeting 64-bit
Windows (`win-x64`). Output lands in
`src/WorksheetWatcher/bin/Release/net8.0-windows10.0.19041.0/win-x64/publish/`, and is
verified to run correctly (confirmed against a live OneNote via `--selftest` and
`--focustest`).

The only other file in that folder is `WorksheetWatcher.config.json` - deliberately left
loose rather than bundled, so it can be edited (poll defaults, exclusions) without
republishing. Everything else - including the OneNote interop assembly - is baked into
the exe.

## Using it

1. Open your Class Notebook in OneNote and let it finish syncing.
2. Start the tool. Pick the notebook from the **Notebook** dropdown (open notebooks
   only, shown by nickname with the storage name in brackets).
3. Pick or type the **Worksheet** page title - the title as it appears once inside each
   student's own section(s), for example "Unit 1 Lesson 3 Worksheet". The dropdown is
   populated from every distinct page title found in the notebook.
4. Pick a poll interval - **5s, 10s, 15s or 30s** (how often changed pages are
   checked) - and click **Start Watching**. All four are safe: a 23-student notebook's
   per-tick hierarchy check costs 50-200ms regardless of interval, so 5s doesn't risk
   overloading OneNote's COM API. The real pacing limit is how many students' pages
   actually changed that tick, since each one costs a separate ~1-2s render - if several
   students type at once a cycle can simply take longer than the chosen interval, which
   is fine (calls are never made concurrently).
5. Across the top is the **activity bar**: one chip per student showing their initials
   (hover for the full name; students who share initials get extra letters). A chip is
   grey until that student has created the page. When a student's work visibly changes
   their chip **blinks red** for a few seconds, then fades back to normal over the rest
   of `flashSeconds` (5 by default) - so a glance shows who is typing right now and who
   is not. The same initials follow each student's name on their tile, so a flashing
   chip can be matched to its tile. Click a chip to jump the grid to that student.
6. Below it, one large tile per student. Each tile is a sharp **close-up of the most
   recent change**, not a shrunken whole page: the page is compared with how it looked
   last time and the tile re-centres on whatever changed. Before any change has been
   seen it shows the end of the student's work. A tile shows grey with "No page yet"
   until the page exists; amber if the section is locked; the caption shows how long
   ago it last changed.
7. Click a tile, or the magnifying-glass button in its bottom-right corner, to see that
   student full screen - again the close-up of the latest change, rendered at your
   screen's resolution. The button at the bottom switches to **Show whole page**: every
   sheet stacked and scrollable, with the close-up area boxed in red. **Close** (or
   Esc) returns to the grid. (There is deliberately no hover-triggered popup - an
   earlier version had one and it proved obtrusive when scanning across the grid.)
8. **Refresh now** forces an immediate check instead of waiting for the next interval;
   **Stop** ends the watch session (OneNote is left completely untouched).

If a student's edits don't seem to be showing up: every poll tick asks OneNote to sync
the notebook from the cloud before checking anything, but that sync itself happens in
OneNote's own background process, not instantly - so expect it to typically show up
within a tick or two of the chosen interval, not necessarily the very next one.

## Configuration

`WorksheetWatcher.config.json` sits next to the executable:

| Key | Meaning |
|---|---|
| `excludedSectionGroupNames` | Top-level section groups never treated as students (Class Notebook system areas). |
| `pollIntervalSeconds` | Which of the 5/10/15/30s presets is pre-selected on start-up; floored at 5s at run time regardless of what is configured. |
| `tileDisplayWidth` / `tileDisplayHeight` | The on-screen size of a grid tile, default 520x390 (large enough to read typing directly, at the cost of fewer tiles per screen - scroll for the rest). Its shape also sets the shape of the close-up, which is rendered at twice this size so it stays crisp. |
| `focusWindowFraction` | How much of the page's width a tile's close-up covers, default 0.5. Smaller zooms in tighter (bigger text, less context); larger shows more. It widens automatically when the change itself is bigger, such as a pasted screenshot. |
| `flashSeconds` | How long a student's initials chip stays highlighted after a change, default 5: blinking red for most of that (up to 3 seconds), then fading. Raise it if you want the red to linger longer. |
| `maxStudents` | Soft cap on students shown at once. The brief specifies 30. |
| `genericGroupLabel` | Row label used when the notebook has no Class Notebook student structure. |

## How the live thumbnail works

Each poll tick first asks OneNote to sync the notebook from the cloud
(`Application.SyncHierarchy`) - without this, a student's edits sit on their own device
or in the cloud and never reach the local hierarchy this app reads, no matter how often
it polls. It then does one cheap hierarchy read to check each student's page
`lastModifiedTime`. Only a page whose time moved gets the heavier treatment:

1. **Export as PDF** (`Application.Publish` with `pfPDF`). PDF rather than EMF because
   OneNote's EMF export contains only the *first A4 sheet* of a page - anything a
   student writes below that is silently cut off - whereas the PDF holds every sheet.
2. **Fingerprint and compare.** Each sheet is rendered at modest size and averaged down
   to a small greyscale grid, then compared with the previous export's. The first run of
   changed rows is where the student is working. ("First", not "biggest": typing in the
   middle pushes everything below it down, which changes all of that too - the first
   changed row is the real edit.) A modified time that moved without anything visibly
   changing - OneNote touches pages when it syncs - is ignored: no redraw, no flash.
3. **Render a close-up** of just that spot, straight from the PDF's vector data, at twice
   the tile's size. Text stays razor sharp however far in it zooms.

PDFs are rendered with Windows' own built-in PDF engine (`Windows.Data.Pdf`), so nothing
extra ships. Only the latest PDF per student is kept (a few hundred KB), which is also
what the full-screen view renders from, at whatever size your screen is. A quiet
classroom costs almost nothing per cycle. All OneNote calls run sequentially on one
dedicated background thread - the UI thread never touches OneNote directly.

## Project layout

```
WorksheetWatcher.sln
src/WorksheetWatcher/
  Program.cs                      entry point (also: --selftest, --focustest, --polltimingtest)
  AppConfig.cs                    config.json loader
  UserSettings.cs                 %AppData% preferences (last notebook/worksheet)
  AppIcon.cs / appicon.ico         app icon
  WorksheetWatcher.config.json     runtime config
  libs/                           OneNote interop assembly (copied from WheresTheWork/libs)
  Models/
    HierarchyModels.cs             NotebookInfo, StudentNode, SectionNode, PageNode
    WatcherModels.cs                StudentPageTarget, ThumbnailState, PollStatus
  Services/
    OneNoteComClient.cs             OneNote interop wrapper (+ Publish for exporting)
    OneNoteHierarchyService.cs      parses hierarchy XML into students/sections/pages
    WorksheetResolutionService.cs   distinct page titles + per-student page matching
    PageExportService.cs            Publish -> PDF bytes
    OpenedPdf.cs                    Windows' built-in PDF renderer (whole sheets or any region)
    PageChangeAnalyzer.cs           fingerprint sheets, find where a page changed, plan the close-up
    WatcherPollingService.cs        background loop: sync, export changed pages, find the change, render
  Forms/
    MainForm.cs                     notebook + worksheet pickers, Start/Stop, activity bar, grid
    ActivityStrip.cs                initials bar that flashes red on change
    StudentThumbnailControl.cs      one student's tile: close-up, name, status, caption
    FullScreenViewForm.cs           full-window close-up, or every sheet stacked
tools/icon-generator/               reused from WheresTheWork (not in the .sln)
```

## Licence

Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International (CC BY-NC-SA
4.0). See [LICENSE](LICENSE).
