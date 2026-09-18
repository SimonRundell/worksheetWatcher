# Worksheet Watcher

A live "progress wall" for a OneNote Class Notebook: pick a notebook and a worksheet
(a page title that recurs once inside each student's own area), and watch a scrollable
grid of up to 30 student thumbnails update as they work - without leaving your desk.
Double-click a tile to see that student full screen. It is read-only: it never writes
back to OneNote.

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

## Using it

1. Open your Class Notebook in OneNote and let it finish syncing.
2. Start the tool. Pick the notebook from the **Notebook** dropdown (open notebooks
   only, shown by nickname with the storage name in brackets).
3. Pick or type the **Worksheet** page title - the title as it appears once inside each
   student's own section(s), for example "Unit 1 Lesson 3 Worksheet". The dropdown is
   populated from every distinct page title found in the notebook.
4. Set the poll interval (default 25s - how often changed pages are checked) and click
   **Start Watching**.
5. The grid fills with one tile per student, rendered from the live page. A tile shows
   grey with "No page yet" until that student creates the page; green once it has
   rendered; amber if the section is locked; the caption shows how long ago it last
   updated.
6. Double-click a tile to see that student full screen; **Close** returns to the grid.
7. **Refresh now** forces an immediate check instead of waiting for the next interval;
   **Stop** ends the watch session (OneNote is left completely untouched).

## Configuration

`WorksheetWatcher.config.json` sits next to the executable:

| Key | Meaning |
|---|---|
| `excludedSectionGroupNames` | Top-level section groups never treated as students (Class Notebook system areas). |
| `pollIntervalSeconds` | Default poll interval; floored at 10s at run time regardless of what is configured. |
| `thumbnailWidth` / `thumbnailHeight` | Base render size for each tile's rasterised page image. |
| `maxStudents` | Soft cap on students shown at once. The brief specifies 30. |
| `genericGroupLabel` | Row label used when the notebook has no Class Notebook student structure. |

## How the live thumbnail works

Each visible page is rasterised with OneNote's own `Application.Publish(pageId, path,
PublishFormat.pfEMF, "")` call - a single-page vector export (ink included), which is
cheaper than reading a whole notebook's content. Every poll tick does one cheap
hierarchy read to check each student's page `lastModifiedTime`; only pages that
actually changed are re-published and re-rasterised, so a quiet classroom costs almost
nothing per cycle. All OneNote calls run sequentially on one dedicated background
thread - the UI thread never touches OneNote directly.

## Project layout

```
WorksheetWatcher.sln
src/WorksheetWatcher/
  Program.cs                      entry point (also: --selftest)
  AppConfig.cs                    config.json loader
  UserSettings.cs                 %AppData% preferences (last notebook/worksheet)
  AppIcon.cs / appicon.ico         app icon
  WorksheetWatcher.config.json     runtime config
  libs/                           OneNote interop assembly (copied from WheresTheWork/libs)
  Models/
    HierarchyModels.cs             NotebookInfo, StudentNode, SectionNode, PageNode
    WatcherModels.cs                StudentPageTarget, ThumbnailState, PollStatus
  Services/
    OneNoteComClient.cs             OneNote interop wrapper (+ Publish for rasterising)
    OneNoteHierarchyService.cs      parses hierarchy XML into students/sections/pages
    WorksheetResolutionService.cs   distinct page titles + per-student page matching
    PageRasterService.cs            Publish -> EMF -> Bitmap
    WatcherPollingService.cs        background loop: diff timestamps, rasterise changed pages
  Forms/
    MainForm.cs                     notebook + worksheet pickers, Start/Stop, thumbnail grid
    StudentThumbnailControl.cs      one student's tile: image, name, status, caption
    FullScreenViewForm.cs           enlarged single-student view
tools/icon-generator/               reused from WheresTheWork (not in the .sln)
```

## Licence

Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International (CC BY-NC-SA
4.0). See [LICENSE](LICENSE).
