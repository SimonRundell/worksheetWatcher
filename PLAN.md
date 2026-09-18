# Worksheet Watcher - Scope and Plan

A live "progress wall" for a OneNote Class Notebook: pick a notebook and a worksheet
(page title), and watch a scrollable grid of up to 30 student thumbnails update as
students work, without leaving your desk. Click a thumbnail to see that student full
screen. Read-only - it never writes to OneNote.

Sibling project to `D:\WherestheWork`, reusing its proven OneNote COM approach and
conventions (see that project's README for the background on why a generated interop is
used instead of late binding).

## Decisions locked in from scoping

| Question | Decision |
|---|---|
| Thumbnail rendering | Rasterised page image via OneNote's own `Publish` export - looks exactly like the real page, ink included. |
| How a "worksheet" is identified | Same page title recurring once inside each student's own section/section-group, same model as WheresTheWork's page matching. |
| Refresh behaviour | Background poll on a timer (default ~25s) + a manual refresh, with a "last updated" stamp per thumbnail. |
| Full screen view | Enlarged version of the same in-app rendered preview, not a jump into real OneNote. |

## How the rasterised thumbnail actually works

The OneNote automation object has a `Publish(hierarchyId, targetFile, PublishFormat, string)`
method. `PublishFormat.pfEMF` exports a **single page** to an Enhanced Metafile - a vector
snapshot that includes ink, text, and images exactly as OneNote renders them. This is a
different, cheaper call than the whole-notebook walk WheresTheWork does, because it targets
one page ID at a time.

Pipeline per page:

1. `Publish(pageId, tempPath.emf, PublishFormat.pfEMF, "")` - writes the EMF to a temp file
   (`Publish` needs a real file path, not a stream).
2. Load it with `System.Drawing.Imaging.Metafile` and draw it into a fixed-size `Bitmap`
   (aspect-fit, letterboxed) sized for the thumbnail control - EMF is vector, so this scales
   cleanly at any thumbnail size.
3. Delete the temp file, cache the `Bitmap` in memory against that student.

This only runs for pages that actually changed (see polling, below), so steady-state cost
is small even with 30 students.

## Avoiding a hammered COM API: change detection before rendering

Every poll tick does **one cheap hierarchy call** (`GetHierarchy` scoped to the notebook,
same XML call WheresTheWork already uses) and reads each target page's
`lastModifiedTime` attribute out of it. That is compared against what we last saw per
student. Only pages whose timestamp actually moved get the expensive `Publish` + rasterise
treatment. So a full 30-student cycle where nobody has typed anything costs one XML call
and zero exports; a cycle where five students just edited costs one XML call plus five
exports.

All COM calls - hierarchy and publish - are made sequentially from a single dedicated
background thread that owns one `Application` COM object for the life of the watch
session, the same "one client, called one call at a time" discipline WheresTheWork's
`ReportBuilder` already uses. The UI thread never touches OneNote directly; updates are
marshalled back via `IProgress<T>` / `Control.Invoke`, matching the existing pattern.

## Identifying "the worksheet" and its 30 students

Reuses `OneNoteHierarchyService.GetStudents()` unchanged (top-level section groups minus
the configured system groups - Content Library, Collaboration Space, Teacher Only - are
the students, exactly as in WheresTheWork). A new, small resolver then:

1. Walks every student's flattened pages once to collect the **distinct page titles** that
   appear across the notebook, so the "Worksheet" dropdown can be populated after picking a
   notebook (mirrors how WheresTheWork's report discovers `PageTitleOrder`).
2. For the chosen title, finds the first matching page (case-insensitive, trimmed) inside
   each student's pages and builds one `StudentPageTarget { StudentName, PageId,
   SectionName }` per student who has it.
3. A student with no matching page yet gets a placeholder thumbnail ("not started / no
   page found") rather than an error - common early in a lesson before everyone has
   created their page.

## Project layout (mirrors WheresTheWork)

```
WorksheetWatcher.sln
src/WorksheetWatcher/
  Program.cs                      entry point
  AppConfig.cs                    config.json loader (poll interval, exclusions, thumb size)
  UserSettings.cs                 %AppData% preferences (interface scale, last notebook/worksheet)
  AppIcon.cs / appicon.ico         app icon (reuse WheresTheWork's icon-generator tool)
  WorksheetWatcher.config.json     runtime config
  libs/                           OneNote interop assembly (copied from WheresTheWork/libs)
  Models/
    HierarchyModels.cs             NotebookInfo, StudentNode, SectionNode, PageNode (reused as-is)
    WatcherModels.cs                StudentPageTarget, ThumbnailState, PollStatus
  Services/
    OneNoteComClient.cs             reused wrapper + new Publish() method
    OneNoteHierarchyService.cs      reused as-is
    WorksheetResolutionService.cs   distinct page titles + per-student page matching
    PageRasterService.cs            Publish -> EMF -> Bitmap
    WatcherPollingService.cs        background loop: diff timestamps, rasterise changed pages, report updates
  Forms/
    MainForm.cs                     notebook + worksheet pickers, Start/Stop watching, thumbnail grid
    StudentThumbnailControl.cs      one student's tile: image, name, last-updated stamp, status
    FullScreenViewForm.cs           enlarged single-student view, same refresh cadence
tools/icon-generator/               reused as-is from WheresTheWork (not in the .sln)
```

## The main window

- **Setup bar**: Notebook dropdown (open notebooks, same nickname format as WheresTheWork)
  -> Worksheet dropdown (populated once a notebook is picked) -> poll interval selector ->
  **Start Watching** / **Stop** button -> manual **Refresh now**.
- **Thumbnail grid**: a scrollable `FlowLayoutPanel` of `StudentThumbnailControl` tiles,
  sized so a normal window shows roughly 6 (3 columns x 2 rows) without scrolling, wrapping
  and scrolling to accommodate up to 30. Each tile shows the rasterised page, the student's
  name, a small "updated 12s ago" stamp, and a status colour (grey = no page yet, green
  flash = just updated, red = section locked / error - same per-page isolation as
  WheresTheWork so one locked section never stops the rest).
- Double-clicking (or an "Enlarge" affordance on) a tile opens **FullScreenViewForm**: the
  same cached bitmap enlarged, refreshing on the same poll cadence, with a **Back** button
  and optional **Prev/Next student** navigation.

## Configuration (`WorksheetWatcher.config.json`)

| Key | Meaning |
|---|---|
| `excludedSectionGroupNames` | Same purpose as WheresTheWork - system areas never treated as students. |
| `pollIntervalSeconds` | Default 25; UI can override per session with a sane floor (e.g. 10s) to protect the COM API. |
| `thumbnailWidth` / `thumbnailHeight` | Base tile render size (before DPI scaling). |
| `maxStudents` | Soft cap, default 30, matching the brief. |
| `genericGroupLabel` | Reused from WheresTheWork for non-Class-Notebook notebooks. |

## Assumptions to confirm (flag if any are wrong)

- A worksheet's page title is unique within a given student's own section tree - first
  match wins, same as WheresTheWork's page-title collection.
- No Excel/report export in this tool - it is a pure live viewer. (Easy to add a "snapshot
  report" later reusing WheresTheWork's `ExcelExporter` if you want one.)
- Bitmaps are in-memory only for the session, no on-disk thumbnail cache between runs.
- This stays read-only: no writes back to OneNote, no ink/selection data ever requested.

## Out of scope for v1 (flag if you want any of these pulled in)

- Marking/annotation from the Watcher itself.
- Historical timeline / activity log of when each student last edited.
- Multi-notebook or multi-worksheet watching in one window.
- Alerting (e.g. "student X hasn't started") - the grid is glance-only for now.

## Licence

Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International (CC BY-NC-SA 4.0),
same as WheresTheWork - `LICENSE` copied over, `.csproj` `<Copyright>` set accordingly.

---

Next step, once you've sanity-checked the above: I'll scaffold the solution (project
files, reused Services/Models copied and adapted, empty Forms wired up) so it builds and
shows the notebook/worksheet pickers end to end, then layer in rasterisation and polling.
