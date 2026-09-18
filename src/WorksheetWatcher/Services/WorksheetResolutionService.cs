using WorksheetWatcher.Models;

namespace WorksheetWatcher.Services;

/// <summary>
/// Turns "a notebook and a worksheet title" into the concrete per-student pages the
/// watcher needs to poll. Built on top of <see cref="OneNoteHierarchyService.GetStudents"/>
/// - it does not talk to OneNote directly.
/// </summary>
public sealed class WorksheetResolutionService
{
    private readonly OneNoteHierarchyService _hierarchy;

    public WorksheetResolutionService(OneNoteHierarchyService hierarchy) => _hierarchy = hierarchy;

    /// <summary>
    /// Every distinct page title in the notebook, in first-seen order, for populating the
    /// "Worksheet" dropdown once a notebook is picked. Mirrors how WheresTheWork discovers
    /// its report's page-title columns.
    /// </summary>
    public IReadOnlyList<string> GetDistinctPageTitles(string notebookId)
    {
        var students = _hierarchy.GetStudents(notebookId);
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var titles = new List<string>();

        foreach (var student in students)
            foreach (var section in student.Sections)
                foreach (var page in section.Pages)
                    if (seen.Add(page.Title)) titles.Add(page.Title);

        return titles;
    }

    /// <summary>
    /// For every student in the notebook, finds the first page whose title matches
    /// <paramref name="worksheetTitle"/> (case-insensitive, trimmed). A student with no
    /// such page yet gets a <see cref="StudentPageTarget.NotStarted"/> placeholder rather
    /// than being dropped, since that is the normal state early in a lesson.
    /// </summary>
    public IReadOnlyList<StudentPageTarget> ResolveWorksheet(string notebookId, string worksheetTitle)
    {
        var students = _hierarchy.GetStudents(notebookId);
        var wanted = worksheetTitle.Trim();
        var targets = new List<StudentPageTarget>();

        foreach (var student in students)
        {
            var match = student.Sections
                .SelectMany(section => section.Pages.Select(page => (Section: section, Page: page)))
                .FirstOrDefault(x => string.Equals(x.Page.Title.Trim(), wanted, StringComparison.OrdinalIgnoreCase));

            targets.Add(match.Page is null
                ? StudentPageTarget.NotStarted(student.Id, student.Name)
                : new StudentPageTarget(
                    StudentId: student.Id,
                    StudentName: student.Name,
                    SectionName: match.Section.Name,
                    PageId: match.Page.Id,
                    PageTitle: match.Page.Title,
                    LastModified: match.Page.LastModified));
        }

        return targets;
    }
}
