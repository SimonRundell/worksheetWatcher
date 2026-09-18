namespace WorksheetWatcher.Models;

/// <summary>
/// A single open notebook as reported by OneNote, used to populate the notebook picker.
/// </summary>
/// <param name="Id">OneNote hierarchy ID (an opaque <c>{...}{...}{n}</c> string).</param>
/// <param name="Name">
/// The notebook's storage name (its folder name), for example "U2FTUCO06B Notebook".
/// </param>
/// <param name="Nickname">
/// The friendly name OneNote shows in its own navigation, for example "Level 2 GRP B".
/// Empty when the notebook has no distinct nickname.
/// </param>
public record NotebookInfo(string Id, string Name, string Nickname = "")
{
    /// <summary>
    /// What the picker shows: the nickname (what the user recognises from OneNote), with
    /// the storage name in brackets when it differs, so a run can still be traced back to
    /// a folder.
    /// </summary>
    public string DisplayName =>
        string.IsNullOrWhiteSpace(Nickname) || Nickname.Equals(Name, StringComparison.OrdinalIgnoreCase)
            ? Name
            : $"{Nickname}  ({Name})";

    /// <summary>True when <paramref name="term"/> is a substring of the name or nickname.</summary>
    public bool Matches(string term) =>
        Name.Contains(term, StringComparison.OrdinalIgnoreCase) ||
        Nickname.Contains(term, StringComparison.OrdinalIgnoreCase);
}

/// <summary>
/// A student's private area within the Class Notebook: the top-level section group,
/// excluding the shared "Content Library" and "Collaboration Space". When the picked
/// notebook is not a Class Notebook this still represents a top-level section group,
/// just labelled generically.
/// </summary>
/// <param name="Id">OneNote hierarchy ID.</param>
/// <param name="Name">Section-group name (the student's name in a Class Notebook).</param>
/// <param name="Sections">Every section beneath this group, flattened across nested sub-groups.</param>
public record StudentNode(string Id, string Name, IReadOnlyList<SectionNode> Sections);

/// <summary>A section within a student's area (for example "Worksheets", "Homework").</summary>
/// <param name="Id">OneNote hierarchy ID.</param>
/// <param name="Name">Section name. Nested groups are prefixed, e.g. "Unit 1 / Worksheets".</param>
/// <param name="Pages">Pages in the section, in notebook order, sub-pages included.</param>
public record SectionNode(string Id, string Name, IReadOnlyList<PageNode> Pages);

/// <summary>A single OneNote page.</summary>
/// <param name="Id">OneNote hierarchy ID, used later with <c>GetPageContent</c> / <c>Publish</c>.</param>
/// <param name="Title">Page title. Untitled pages come back as "(untitled page)".</param>
/// <param name="LastModified">Last modified timestamp from the hierarchy XML.</param>
/// <param name="PageLevel">1 for a top-level page, 2 or 3 for a sub-page.</param>
public record PageNode(string Id, string Title, DateTime LastModified, int PageLevel);
