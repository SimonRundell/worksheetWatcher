using System.Xml.Linq;
using WorksheetWatcher.Models;

namespace WorksheetWatcher.Services;

/// <summary>
/// Turns the raw OneNote hierarchy XML into the app's data model: open notebooks for
/// the picker, and a notebook's students -&gt; sections -&gt; pages tree for the watcher.
///
/// The hierarchy XML is namespaced
/// (<c>http://schemas.microsoft.com/office/onenote/2013/onenote</c>); this class works
/// entirely through <see cref="XDocument"/> and an <see cref="XNamespace"/> constant
/// rather than string matching.
/// </summary>
public sealed class OneNoteHierarchyService : IDisposable
{
    /// <summary>The OneNote 2013 hierarchy/content XML namespace.</summary>
    public static readonly XNamespace One = "http://schemas.microsoft.com/office/onenote/2013/onenote";

    private readonly OneNoteComClient _com;
    private readonly AppConfig _config;
    private readonly bool _ownsClient;

    /// <summary>
    /// Set after <see cref="GetStudents"/> runs: false when the picked notebook had no
    /// section groups to exclude (so it is probably not a Class Notebook and the groups
    /// are labelled generically).
    /// </summary>
    public bool LooksLikeClassNotebook { get; private set; } = true;

    /// <param name="com">A connected <see cref="OneNoteComClient"/>. Ownership is not taken unless created here.</param>
    /// <param name="config">Loaded application config (exclusion list, labels).</param>
    public OneNoteHierarchyService(OneNoteComClient com, AppConfig config)
    {
        _com = com;
        _config = config;
        _ownsClient = false;
    }

    /// <summary>Convenience constructor that creates and connects its own COM client.</summary>
    public OneNoteHierarchyService(AppConfig config)
    {
        _config = config;
        _com = new OneNoteComClient();
        _com.Connect();
        _ownsClient = true;
    }

    /// <summary>The underlying COM client, for services that need page content or rasterisation.</summary>
    public OneNoteComClient Client => _com;

    /// <summary>Ensures the COM client is connected.</summary>
    public void Connect()
    {
        if (!_com.IsConnected) _com.Connect();
    }

    /// <summary>Every notebook currently open in OneNote, for the notebook picker.</summary>
    public IReadOnlyList<NotebookInfo> GetOpenNotebooks()
    {
        Connect();
        var xml = _com.GetNotebooksXml();
        var doc = XDocument.Parse(xml);

        return doc.Descendants(One + "Notebook")
            .Select(n => new NotebookInfo(
                Id: (string?)n.Attribute("ID") ?? string.Empty,
                Name: (string?)n.Attribute("name") ?? "(unnamed notebook)",
                Nickname: (string?)n.Attribute("nickname") ?? string.Empty))
            .Where(n => n.Id.Length > 0)
            .ToList();
    }

    /// <summary>
    /// Walks the notebook's full page-scope hierarchy and returns its students
    /// (top-level section groups, minus the configured system groups and any recycle
    /// bin), each with a flattened list of sections and their pages.
    /// </summary>
    public IReadOnlyList<StudentNode> GetStudents(string notebookId)
    {
        Connect();
        var xml = _com.GetNotebookTreeXml(notebookId);
        var doc = XDocument.Parse(xml);

        var notebook = doc.Descendants(One + "Notebook").FirstOrDefault()
            ?? throw new InvalidOperationException("The hierarchy XML did not contain a <Notebook> element.");

        var excluded = new HashSet<string>(
            _config.ExcludedSectionGroupNames.Select(NormaliseGroupName),
            StringComparer.OrdinalIgnoreCase);

        // Top-level section groups directly under the notebook.
        var topGroups = notebook.Elements(One + "SectionGroup").ToList();

        var keptGroups = topGroups
            .Where(g => !IsRecycleBin(g))
            .Where(g => !excluded.Contains(NormaliseGroupName((string?)g.Attribute("name"))))
            .ToList();

        // Heuristic from the design: if nothing was excluded, this probably is not a
        // Class Notebook, so report groups generically rather than as "students".
        LooksLikeClassNotebook = keptGroups.Count < topGroups.Count(g => !IsRecycleBin(g));

        var students = new List<StudentNode>();

        foreach (var group in keptGroups)
        {
            var name = (string?)group.Attribute("name") ?? _config.GenericGroupLabel;
            var id = (string?)group.Attribute("ID") ?? string.Empty;
            var sections = new List<SectionNode>();
            CollectSections(group, prefix: null, sections);
            students.Add(new StudentNode(id, name, sections));
        }

        // Some notebooks also place sections directly under the notebook root (rare for a
        // Class Notebook, common elsewhere). Surface those under a synthetic group so they
        // are not silently dropped.
        var looseSections = new List<SectionNode>();
        foreach (var section in notebook.Elements(One + "Section").Where(s => !IsRecycleBin(s)))
            AddSection(section, prefix: null, looseSections);

        if (looseSections.Count > 0)
        {
            students.Add(new StudentNode(
                Id: (string?)notebook.Attribute("ID") ?? "loose",
                Name: LooksLikeClassNotebook ? "(notebook root)" : _config.GenericGroupLabel,
                Sections: looseSections));
        }

        return students;
    }

    /// <summary>
    /// Recursively walks nested section groups, flattening every section into
    /// <paramref name="into"/> with a slash-joined name prefix so nesting is still legible.
    /// </summary>
    private void CollectSections(XElement group, string? prefix, List<SectionNode> into)
    {
        var groupName = (string?)group.Attribute("name");
        var childPrefix = string.IsNullOrEmpty(prefix)
            ? (string.IsNullOrEmpty(groupName) ? null : groupName)
            : $"{prefix} / {groupName}";

        foreach (var section in group.Elements(One + "Section").Where(s => !IsRecycleBin(s)))
            AddSection(section, childPrefix, into);

        foreach (var sub in group.Elements(One + "SectionGroup").Where(g => !IsRecycleBin(g)))
            CollectSections(sub, childPrefix, into);
    }

    /// <summary>Materialises one section and its pages (sub-pages included).</summary>
    private void AddSection(XElement section, string? prefix, List<SectionNode> into)
    {
        var rawName = (string?)section.Attribute("name") ?? "(unnamed section)";
        var name = string.IsNullOrEmpty(prefix) ? rawName : $"{prefix} / {rawName}";
        var id = (string?)section.Attribute("ID") ?? string.Empty;

        var pages = section.Elements(One + "Page")
            .Select(ParsePage)
            .ToList();

        into.Add(new SectionNode(id, name, pages));
    }

    /// <summary>Reads a single <c>&lt;one:Page&gt;</c> element into a <see cref="PageNode"/>.</summary>
    private static PageNode ParsePage(XElement page)
    {
        var title = (string?)page.Attribute("name");
        if (string.IsNullOrWhiteSpace(title)) title = "(untitled page)";

        var lastModified = ParseDate((string?)page.Attribute("lastModifiedTime"))
                           ?? ParseDate((string?)page.Attribute("dateTime"))
                           ?? DateTime.MinValue;

        var level = 1;
        if (int.TryParse((string?)page.Attribute("pageLevel"), out var parsed) && parsed > 0)
            level = parsed;

        return new PageNode(
            Id: (string?)page.Attribute("ID") ?? string.Empty,
            Title: title!,
            LastModified: lastModified,
            PageLevel: level);
    }

    private static DateTime? ParseDate(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        return DateTime.TryParse(
            value,
            System.Globalization.CultureInfo.InvariantCulture,
            System.Globalization.DateTimeStyles.AdjustToUniversal | System.Globalization.DateTimeStyles.AssumeUniversal,
            out var dt)
            ? dt.ToLocalTime()
            : null;
    }

    /// <summary>
    /// Canonical form of a section-group name for exclusion matching: trimmed, leading
    /// punctuation/underscores stripped, lower-cased, internal whitespace collapsed. This
    /// lets a config entry of "Content Library" match OneNote's "_Content Library",
    /// "Content Library ", and similar per-tenant variants.
    /// </summary>
    private static string NormaliseGroupName(string? name)
    {
        if (string.IsNullOrWhiteSpace(name)) return string.Empty;
        var trimmed = name.Trim().TrimStart('_', '-', ' ', '.', '*');
        var collapsed = System.Text.RegularExpressions.Regex.Replace(trimmed, @"\s+", " ");
        return collapsed.ToLowerInvariant();
    }

    /// <summary>True when an element is OneNote's recycle bin / deleted-pages container.</summary>
    private static bool IsRecycleBin(XElement element) =>
        string.Equals((string?)element.Attribute("isRecycleBin"), "true", StringComparison.OrdinalIgnoreCase) ||
        string.Equals((string?)element.Attribute("isInRecycleBin"), "true", StringComparison.OrdinalIgnoreCase) ||
        string.Equals((string?)element.Attribute("isDeletedPages"), "true", StringComparison.OrdinalIgnoreCase);

    /// <summary>Releases the COM client if this instance created it.</summary>
    public void Dispose()
    {
        if (_ownsClient) _com.Dispose();
    }
}
