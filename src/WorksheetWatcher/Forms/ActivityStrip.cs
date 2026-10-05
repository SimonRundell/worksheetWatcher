using System.Drawing;
using System.Windows.Forms;

namespace WorksheetWatcher.Forms;

/// <summary>
/// A bar of one small chip per student, showing their initials, so the whole class can be
/// read at a glance without scrolling the tiles. When a student's work changes their chip
/// blinks red, then fades back to normal - so who is actively typing and who is coasting
/// stands out immediately. Students who have not started the worksheet are greyed out.
/// Hovering a chip shows the full name; clicking it raises <see cref="ChipClicked"/>.
/// </summary>
public sealed class ActivityStrip : FlowLayoutPanel
{
    private static readonly Color Alert = Color.FromArgb(0xE5, 0x39, 0x35);
    private static readonly Color AlertPale = Color.FromArgb(0xFF, 0xCD, 0xD2);
    private static readonly Color Idle = Color.FromArgb(0xDD, 0xE6, 0xF0);
    private static readonly Color NotStarted = Color.FromArgb(0xF1, 0xF1, 0xF1);

    private const double BlinkSeconds = 3;
    private const int BlinkPhaseMilliseconds = 300;

    private sealed class Chip : Label
    {
        public string StudentId = string.Empty;
        public string FullName = string.Empty;
        public bool Started;
        public DateTime? LastChange;
    }

    private readonly Dictionary<string, Chip> _chips = new();
    private readonly ToolTip _tip = new();
    private readonly System.Windows.Forms.Timer _timer = new() { Interval = 150 };

    /// <summary>How long, in seconds, a chip stays highlighted after a change (blinking, then fading).</summary>
    public double FlashSeconds { get; set; } = 20;

    /// <summary>Raised with the student's ID when their chip is clicked.</summary>
    public event EventHandler<string>? ChipClicked;

    public ActivityStrip()
    {
        Dock = DockStyle.Top;
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        WrapContents = true;
        FlowDirection = FlowDirection.LeftToRight;
        Padding = new Padding(6, 4, 6, 4);
        BackColor = Color.White;

        _timer.Tick += (_, _) => RefreshColours();
        _timer.Start();
    }

    /// <summary>Adds a chip for the student if there is not one yet, and records whether they have started.</summary>
    public void SetStudent(string studentId, string name, bool started)
    {
        if (!_chips.TryGetValue(studentId, out var chip))
        {
            chip = new Chip
            {
                StudentId = studentId,
                Size = new Size(48, 30),
                Margin = new Padding(3),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Trebuchet MS", 10f, FontStyle.Bold),
                BorderStyle = BorderStyle.FixedSingle,
                Cursor = Cursors.Hand
            };
            chip.Click += (_, _) => ChipClicked?.Invoke(this, chip.StudentId);
            _chips[studentId] = chip;
            Controls.Add(chip);
        }

        chip.FullName = name;
        chip.Started = started;
        _tip.SetToolTip(chip, name);
        RelabelAll();
        RefreshColours();
    }

    /// <summary>Starts a chip's red flash.</summary>
    public void Flash(string studentId)
    {
        if (!_chips.TryGetValue(studentId, out var chip)) return;
        chip.LastChange = DateTime.Now;
        RefreshColours();
    }

    /// <summary>Removes every chip - used when a new watch session starts.</summary>
    public void Clear()
    {
        foreach (var chip in _chips.Values) chip.Dispose();
        _chips.Clear();
        Controls.Clear();
    }

    /// <summary>
    /// Initials are the first and last name's first letters. Students whose initials
    /// collide are given more letters, trying more of the surname, then more of the first
    /// name, then both, until the group is told apart; anything still colliding gets a
    /// number. Recalculated for the whole bar whenever a chip is added, since a collision
    /// only shows up once the second student arrives.
    /// </summary>
    private void RelabelAll()
    {
        var chips = _chips.Values.ToList();
        var labels = chips.ToDictionary(c => c, c => Initials(c.FullName, 0));

        foreach (var clash in chips.GroupBy(c => labels[c]).Where(g => g.Count() > 1).ToList())
        {
            for (var extra = 1; extra <= 3; extra++)
            {
                var candidate = clash.ToDictionary(c => c, c => Initials(c.FullName, extra));
                var told = candidate.Values.Distinct().Count() == candidate.Count;
                if (!told && extra < 3) continue;

                foreach (var (chip, label) in candidate) labels[chip] = label;
                break;
            }
        }

        foreach (var clash in chips.GroupBy(c => labels[c]).Where(g => g.Count() > 1).ToList())
        {
            var n = 1;
            foreach (var chip in clash) labels[chip] += n++;
        }

        foreach (var chip in chips) chip.Text = labels[chip];
    }

    /// <summary>
    /// <paramref name="extra"/> 0 is first + last initial ("JS"); 1 adds a letter of the
    /// surname ("JSM"); 2 adds a letter of the first name ("JOS"); 3 adds both ("JOSM").
    /// </summary>
    private static string Initials(string name, int extra)
    {
        var parts = name.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 0) return "?";

        if (parts.Length == 1)
            return parts[0][..Math.Min(2, parts[0].Length)].ToUpperInvariant();

        var first = parts[0];
        var last = parts[^1];
        var firstLetters = extra >= 2 ? Math.Min(2, first.Length) : 1;
        var lastLetters = extra is 1 or 3 ? Math.Min(2, last.Length) : 1;
        return (first[..firstLetters] + last[..lastLetters]).ToUpperInvariant();
    }

    /// <summary>
    /// Blinks for the first few seconds after a change, then fades from red back to the
    /// normal colour over the rest of <see cref="FlashSeconds"/>.
    /// </summary>
    private void RefreshColours()
    {
        var now = DateTime.Now;
        foreach (var chip in _chips.Values)
        {
            Color back;
            if (!chip.Started)
            {
                back = NotStarted;
            }
            else if (chip.LastChange is not { } changed || (now - changed).TotalSeconds >= FlashSeconds)
            {
                back = Idle;
            }
            else
            {
                var age = (now - changed).TotalSeconds;
                if (age < BlinkSeconds)
                {
                    back = (int)(age * 1000 / BlinkPhaseMilliseconds) % 2 == 0 ? Alert : AlertPale;
                }
                else
                {
                    var t = Math.Clamp((age - BlinkSeconds) / Math.Max(1, FlashSeconds - BlinkSeconds), 0, 1);
                    back = Blend(Alert, Idle, t);
                }
            }

            if (chip.BackColor != back) chip.BackColor = back;

            var fore = !chip.Started ? Color.Silver
                : Luminance(back) < 150 ? Color.White
                : Color.FromArgb(0x24, 0x2F, 0x3D);
            if (chip.ForeColor != fore) chip.ForeColor = fore;
        }
    }

    private static Color Blend(Color from, Color to, double t) => Color.FromArgb(
        (int)(from.R + (to.R - from.R) * t),
        (int)(from.G + (to.G - from.G) * t),
        (int)(from.B + (to.B - from.B) * t));

    private static double Luminance(Color c) => 0.299 * c.R + 0.587 * c.G + 0.114 * c.B;

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _timer.Dispose();
            _tip.Dispose();
        }
        base.Dispose(disposing);
    }
}
