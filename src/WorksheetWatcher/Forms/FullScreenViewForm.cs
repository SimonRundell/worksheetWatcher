using System.Drawing;
using System.Windows.Forms;

namespace WorksheetWatcher.Forms;

/// <summary>
/// Enlarged single-student view: the same rendered preview as its thumbnail, just full
/// window sized. Non-modal, so the main grid keeps polling and updating behind it;
/// <see cref="MainForm"/> forwards this student's next render here via
/// <see cref="UpdateImage"/> for as long as the window stays open.
/// </summary>
public sealed class FullScreenViewForm : Form
{
    private readonly PictureBox _picture = new()
    {
        Dock = DockStyle.Fill,
        SizeMode = PictureBoxSizeMode.Zoom,
        BackColor = Color.White
    };

    private readonly Label _nameLabel = new()
    {
        Dock = DockStyle.Top,
        Height = 44,
        TextAlign = ContentAlignment.MiddleCenter,
        Font = new Font("Trebuchet MS", 16f, FontStyle.Bold)
    };

    /// <summary>The student this window is showing, so the host knows whether to forward an update.</summary>
    public string StudentId { get; }

    public FullScreenViewForm(string studentId, string studentName, Image? currentImage)
    {
        StudentId = studentId;

        Text = $"{studentName} - Worksheet Watcher";
        Icon = AppIcon.Value;
        StartPosition = FormStartPosition.CenterParent;
        WindowState = FormWindowState.Maximized;
        MinimumSize = new Size(640, 480);
        Font = new Font("Trebuchet MS", 9.75f);

        _nameLabel.Text = studentName;
        _picture.Image = currentImage;

        var btnClose = new Button
        {
            Text = "Close",
            Dock = DockStyle.Bottom,
            Height = 36,
            Font = new Font("Trebuchet MS", 10f)
        };
        btnClose.Click += (_, _) => Close();

        Controls.Add(_picture);
        Controls.Add(btnClose);
        Controls.Add(_nameLabel);
    }

    /// <summary>Swaps in a freshly rendered image, called by the host when this student's tile updates.</summary>
    public void UpdateImage(Image? image) => _picture.Image = image;
}
