using System.Drawing;
using System.Windows.Forms;

namespace OfficePicture.Core;

public sealed class PreviewPane : UserControl
{
    private readonly PictureBox _pictureBox;
    private readonly Label _titleLabel;
    private readonly Label _statusLabel;
    private Image? _currentImage;

    public PreviewPane()
    {
        BackColor = Color.White;
        Dock = DockStyle.Fill;
        MinimumSize = new Size(240, 160);

        _titleLabel = new Label
        {
            Dock = DockStyle.Top,
            Height = 34,
            Padding = new Padding(12, 10, 8, 0),
            Font = new Font("Segoe UI", 10F, FontStyle.Bold),
            Text = "图片预览"
        };

        _statusLabel = new Label
        {
            Dock = DockStyle.Bottom,
            Height = 34,
            Padding = new Padding(12, 8, 8, 0),
            ForeColor = Color.DimGray,
            AutoEllipsis = true,
            Text = "在 Word、PowerPoint 或 Excel 中选中图片"
        };

        _pictureBox = new PictureBox
        {
            Dock = DockStyle.Fill,
            BackColor = Color.FromArgb(245, 247, 250),
            SizeMode = PictureBoxSizeMode.Zoom,
            BorderStyle = BorderStyle.FixedSingle
        };

        Controls.Add(_pictureBox);
        Controls.Add(_statusLabel);
        Controls.Add(_titleLabel);
    }

    public void ShowImage(Image image, string host, string source)
    {
        var copy = new Bitmap(image);
        var old = _currentImage;
        _currentImage = copy;
        _pictureBox.Image = copy;
        _titleLabel.Text = $"图片预览 · {host}";
        _statusLabel.Text = $"{source} · {copy.Width} × {copy.Height}px";
        old?.Dispose();
    }

    public void ShowMessage(string message)
    {
        _titleLabel.Text = "图片预览";
        _statusLabel.Text = message;
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _pictureBox.Image = null;
            _currentImage?.Dispose();
        }
        base.Dispose(disposing);
    }
}
