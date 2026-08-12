using System;
using System.Drawing;
using System.Windows.Forms;

namespace OfficePicture.Core;

public sealed class ImagePreviewForm : Form
{
    private const float MinZoom = 0.1F;
    private const float MaxZoom = 8F;

    private readonly Panel _viewport;
    private readonly PictureBox _pictureBox;
    private readonly ToolStripLabel _zoomLabel;
    private readonly Image _image;
    private float _zoom = 1F;
    private bool _fitToWindow = true;

    public ImagePreviewForm(Image image, string host)
    {
        _image = new Bitmap(image);
        Text = $"图片预览 - {host}";
        StartPosition = FormStartPosition.CenterParent;
        Size = new Size(1000, 720);
        MinimumSize = new Size(480, 360);
        BackColor = Color.FromArgb(28, 28, 30);
        KeyPreview = true;
        ShowIcon = false;

        var toolStrip = new ToolStrip
        {
            GripStyle = ToolStripGripStyle.Hidden,
            BackColor = Color.FromArgb(245, 245, 247),
            Padding = new Padding(8, 3, 8, 3),
            Renderer = new ToolStripSystemRenderer()
        };
        toolStrip.Items.Add(CreateButton("适应窗口", (_, _) => FitImage()));
        toolStrip.Items.Add(CreateButton("100%", (_, _) => SetZoom(1F)));
        toolStrip.Items.Add(new ToolStripSeparator());
        toolStrip.Items.Add(CreateButton("－", (_, _) => ChangeZoom(1F / 1.2F)));
        toolStrip.Items.Add(CreateButton("＋", (_, _) => ChangeZoom(1.2F)));
        _zoomLabel = new ToolStripLabel("100%") { Margin = new Padding(8, 1, 8, 2) };
        toolStrip.Items.Add(_zoomLabel);
        toolStrip.Items.Add(new ToolStripSeparator());
        toolStrip.Items.Add(new ToolStripLabel($"原图 {_image.Width} × {_image.Height}px"));

        _viewport = new Panel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            BackColor = Color.FromArgb(28, 28, 30),
            Padding = new Padding(20)
        };
        _pictureBox = new PictureBox
        {
            Image = _image,
            SizeMode = PictureBoxSizeMode.Zoom,
            BackColor = Color.Transparent,
            TabStop = false
        };
        _viewport.Controls.Add(_pictureBox);

        Controls.Add(_viewport);
        Controls.Add(toolStrip);
        toolStrip.Dock = DockStyle.Top;

        Shown += (_, _) => FitImage();
        Resize += (_, _) => { if (_fitToWindow) FitImage(); };
        MouseWheel += OnPreviewMouseWheel;
        _viewport.MouseWheel += OnPreviewMouseWheel;
        _pictureBox.MouseWheel += OnPreviewMouseWheel;
        KeyDown += (_, e) =>
        {
            if (e.KeyCode == Keys.Escape) Close();
            if (e.KeyCode == Keys.Add || e.KeyCode == Keys.Oemplus) ChangeZoom(1.2F);
            if (e.KeyCode == Keys.Subtract || e.KeyCode == Keys.OemMinus) ChangeZoom(1F / 1.2F);
        };
    }

    public static void ShowPreview(Image image, string host, IWin32Window? owner = null)
    {
        using var form = new ImagePreviewForm(image, host);
        if (owner is null) form.ShowDialog();
        else form.ShowDialog(owner);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _pictureBox.Image = null;
            _image.Dispose();
        }
        base.Dispose(disposing);
    }

    private static ToolStripButton CreateButton(string text, EventHandler onClick)
    {
        var button = new ToolStripButton(text) { DisplayStyle = ToolStripItemDisplayStyle.Text };
        button.Click += onClick;
        return button;
    }

    private void OnPreviewMouseWheel(object? sender, MouseEventArgs e) =>
        ChangeZoom(e.Delta > 0 ? 1.15F : 1F / 1.15F);

    private void ChangeZoom(float factor) => SetZoom(_zoom * factor);

    private void FitImage()
    {
        if (_viewport.ClientSize.Width <= 40 || _viewport.ClientSize.Height <= 40) return;
        var widthRatio = (_viewport.ClientSize.Width - 40F) / _image.Width;
        var heightRatio = (_viewport.ClientSize.Height - 40F) / _image.Height;
        _fitToWindow = true;
        ApplyZoom(Math.Min(widthRatio, heightRatio));
    }

    private void SetZoom(float zoom)
    {
        _fitToWindow = false;
        ApplyZoom(zoom);
    }

    private void ApplyZoom(float zoom)
    {
        _zoom = Math.Max(MinZoom, Math.Min(MaxZoom, zoom));
        _pictureBox.Size = new Size(
            Math.Max(1, (int)Math.Round(_image.Width * _zoom)),
            Math.Max(1, (int)Math.Round(_image.Height * _zoom)));
        _pictureBox.Location = new Point(
            Math.Max(20, (_viewport.ClientSize.Width - _pictureBox.Width) / 2),
            Math.Max(20, (_viewport.ClientSize.Height - _pictureBox.Height) / 2));
        _zoomLabel.Text = $"{_zoom:P0}";
        _viewport.AutoScrollMinSize = new Size(_pictureBox.Right + 20, _pictureBox.Bottom + 20);
    }
}
