using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OfficePicture.Core;

public sealed class ImagePreviewForm : Form
{
    private const float MinZoom = 0.1F;
    private const float MaxZoom = 8F;

    private readonly Panel _viewport;
    private readonly PixelAccuratePictureBox _pictureBox;
    private readonly ToolStripLabel _zoomLabel;
    private readonly Image _image;
    private float _zoom = 1F;
    private bool _fitToWindow = true;

    public ImagePreviewForm(Image image, string host)
    {
        _image = new Bitmap(image);
        Text = $"图片预览 - {host}";
        StartPosition = FormStartPosition.Manual;
        Size = new Size(1000, 720);
        MinimumSize = new Size(480, 360);
        BackColor = Color.FromArgb(28, 28, 30);
        ForeColor = Color.White;
        AutoScaleMode = AutoScaleMode.None;
        KeyPreview = true;
        ShowIcon = false;
        ShowInTaskbar = false;
        FormBorderStyle = FormBorderStyle.None;
        Padding = new Padding(1);

        var toolStrip = new ToolStrip
        {
            GripStyle = ToolStripGripStyle.Hidden,
            BackColor = Color.FromArgb(38, 38, 41),
            ForeColor = Color.White,
            Padding = new Padding(12, 6, 12, 6),
            Height = 44,
            Renderer = new ToolStripSystemRenderer()
        };
        toolStrip.Items.Add(new ToolStripLabel($"图片预览 · {host}")
        {
            Font = new Font(SystemFonts.MessageBoxFont, FontStyle.Bold),
            ForeColor = Color.White,
            Margin = new Padding(0, 1, 16, 2)
        });
        toolStrip.Items.Add(CreateButton("适应窗口", (_, _) => FitImage()));
        toolStrip.Items.Add(CreateButton("100%", (_, _) => SetZoom(1F)));
        toolStrip.Items.Add(new ToolStripSeparator());
        toolStrip.Items.Add(CreateButton("－", (_, _) => ChangeZoom(1F / 1.2F)));
        toolStrip.Items.Add(CreateButton("＋", (_, _) => ChangeZoom(1.2F)));
        _zoomLabel = new ToolStripLabel("100%") { Margin = new Padding(8, 1, 8, 2) };
        toolStrip.Items.Add(_zoomLabel);
        toolStrip.Items.Add(new ToolStripSeparator());
        var closeButton = CreateButton("关闭  ✕", (_, _) => Close());
        closeButton.Alignment = ToolStripItemAlignment.Right;
        toolStrip.Items.Add(closeButton);
        var sizeLabel = new ToolStripLabel($"原始媒体 {_image.Width} × {_image.Height}px")
        {
            Alignment = ToolStripItemAlignment.Right,
            ForeColor = Color.Gainsboro
        };
        toolStrip.Items.Add(sizeLabel);

        _viewport = new Panel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            BackColor = Color.FromArgb(28, 28, 30)
        };
        _pictureBox = new PixelAccuratePictureBox
        {
            Image = _image,
            BackColor = Color.FromArgb(28, 28, 30),
            TabStop = false
        };
        _viewport.Controls.Add(_pictureBox);

        Controls.Add(_viewport);
        Controls.Add(toolStrip);
        toolStrip.Dock = DockStyle.Top;

        Shown += (_, _) => FitImage();
        Resize += (_, _) =>
        {
            if (_fitToWindow) FitImage();
            else ApplyZoom(_zoom, GetViewportCenter());
        };
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
        form.PlaceOverOwner(owner);
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
        var button = new ToolStripButton(text)
        {
            DisplayStyle = ToolStripItemDisplayStyle.Text,
            ForeColor = Color.White,
            AutoSize = true,
            Margin = new Padding(2, 1, 2, 2)
        };
        button.Click += onClick;
        return button;
    }

    private void PlaceOverOwner(IWin32Window? owner)
    {
        const int inset = 12;
        if (owner is not null && owner.Handle != IntPtr.Zero && GetWindowRect(owner.Handle, out var ownerRect))
        {
            Bounds = new Rectangle(
                ownerRect.Left + inset,
                ownerRect.Top + inset,
                Math.Max(MinimumSize.Width, ownerRect.Right - ownerRect.Left - inset * 2),
                Math.Max(MinimumSize.Height, ownerRect.Bottom - ownerRect.Top - inset * 2));
            return;
        }

        var area = Screen.FromPoint(Cursor.Position).WorkingArea;
        Bounds = new Rectangle(
            area.Left + Math.Max(0, (area.Width - Width) / 2),
            area.Top + Math.Max(0, (area.Height - Height) / 2),
            Math.Min(Width, area.Width),
            Math.Min(Height, area.Height));
    }

    private void OnPreviewMouseWheel(object? sender, MouseEventArgs e)
    {
        if (e is HandledMouseEventArgs handled) handled.Handled = true;
        var focus = _viewport.PointToClient(Cursor.Position);
        ChangeZoom(e.Delta > 0 ? 1.15F : 1F / 1.15F, focus);
    }

    private void ChangeZoom(float factor, Point? focus = null) =>
        SetZoom(_zoom * factor, focus);

    private void FitImage()
    {
        if (_viewport.ClientSize.Width <= 40 || _viewport.ClientSize.Height <= 40) return;
        var widthRatio = (_viewport.ClientSize.Width - 40F) / _image.Width;
        var heightRatio = (_viewport.ClientSize.Height - 40F) / _image.Height;
        _fitToWindow = true;
        ApplyZoom(Math.Min(widthRatio, heightRatio), resetScroll: true);
    }

    private void SetZoom(float zoom, Point? focus = null)
    {
        _fitToWindow = false;
        ApplyZoom(zoom, focus ?? GetViewportCenter());
    }

    private void ApplyZoom(float zoom, Point? focus = null, bool resetScroll = false)
    {
        var focalPoint = focus ?? GetViewportCenter();
        var imageFocusX = _pictureBox.Width > 0
            ? Math.Max(0F, Math.Min(1F, (focalPoint.X - _pictureBox.Left) / (float)_pictureBox.Width))
            : 0.5F;
        var imageFocusY = _pictureBox.Height > 0
            ? Math.Max(0F, Math.Min(1F, (focalPoint.Y - _pictureBox.Top) / (float)_pictureBox.Height))
            : 0.5F;

        _zoom = Math.Max(MinZoom, Math.Min(MaxZoom, zoom));
        var scaledSize = new Size(
            Math.Max(1, (int)Math.Round(_image.Width * _zoom)),
            Math.Max(1, (int)Math.Round(_image.Height * _zoom)));

        _viewport.SuspendLayout();
        _viewport.AutoScrollPosition = Point.Empty;
        _pictureBox.Size = scaledSize;
        _pictureBox.Location = GetPictureLocation(scaledSize);

        // Right/Bottom include the current negative scroll offset. Computing the
        // virtual canvas from them makes the far edge unreachable after zooming
        // while already scrolled. The image dimensions are stable virtual values.
        var virtualSize = new Size(scaledSize.Width + 40, scaledSize.Height + 40);
        _viewport.AutoScrollMinSize = virtualSize;

        if (!resetScroll)
        {
            var targetX = _pictureBox.Left + (int)Math.Round(imageFocusX * scaledSize.Width) - focalPoint.X;
            var targetY = _pictureBox.Top + (int)Math.Round(imageFocusY * scaledSize.Height) - focalPoint.Y;
            var maxX = Math.Max(0, virtualSize.Width - _viewport.ClientSize.Width);
            var maxY = Math.Max(0, virtualSize.Height - _viewport.ClientSize.Height);
            _viewport.AutoScrollPosition = new Point(
                Math.Max(0, Math.Min(maxX, targetX)),
                Math.Max(0, Math.Min(maxY, targetY)));
        }

        _zoomLabel.Text = $"{_zoom:P0}";
        _viewport.ResumeLayout(true);
        _viewport.Invalidate(true);
        _viewport.Update();
    }

    private Point GetViewportCenter() =>
        new(_viewport.ClientSize.Width / 2, _viewport.ClientSize.Height / 2);

    private Point GetPictureLocation(Size imageSize) => new(
        Math.Max(20, (_viewport.ClientSize.Width - imageSize.Width) / 2),
        Math.Max(20, (_viewport.ClientSize.Height - imageSize.Height) / 2));

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeRect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetWindowRect(IntPtr window, out NativeRect rect);

    private sealed class PixelAccuratePictureBox : PictureBox
    {
        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.Clear(BackColor);
            if (Image is null) return;

            if (ClientSize.Width == Image.Width && ClientSize.Height == Image.Height)
            {
                e.Graphics.DrawImageUnscaled(Image, Point.Empty);
                return;
            }

            e.Graphics.CompositingMode = CompositingMode.SourceCopy;
            e.Graphics.CompositingQuality = CompositingQuality.HighQuality;
            e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            e.Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
            e.Graphics.DrawImage(Image, ClientRectangle);
        }
    }
}
