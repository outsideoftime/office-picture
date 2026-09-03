using OfficePicture.Core;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System;
using System.Diagnostics;
using System.IO;

namespace OfficePicture.PowerPointAddIn;

public partial class ThisAddIn
{
    private const string PictureContextMenuName = "Pictures Context Menu";
    private const string PicturePreviewButtonTag = "OfficePicture.PowerPoint.PicturePreview";

    private bool _previewOpen;
    private System.DateTime _suppressPreviewUntilUtc;
    private Office.CommandBarButton? _picturePreviewButton;

    private void ThisAddIn_Startup(object sender, System.EventArgs e)
    {
        Application.WindowBeforeDoubleClick += Application_WindowBeforeDoubleClick;
        InstallPictureContextMenu();
    }

    private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
    {
        Application.WindowBeforeDoubleClick -= Application_WindowBeforeDoubleClick;
        RemovePictureContextMenu();
    }

    private void Application_WindowBeforeDoubleClick(PowerPoint.Selection selection, ref bool cancel)
    {
        if (_previewOpen || System.DateTime.UtcNow < _suppressPreviewUntilUtc)
        {
            cancel = true;
            return;
        }

        try
        {
            // PowerPoint can briefly report ppSelectionNone/another selection
            // type while it is completing the second click. ShapeRange is the
            // more reliable source for the object under the pointer.
            if (!TryGetPictureShape(selection, out var shape) || shape is null) return;
            cancel = true;
            PreviewPicture(selection, shape);
        }
        catch { /* Office selection can be transient during a double-click. */ }
    }

    private void PicturePreviewButton_Click(Office.CommandBarButton control, ref bool cancelDefault)
    {
        cancelDefault = true;
        PreviewActivePicture();
    }

    internal void PreviewActivePicture()
    {
        try
        {
            var selection = Application.ActiveWindow.Selection;
            if (!TryGetPictureShape(selection, out var shape) || shape is null) return;
            PreviewPicture(selection, shape);
        }
        catch { /* The active window can change while the context menu is closing. */ }
    }

    private void PreviewPicture(PowerPoint.Selection selection, PowerPoint.Shape shape)
    {
        if (_previewOpen || System.DateTime.UtcNow < _suppressPreviewUntilUtc) return;

        _previewOpen = true;
        try
        {
            if (!TryGetOriginalImage(selection, shape, out var image) || image is null) return;
            using (image)
                ImagePreviewForm.ShowPreview(image, "PowerPoint", GetPowerPointOwner());
        }
        finally
        {
            _previewOpen = false;
            _suppressPreviewUntilUtc = System.DateTime.UtcNow.AddMilliseconds(400);
        }
    }

    private void InstallPictureContextMenu()
    {
        try
        {
            var menu = Application.CommandBars[PictureContextMenuName];
            if (menu is null) return;

            foreach (Office.CommandBarControl control in menu.Controls)
            {
                if (!string.Equals(control.Tag, PicturePreviewButtonTag, StringComparison.Ordinal)) continue;
                if (control is Office.CommandBarButton existingButton)
                {
                    _picturePreviewButton = existingButton;
                    _picturePreviewButton.Click += PicturePreviewButton_Click;
                }
                return;
            }

            var addedControl = menu.Controls.Add(
                Office.MsoControlType.msoControlButton,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                true);
            _picturePreviewButton = (Office.CommandBarButton)addedControl;
            _picturePreviewButton.Caption = "图片预览";
            _picturePreviewButton.DescriptionText = "预览当前 PowerPoint 图片";
            _picturePreviewButton.TooltipText = "预览图片";
            _picturePreviewButton.Tag = PicturePreviewButtonTag;
            _picturePreviewButton.BeginGroup = true;
            _picturePreviewButton.Visible = true;
            _picturePreviewButton.Enabled = true;
            _picturePreviewButton.Click += PicturePreviewButton_Click;
        }
        catch
        {
            // Some PowerPoint builds or menu customizers may not expose the
            // legacy CommandBars collection. Double-click remains available.
        }
    }

    private void RemovePictureContextMenu()
    {
        try
        {
            if (_picturePreviewButton is null) return;
            _picturePreviewButton.Click -= PicturePreviewButton_Click;
            _picturePreviewButton.Delete();
        }
        catch { }
        finally
        {
            _picturePreviewButton = null;
        }
    }

    private bool TryGetOriginalImage(
        PowerPoint.Selection selection,
        PowerPoint.Shape shape,
        out System.Drawing.Image? image)
    {
        image = null;
        string? temporaryCopy = null;
        try
        {
            try
            {
                var presentation = Application.ActivePresentation;
                var packagePath = presentation.FullName;
                if (presentation.Saved != Office.MsoTriState.msoTrue ||
                    !File.Exists(packagePath) ||
                    !IsOpenXmlPresentation(packagePath))
                {
                    temporaryCopy = Path.Combine(Path.GetTempPath(), $"OfficePicture-{Guid.NewGuid():N}.pptx");
                    presentation.SaveCopyAs(
                        temporaryCopy,
                        PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                        Office.MsoTriState.msoFalse);
                    packagePath = temporaryCopy;
                }

                if (OpenXmlImageExtractor.TryExtractPowerPointImage(
                    packagePath,
                    GetSlideIndex(selection),
                    shape.Id,
                    shape.Name,
                    out image))
                    return true;
            }
            catch
            {
                // Fall through to the host copy path below. PowerPoint's COM
                // selection and SaveCopyAs can be transient during a double-click.
            }

            // The package path is preferred because it preserves the original
            // pixels. Copying is a compatibility fallback for legacy/unsaved
            // presentations and shapes that cannot be matched in the package.
            return ClipboardImageCapture.TryCapture(shape.Copy, out image) && image is not null;
        }
        finally
        {
            if (temporaryCopy is not null)
            {
                try { File.Delete(temporaryCopy); }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
            }
        }
    }

    private static bool TryGetPictureShape(
        PowerPoint.Selection selection,
        out PowerPoint.Shape? shape)
    {
        shape = null;
        try
        {
            var shapeRange = selection.ShapeRange;
            if (shapeRange is null || shapeRange.Count == 0) return false;

            var selectedShape = shapeRange[1];
            if (!IsPicture(selectedShape.Type)) return false;

            shape = selectedShape;
            return true;
        }
        catch
        {
            return false;
        }
    }

    private int GetSlideIndex(PowerPoint.Selection selection)
    {
        try
        {
            var slideRange = selection.SlideRange;
            if (slideRange is not null && slideRange.Count > 0)
                return slideRange[1].SlideIndex;
        }
        catch { }

        return 0;
    }

    private static NativeWindowOwner? GetPowerPointOwner()
    {
        try
        {
            var handle = Process.GetCurrentProcess().MainWindowHandle;
            return handle == IntPtr.Zero ? null : new NativeWindowOwner(handle);
        }
        catch
        {
            return null;
        }
    }

    private static bool IsOpenXmlPresentation(string path)
    {
        var extension = Path.GetExtension(path);
        return extension.Equals(".pptx", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".pptm", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".ppsx", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".ppsm", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".potx", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".potm", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsPicture(Office.MsoShapeType type)
    {
        var value = (int)type;
        return value == 13 || value == 11 || value == 28 || value == 29;
    }
}
