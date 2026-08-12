using OfficePicture.Core;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System;
using System.IO;

namespace OfficePicture.PowerPointAddIn;

public partial class ThisAddIn
{
    private bool _previewOpen;
    private System.DateTime _suppressPreviewUntilUtc;

    private void ThisAddIn_Startup(object sender, System.EventArgs e)
    {
        Application.WindowBeforeDoubleClick += Application_WindowBeforeDoubleClick;
    }

    private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
    {
        Application.WindowBeforeDoubleClick -= Application_WindowBeforeDoubleClick;
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
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes || selection.ShapeRange.Count == 0) return;
            var shape = selection.ShapeRange[1];
            if (!IsPicture(shape.Type)) return;
            cancel = true;
            _previewOpen = true;
            try
            {
                if (!TryGetOriginalImage(selection, shape, out var image) || image is null) return;
                using (image)
                    ImagePreviewForm.ShowPreview(image, "PowerPoint", new NativeWindowOwner(new System.IntPtr(Application.HWND)));
            }
            finally
            {
                _previewOpen = false;
                _suppressPreviewUntilUtc = System.DateTime.UtcNow.AddMilliseconds(400);
            }
        }
        catch { /* Office selection can be transient during a double-click. */ }
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

            return OpenXmlImageExtractor.TryExtractPowerPointImage(
                packagePath,
                selection.SlideRange[1].SlideIndex,
                shape.Id,
                shape.Name,
                out image);
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
