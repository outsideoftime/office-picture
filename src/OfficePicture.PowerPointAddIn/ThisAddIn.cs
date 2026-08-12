using OfficePicture.Core;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OfficePicture.PowerPointAddIn;

public partial class ThisAddIn
{
    private bool _previewOpen;

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
        if (_previewOpen) return;
        try
        {
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes || selection.ShapeRange.Count == 0) return;
            var shape = selection.ShapeRange[1];
            if (!IsPicture(shape.Type)) return;
            cancel = true;
            if (!ClipboardImageCapture.TryCapture(shape.Copy, out var image) || image is null) return;

            _previewOpen = true;
            try
            {
                using (image)
                    ImagePreviewForm.ShowPreview(image, "PowerPoint", new NativeWindowOwner(new System.IntPtr(Application.HWND)));
            }
            finally { _previewOpen = false; }
        }
        catch { /* Office selection can be transient during a double-click. */ }
    }

    private static bool IsPicture(Office.MsoShapeType type)
    {
        var value = (int)type;
        return value == 13 || value == 11 || value == 28 || value == 29;
    }
}
