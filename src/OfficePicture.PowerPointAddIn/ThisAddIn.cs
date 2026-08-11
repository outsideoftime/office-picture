using OfficePicture.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OfficePicture.PowerPointAddIn;

public partial class ThisAddIn
{
    private PreviewPane? _previewPane;
    private Microsoft.Office.Tools.CustomTaskPane? _taskPane;

    private void ThisAddIn_Startup(object sender, System.EventArgs e)
    {
        _previewPane = new PreviewPane();
        _taskPane = CustomTaskPanes.Add(_previewPane, "图片预览");
        _taskPane.Width = 360;
        _taskPane.Visible = true;
        Application.WindowSelectionChange += Application_WindowSelectionChange;
    }

    private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
    {
        Application.WindowSelectionChange -= Application_WindowSelectionChange;
        _taskPane?.Dispose();
    }

    private void Application_WindowSelectionChange(PowerPoint.Selection selection)
    {
        try
        {
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes || selection.ShapeRange.Count == 0) return;
            var shape = selection.ShapeRange[1];
            if (ClipboardImageCapture.TryCapture(shape.Copy, out var image) && image is not null)
            {
                using (image) _previewPane?.ShowImage(image, "PowerPoint", "当前选中图片");
            }
        }
        catch { /* Office selection can be transient while changing. */ }
    }
}
