using OfficePicture.Core;
using Word = Microsoft.Office.Interop.Word;

namespace OfficePicture.WordAddIn;

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

    private void Application_WindowSelectionChange(Word.Selection selection)
    {
        try
        {
            var hasInlinePicture = selection.InlineShapes.Count > 0;
            if (!hasInlinePicture) { try { hasInlinePicture = selection.ShapeRange.Count > 0; } catch { } }
            if (!hasInlinePicture) return;
            if (ClipboardImageCapture.TryCapture(selection.Copy, out var image) && image is not null)
            {
                using (image) _previewPane?.ShowImage(image, "Word", "当前选中图片");
            }
        }
        catch { /* Office selection can be transient while changing. */ }
    }
}
