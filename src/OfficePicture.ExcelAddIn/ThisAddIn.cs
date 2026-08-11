using Excel = Microsoft.Office.Interop.Excel;
using OfficePicture.Core;

namespace OfficePicture.ExcelAddIn;

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

    private void Application_WindowSelectionChange(Excel.Range target)
    {
        try
        {
            var selection = Application.Selection;
            if (selection is not Excel.Shape shape) return;
            if (shape.Type != Microsoft.Office.Core.MsoShapeType.msoPicture && shape.Type != Microsoft.Office.Core.MsoShapeType.msoLinkedPicture) return;
            if (ClipboardImageCapture.TryCapture(shape.Copy, out var image) && image is not null)
            {
                using (image) _previewPane?.ShowImage(image, "Excel", "当前选中图片");
            }
        }
        catch { /* Excel can return a transient selection during mouse interaction. */ }
    }
}
