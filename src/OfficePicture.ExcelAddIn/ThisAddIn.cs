using Excel = Microsoft.Office.Interop.Excel;
using OfficePicture.Core;

namespace OfficePicture.ExcelAddIn;

public partial class ThisAddIn
{
    private PreviewPane? _previewPane;
    private Microsoft.Office.Tools.CustomTaskPane? _taskPane;
    private System.Windows.Forms.Timer? _selectionTimer;

    private void ThisAddIn_Startup(object sender, System.EventArgs e)
    {
        _previewPane = new PreviewPane();
        _taskPane = CustomTaskPanes.Add(_previewPane, "图片预览");
        _taskPane.Width = 360;
        _taskPane.Visible = true;
        _selectionTimer = new System.Windows.Forms.Timer { Interval = 300 };
        _selectionTimer.Tick += SelectionTimer_Tick;
        _selectionTimer.Start();
    }

    private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
    {
        if (_selectionTimer is not null)
        {
            _selectionTimer.Stop();
            _selectionTimer.Tick -= SelectionTimer_Tick;
            _selectionTimer.Dispose();
        }
        _taskPane?.Dispose();
    }

    private void SelectionTimer_Tick(object? sender, System.EventArgs e)
    {
        try
        {
            var selection = Application.Selection;
            if (selection is not Excel.Shape shape) return;
            if (ClipboardImageCapture.TryCapture(shape.Copy, out var image) && image is not null)
            {
                using (image) _previewPane?.ShowImage(image, "Excel", "当前选中图片");
            }
        }
        catch { /* Excel can return a transient selection during mouse interaction. */ }
    }
}
