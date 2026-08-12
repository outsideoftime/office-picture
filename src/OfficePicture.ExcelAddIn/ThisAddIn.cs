using Excel = Microsoft.Office.Interop.Excel;
using OfficePicture.Core;
using Office = Microsoft.Office.Core;
using System.Reflection;

namespace OfficePicture.ExcelAddIn;

public partial class ThisAddIn
{
    private OfficeDoubleClickHook? _doubleClickHook;
    private bool _previewOpen;

    private void ThisAddIn_Startup(object sender, System.EventArgs e)
    {
        _doubleClickHook = new OfficeDoubleClickHook("EXCEL7", PreviewSelectedPicture);
    }

    private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
    {
        _doubleClickHook?.Dispose();
        _doubleClickHook = null;
    }

    private void PreviewSelectedPicture()
    {
        if (_previewOpen) return;
        try
        {
            object selection = Application.Selection;
            if (selection is null || selection is Excel.Range) return;
            if (!IsPictureSelection(selection)) return;
            if (!ClipboardImageCapture.TryCapture(
                    () => Application.CommandBars.ExecuteMso("Copy"), out var image) || image is null) return;

            _previewOpen = true;
            try
            {
                using (image)
                    ImagePreviewForm.ShowPreview(image, "Excel", new NativeWindowOwner(new System.IntPtr(Application.Hwnd)));
            }
            finally { _previewOpen = false; }
        }
        catch { /* Excel can return a transient selection during a double-click. */ }
    }

    private bool IsPictureSelection(object selection)
    {
        try
        {
            var name = selection.GetType().InvokeMember(
                "Name", BindingFlags.GetProperty, null, selection, null);
            object activeSheet = Application.ActiveSheet;
            if (name is null || activeSheet is not Excel.Worksheet sheet) return false;
            var type = sheet.Shapes.Item(name).Type;
            var value = (int)type;
            return value == 13 || value == 11 || value == 28 || value == 29;
        }
        catch { return false; }
    }
}
