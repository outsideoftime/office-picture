using Excel = Microsoft.Office.Interop.Excel;
using OfficePicture.Core;
using Office = Microsoft.Office.Core;
using System;
using System.IO;
using System.Reflection;

namespace OfficePicture.ExcelAddIn;

public partial class ThisAddIn
{
    private OfficeDoubleClickHook? _doubleClickHook;
    private bool _previewOpen;
    private System.DateTime _suppressPreviewUntilUtc;

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
        if (_previewOpen || System.DateTime.UtcNow < _suppressPreviewUntilUtc) return;
        try
        {
            object selection = Application.Selection;
            if (selection is null || selection is Excel.Range) return;
            if (!TryGetPictureSelection(selection, out var sheet, out var shape) || sheet is null || shape is null) return;
            _previewOpen = true;
            try
            {
                if (!TryGetOriginalImage(sheet, shape, out var image) || image is null) return;
                using (image)
                    ImagePreviewForm.ShowPreview(image, "Excel", new NativeWindowOwner(new System.IntPtr(Application.Hwnd)));
            }
            finally
            {
                _previewOpen = false;
                _suppressPreviewUntilUtc = System.DateTime.UtcNow.AddMilliseconds(400);
            }
        }
        catch { /* Excel can return a transient selection during a double-click. */ }
    }

    private bool TryGetPictureSelection(
        object selection,
        out Excel.Worksheet? sheet,
        out Excel.Shape? shape)
    {
        sheet = null;
        shape = null;
        try
        {
            var name = selection.GetType().InvokeMember(
                "Name", BindingFlags.GetProperty, null, selection, null);
            object activeSheet = Application.ActiveSheet;
            if (name is null || activeSheet is not Excel.Worksheet worksheet) return false;
            var selectedShape = worksheet.Shapes.Item(name);
            var value = (int)selectedShape.Type;
            if (value != 13 && value != 11 && value != 28 && value != 29) return false;

            sheet = worksheet;
            shape = selectedShape;
            return true;
        }
        catch { return false; }
    }

    private bool TryGetOriginalImage(
        Excel.Worksheet sheet,
        Excel.Shape shape,
        out System.Drawing.Image? image)
    {
        image = null;
        string? temporaryCopy = null;
        try
        {
            var workbook = Application.ActiveWorkbook;
            var extension = GetOpenXmlWorkbookExtension(workbook.FileFormat);
            if (extension is null) return false;

            var packagePath = workbook.FullName;
            if (!workbook.Saved || !File.Exists(packagePath) || !IsOpenXmlWorkbook(packagePath))
            {
                temporaryCopy = Path.Combine(Path.GetTempPath(), $"OfficePicture-{Guid.NewGuid():N}{extension}");
                workbook.SaveCopyAs(temporaryCopy);
                packagePath = temporaryCopy;
            }

            return OpenXmlImageExtractor.TryExtractExcelImage(
                packagePath,
                sheet.Name,
                shape.ID,
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

    private static string? GetOpenXmlWorkbookExtension(Excel.XlFileFormat format)
    {
        switch ((int)format)
        {
            case 51: return ".xlsx";
            case 52: return ".xlsm";
            case 53: return ".xltm";
            case 54: return ".xltx";
            default: return null;
        }
    }

    private static bool IsOpenXmlWorkbook(string path)
    {
        var extension = Path.GetExtension(path);
        return extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".xltx", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".xltm", StringComparison.OrdinalIgnoreCase);
    }

}
