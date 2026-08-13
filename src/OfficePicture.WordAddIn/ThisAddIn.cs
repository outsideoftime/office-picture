using OfficePicture.Core;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace OfficePicture.WordAddIn;

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

    private void Application_WindowBeforeDoubleClick(Word.Selection selection, ref bool cancel)
    {
        if (_previewOpen || System.DateTime.UtcNow < _suppressPreviewUntilUtc)
        {
            cancel = true;
            return;
        }

        try
        {
            if (!IsPicture(selection)) return;

            cancel = true;
            _previewOpen = true;
            try
            {
                if (!TryGetOriginalImage(selection, out var image) || image is null) return;
                using (image)
                    ImagePreviewForm.ShowPreview(image, "Word", new NativeWindowOwner(new System.IntPtr(Application.ActiveWindow.Hwnd)));
            }
            finally
            {
                _previewOpen = false;
                _suppressPreviewUntilUtc = System.DateTime.UtcNow.AddMilliseconds(400);
            }
        }
        catch { /* Office selection can be transient during a double-click. */ }
    }

    private static bool TryGetOriginalImage(Word.Selection selection, out System.Drawing.Image? image)
    {
        image = null;
        if (selection.InlineShapes.Count > 0)
        {
            var selectedShape = selection.InlineShapes[1];
            var ordinal = GetInlineShapeOrdinal(selection.Document, selectedShape);
            if (ordinal < 1) return false;

            return OpenXmlImageExtractor.TryExtractWordImageByOrdinal(
                selection.Document.WordOpenXML,
                ordinal,
                out image);
        }

        if (selection.ShapeRange.Count == 0) return false;
        var shape = selection.ShapeRange[1];

        return OpenXmlImageExtractor.TryExtractWordImage(
            selection.Document.WordOpenXML,
            shape.ID,
            shape.Name,
            out image);
    }

    private static int GetInlineShapeOrdinal(Word.Document document, Word.InlineShape selectedShape)
    {
        var selectedRange = selectedShape.Range;
        for (var index = 1; index <= document.InlineShapes.Count; index++)
        {
            var candidateRange = document.InlineShapes[index].Range;
            if (candidateRange.Start == selectedRange.Start &&
                candidateRange.StoryType == selectedRange.StoryType)
                return index;
        }

        return -1;
    }

    private static bool IsPicture(Word.Selection selection)
    {
        if (selection.InlineShapes.Count > 0)
        {
            var type = selection.InlineShapes[1].Type;
            return type == Word.WdInlineShapeType.wdInlineShapePicture ||
                   type == Word.WdInlineShapeType.wdInlineShapeLinkedPicture;
        }

        try
        {
            if (selection.ShapeRange.Count == 0) return false;
            var type = selection.ShapeRange[1].Type;
            var value = (int)type;
            return value == 13 || value == 11 || value == 28 || value == 29;
        }
        catch { return false; }
    }
}
