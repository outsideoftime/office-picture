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
            var xml = selection.InlineShapes[1].Range.WordOpenXML;
            return OpenXmlImageExtractor.TryExtractWordImage(xml, null, null, out image);
        }

        if (selection.ShapeRange.Count == 0) return false;
        var shape = selection.ShapeRange[1];

        if (OpenXmlImageExtractor.TryExtractWordImage(
                selection.Range.WordOpenXML, shape.ID, shape.Name, out image))
            return true;

        var anchorParagraph = shape.Anchor.Paragraphs[1].Range;
        return OpenXmlImageExtractor.TryExtractWordImage(
            anchorParagraph.WordOpenXML, shape.ID, shape.Name, out image);
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
