using OfficePicture.Core;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;

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
            _previewOpen = true;
            try
            {
                if (!TryGetImageUnderMouse(out var image) || image is null) return;
                cancel = true;
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

    private bool TryGetImageUnderMouse(out System.Drawing.Image? image)
    {
        image = null;
        if (!GetCursorPos(out var cursor)) return false;

        object hitObject = Application.ActiveWindow.RangeFromPoint(cursor.X, cursor.Y);
        if (hitObject is not Word.Shape hitShape || !IsPicture(hitShape.Type)) return false;

        var document = Application.ActiveDocument;
        var inlineOrdinal = GetInlineShapeOrdinal(document, hitShape.Anchor);
        if (inlineOrdinal > 0)
            return OpenXmlImageExtractor.TryExtractWordImageByOrdinal(
                document.WordOpenXML,
                inlineOrdinal,
                out image);

        return OpenXmlImageExtractor.TryExtractWordImage(
            document.WordOpenXML,
            hitShape.ID,
            hitShape.Name,
            out image);
    }

    private static int GetInlineShapeOrdinal(Word.Document document, Word.Range hitAnchor)
    {
        for (var index = 1; index <= document.InlineShapes.Count; index++)
        {
            var candidateRange = document.InlineShapes[index].Range;
            if (candidateRange.Start == hitAnchor.Start &&
                candidateRange.StoryType == hitAnchor.StoryType)
                return index;
        }

        return -1;
    }

    private static bool IsPicture(Office.MsoShapeType type)
    {
        var value = (int)type;
        return value == 13 || value == 11 || value == 28 || value == 29;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct CursorPoint
    {
        public int X;
        public int Y;
    }

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetCursorPos(out CursorPoint point);
}
