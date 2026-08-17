using OfficePicture.Core;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.IO;

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
        {
            if (TryGetPackagePath(document, out var packagePath))
            {
                if (TryGetInlineShapeIdentity(document, inlineOrdinal, out var anchorId, out var editId) &&
                    OpenXmlImageExtractor.TryExtractWordImageByIdentityFromPackage(
                        packagePath,
                        anchorId,
                        editId,
                        out image))
                    return true;

                // Ordinals are only safe when the package and the in-memory document
                // are in sync. Unsaved insertions/deletions can shift every later image.
                if (document.Saved &&
                    OpenXmlImageExtractor.TryExtractWordImageByOrdinalFromPackage(
                        packagePath,
                        inlineOrdinal,
                        out image))
                    return true;
            }

            return OpenXmlImageExtractor.TryExtractWordImageByOrdinal(
                document.WordOpenXML,
                inlineOrdinal,
                out image);
        }

        return OpenXmlImageExtractor.TryExtractWordImage(
            document.WordOpenXML,
            hitShape.ID,
            hitShape.Name,
            out image);
    }

    private static bool TryGetPackagePath(Word.Document document, out string packagePath)
    {
        packagePath = string.Empty;
        try
        {
            if (!File.Exists(document.FullName)) return false;
            packagePath = document.FullName;
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static bool TryGetInlineShapeIdentity(
        Word.Document document,
        int inlineOrdinal,
        out string anchorId,
        out string editId)
    {
        anchorId = string.Empty;
        editId = string.Empty;
        Word.InlineShape? inlineShape = null;
        try
        {
            inlineShape = document.InlineShapes[inlineOrdinal];
            var anchorValue = unchecked((uint)inlineShape.AnchorID);
            var editValue = unchecked((uint)inlineShape.EditID);
            if (anchorValue != 0) anchorId = anchorValue.ToString("X8");
            if (editValue != 0) editId = editValue.ToString("X8");
            return anchorId.Length > 0 || editId.Length > 0;
        }
        catch
        {
            return false;
        }
        finally
        {
            if (inlineShape is not null)
                Marshal.FinalReleaseComObject(inlineShape);
        }
    }

    private static int GetInlineShapeOrdinal(Word.Document document, Word.Range hitAnchor)
    {
        if (hitAnchor.StoryType != Word.WdStoryType.wdMainTextStory) return -1;

        Word.Range? prefixRange = null;
        try
        {
            prefixRange = document.Range(0, hitAnchor.End);
            return prefixRange.InlineShapes.Count;
        }
        catch
        {
            return -1;
        }
        finally
        {
            if (prefixRange is not null)
                Marshal.FinalReleaseComObject(prefixRange);
        }
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
