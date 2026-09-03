using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OfficePicture.Core;

/// <summary>
/// Captures an Office selection through the clipboard and restores the
/// clipboard contents afterwards. This is used as a fallback when a host's
/// Open XML package cannot be read or the selected object cannot be matched.
/// </summary>
public static class ClipboardImageCapture
{
    public static bool TryCapture(Action copySelection, out Image? image)
    {
        image = null;
        IDataObject? previousClipboard = null;
        try
        {
            try { previousClipboard = SnapshotClipboard(); }
            catch (ExternalException) { }

            copySelection();
            Application.DoEvents();

            if (!Clipboard.ContainsImage()) return false;

            var clipboardImage = Clipboard.GetImage();
            if (clipboardImage is null) return false;

            image = new Bitmap(clipboardImage);
            return true;
        }
        catch (ExternalException)
        {
            return false;
        }
        catch (InvalidOperationException)
        {
            return false;
        }
        finally
        {
            if (previousClipboard is not null)
            {
                try { Clipboard.SetDataObject(previousClipboard, true); }
                catch (ExternalException) { }
            }
        }
    }

    private static IDataObject? SnapshotClipboard()
    {
        var source = Clipboard.GetDataObject();
        if (source is null) return null;

        var snapshot = new DataObject();
        foreach (var format in source.GetFormats())
        {
            try
            {
                var data = source.GetData(format);
                if (data is not null) snapshot.SetData(format, data);
            }
            catch (ExternalException) { }
        }
        return snapshot;
    }
}
