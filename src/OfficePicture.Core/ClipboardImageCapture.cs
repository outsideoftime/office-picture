using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OfficePicture.Core;

public static class ClipboardImageCapture
{
    public static bool TryCapture(Action copySelection, out Image? image)
    {
        image = null;
        try
        {
            copySelection();
            Application.DoEvents();

            if (!Clipboard.ContainsImage())
            {
                return false;
            }

            var clipboardImage = Clipboard.GetImage();
            if (clipboardImage is null)
            {
                return false;
            }

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
    }
}
