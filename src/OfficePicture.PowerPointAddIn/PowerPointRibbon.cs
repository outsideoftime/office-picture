using System;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace OfficePicture.PowerPointAddIn;

[ComVisible(true)]
public sealed class PowerPointRibbon : Office.IRibbonExtensibility
{
    private const string PictureContextMenuXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
        "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">" +
        "<contextMenus>" +
        "<contextMenu idMso=\"ContextMenuPicture\">" +
        "<button id=\"OfficePicturePowerPointPreviewRibbon\" " +
        "label=\"图片预览\" " +
        "screentip=\"预览图片\" " +
        "supertip=\"预览当前 PowerPoint 图片\" " +
        "imageMso=\"PictureInsertFromFile\" " +
        "onAction=\"OnPicturePreview\"/>" +
        "</contextMenu>" +
        "</contextMenus>" +
        "</customUI>";

    private readonly ThisAddIn _addIn;

    public PowerPointRibbon(ThisAddIn addIn)
    {
        _addIn = addIn;
    }

    public string GetCustomUI(string ribbonId)
    {
        return PictureContextMenuXml;
    }

    public void OnPicturePreview(Office.IRibbonControl control)
    {
        try
        {
            _addIn.PreviewActivePicture();
        }
        catch
        {
            // The active selection can change while the context menu closes.
        }
    }
}
