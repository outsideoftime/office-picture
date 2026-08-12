using System;
using System.Windows.Forms;

namespace OfficePicture.Core;

public sealed class NativeWindowOwner : IWin32Window
{
    public NativeWindowOwner(IntPtr handle) => Handle = handle;
    public IntPtr Handle { get; }
}
