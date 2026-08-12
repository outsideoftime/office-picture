using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace OfficePicture.Core;

public sealed class OfficeDoubleClickHook : IDisposable
{
    private const int WhGetMessage = 3;
    private const uint WmLButtonDblClk = 0x0203;

    private readonly Action _onDoubleClick;
    private readonly string _requiredWindowClass;
    private readonly HookProc _hookProc;
    private readonly Control _dispatcher;
    private IntPtr _hook;
    private bool _callbackPending;

    public OfficeDoubleClickHook(string requiredWindowClass, Action onDoubleClick)
    {
        _requiredWindowClass = requiredWindowClass;
        _onDoubleClick = onDoubleClick;
        _hookProc = HookCallback;
        _dispatcher = new Control();
        _ = _dispatcher.Handle;
        _hook = SetWindowsHookEx(WhGetMessage, _hookProc, IntPtr.Zero, GetCurrentThreadId());
        if (_hook == IntPtr.Zero)
        {
            _dispatcher.Dispose();
            throw new InvalidOperationException("无法监听 Office 双击消息。");
        }
    }

    public void Dispose()
    {
        if (_hook != IntPtr.Zero)
        {
            UnhookWindowsHookEx(_hook);
            _hook = IntPtr.Zero;
        }
        _dispatcher.Dispose();
    }

    private IntPtr HookCallback(int code, IntPtr wParam, IntPtr lParam)
    {
        if (code >= 0 && !_callbackPending)
        {
            var message = Marshal.PtrToStructure<NativeMessage>(lParam);
            if (message.Message == WmLButtonDblClk && IsInsideRequiredWindow(message.HWnd))
            {
                _callbackPending = true;
                _dispatcher.BeginInvoke((Action)(() =>
                {
                    try { _onDoubleClick(); }
                    finally { _callbackPending = false; }
                }));
            }
        }
        return CallNextHookEx(_hook, code, wParam, lParam);
    }

    private bool IsInsideRequiredWindow(IntPtr window)
    {
        var className = new StringBuilder(128);
        while (window != IntPtr.Zero)
        {
            className.Clear();
            GetClassName(window, className, className.Capacity);
            if (className.ToString().IndexOf(_requiredWindowClass, StringComparison.OrdinalIgnoreCase) >= 0)
                return true;
            window = GetParent(window);
        }
        return false;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeMessage
    {
        public IntPtr HWnd;
        public uint Message;
        public UIntPtr WParam;
        public IntPtr LParam;
        public uint Time;
        public System.Drawing.Point Point;
    }

    private delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, HookProc callback, IntPtr module, uint threadId);
    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnhookWindowsHookEx(IntPtr hook);
    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr wParam, IntPtr lParam);
    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")]
    private static extern int GetClassName(IntPtr window, StringBuilder className, int maxCount);
    [DllImport("user32.dll")]
    private static extern IntPtr GetParent(IntPtr window);
}
