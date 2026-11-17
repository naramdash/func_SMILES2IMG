using System;
using System.Runtime.InteropServices;

namespace Smiles2Img.Editor;

/// <summary>
/// Thread-level Win32 WH_MOUSE hook that intercepts WM_LBUTTONDBLCLK
/// when Office native double-click events do not fire for floating shapes or 3D models.
/// </summary>
public sealed class ThreadMouseDoubleClickHook : IDisposable
{
    private const int WH_MOUSE = 7;
    private const int WM_LBUTTONDBLCLK = 0x0203;

    private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

    private readonly Func<bool> _onDoubleClick;
    private readonly HookProc _hookProc;
    private IntPtr _hookHandle;
    private bool _isDispatching;

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();

    /// <summary>
    /// Installs a thread-level double-click hook. The callback returns true to suppress the host's default handling.
    /// </summary>
    public ThreadMouseDoubleClickHook(Func<bool> onDoubleClick)
    {
        _onDoubleClick = onDoubleClick ?? throw new ArgumentNullException(nameof(onDoubleClick));
        _hookProc = Callback;
        try
        {
            _hookHandle = SetWindowsHookEx(WH_MOUSE, _hookProc, IntPtr.Zero, GetCurrentThreadId());
        }
        catch { }
    }

    private IntPtr Callback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && wParam == (IntPtr)WM_LBUTTONDBLCLK && !_isDispatching)
        {
            _isDispatching = true;
            try
            {
                if (_onDoubleClick())
                    return (IntPtr)1;
            }
            catch { }
            finally
            {
                _isDispatching = false;
            }
        }
        return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
    }

    public void Dispose()
    {
        if (_hookHandle != IntPtr.Zero)
        {
            try { UnhookWindowsHookEx(_hookHandle); } catch { }
            _hookHandle = IntPtr.Zero;
        }
    }
}
