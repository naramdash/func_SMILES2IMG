using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace Smiles2Img.Editor;

/// <summary>
/// Thread-level Win32 WH_MOUSE hook that intercepts WM_LBUTTONDBLCLK
/// on Office canvas client windows when native double-click events do not fire for floating shapes or 3D models.
/// Never invokes modal UI synchronously inside the hook procedure.
/// </summary>
public sealed class ThreadMouseDoubleClickHook : IDisposable
{
    private const int WH_MOUSE = 7;
    private const int HC_ACTION = 0;
    private const int WM_LBUTTONDBLCLK = 0x0203;
    private const uint HTCLIENT = 1;
    private const int WM_APP_DISPATCH = 0x8000 + 0x201;
    private static readonly IntPtr HwndMessage = new(-3);

    private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT
    {
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MOUSEHOOKSTRUCT
    {
        public POINT pt;
        public IntPtr hwnd;
        public uint wHitTestCode;
        public IntPtr dwExtraInfo;
    }

    private sealed class DispatchWindow : NativeWindow
    {
        private readonly ConcurrentQueue<Action> _queue = new();
        private readonly Func<bool> _isBusy;
        private readonly Action<bool> _setBusy;

        public DispatchWindow(Func<bool> isBusy, Action<bool> setBusy)
        {
            _isBusy = isBusy;
            _setBusy = setBusy;
            CreateHandle(new CreateParams { Parent = HwndMessage });
        }

        public bool TryPost(Action action)
        {
            if (Handle == IntPtr.Zero) return false;
            _queue.Enqueue(action);
            return PostMessage(Handle, WM_APP_DISPATCH, IntPtr.Zero, IntPtr.Zero);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_APP_DISPATCH)
            {
                while (_queue.TryDequeue(out var action))
                {
                    if (_isBusy()) continue;
                    _setBusy(true);
                    try { action(); }
                    catch { }
                    finally { _setBusy(false); }
                }
                return;
            }
            base.WndProc(ref m);
        }
    }

    private readonly Func<bool>? _onDoubleClick;
    private readonly Func<Action?>? _tryPrepareDeferredAction;
    private readonly HashSet<string>? _allowedWindowClasses;
    private readonly HookProc _hookProc;
    private DispatchWindow? _dispatchWindow;
    private IntPtr _hookHandle;
    private bool _isDispatching;

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool PostMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();

    /// <summary>
    /// Installs a thread-level double-click hook whose synchronous callback MUST defer any UI/modal work itself
    /// (e.g., via ExcelAsyncUtil.QueueAsMacro). Returns true to suppress the host's default handling.
    /// </summary>
    public ThreadMouseDoubleClickHook(Func<bool> onDoubleClick, params string[] allowedWindowClasses)
    {
        _onDoubleClick = onDoubleClick ?? throw new ArgumentNullException(nameof(onDoubleClick));
        _allowedWindowClasses = BuildClassSet(allowedWindowClasses);
        _hookProc = Callback;
        InstallHook();
    }

    /// <summary>
    /// Installs a thread-level double-click hook that inspects selection synchronously inside the hook,
    /// then posts any non-null modal/editor action onto a hidden message-only window outside the hook callback.
    /// </summary>
    public ThreadMouseDoubleClickHook(Func<Action?> tryPrepareDeferredAction, params string[] allowedWindowClasses)
    {
        _tryPrepareDeferredAction = tryPrepareDeferredAction ?? throw new ArgumentNullException(nameof(tryPrepareDeferredAction));
        _allowedWindowClasses = BuildClassSet(allowedWindowClasses);
        _hookProc = Callback;
        try
        {
            _dispatchWindow = new DispatchWindow(() => _isDispatching, b => _isDispatching = b);
        }
        catch { }
        InstallHook();
    }

    private static HashSet<string>? BuildClassSet(string[]? classes)
    {
        if (classes == null || classes.Length == 0) return null;
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var c in classes)
        {
            if (!string.IsNullOrWhiteSpace(c))
                set.Add(c.Trim());
        }
        return set.Count > 0 ? set : null;
    }

    private void InstallHook()
    {
        try
        {
            _hookHandle = SetWindowsHookEx(WH_MOUSE, _hookProc, IntPtr.Zero, GetCurrentThreadId());
        }
        catch { }
    }

    internal static bool ShouldHandleDoubleClick(
        int nCode,
        IntPtr wParam,
        uint hitTestCode,
        string? windowClassName,
        IReadOnlyCollection<string>? allowedWindowClasses)
    {
        if (nCode != HC_ACTION || wParam != (IntPtr)WM_LBUTTONDBLCLK || hitTestCode != HTCLIENT)
            return false;

        if (allowedWindowClasses == null || allowedWindowClasses.Count == 0)
            return true;

        if (string.IsNullOrWhiteSpace(windowClassName))
            return false;

        foreach (var allowed in allowedWindowClasses)
        {
            if (string.Equals(allowed, windowClassName, StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

    private IntPtr Callback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode < 0)
            return CallNextHookEx(_hookHandle, nCode, wParam, lParam);

        if (!_isDispatching && nCode == HC_ACTION && wParam == (IntPtr)WM_LBUTTONDBLCLK && lParam != IntPtr.Zero)
        {
            try
            {
                var info = Marshal.PtrToStructure<MOUSEHOOKSTRUCT>(lParam);
                string? className = null;
                if (info.hwnd != IntPtr.Zero)
                {
                    var sb = new StringBuilder(64);
                    if (GetClassName(info.hwnd, sb, sb.Capacity) > 0)
                        className = sb.ToString();
                }

                if (ShouldHandleDoubleClick(nCode, wParam, info.wHitTestCode, className, _allowedWindowClasses))
                {
                    if (_tryPrepareDeferredAction != null && _dispatchWindow != null)
                    {
                        var deferred = _tryPrepareDeferredAction();
                        if (deferred != null && _dispatchWindow.TryPost(deferred))
                            return (IntPtr)1;
                    }
                    else if (_onDoubleClick != null)
                    {
                        _isDispatching = true;
                        try
                        {
                            if (_onDoubleClick())
                                return (IntPtr)1;
                        }
                        finally
                        {
                            _isDispatching = false;
                        }
                    }
                }
            }
            catch { }
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
        if (_dispatchWindow != null)
        {
            try { _dispatchWindow.DestroyHandle(); } catch { }
            _dispatchWindow = null;
        }
    }
}
