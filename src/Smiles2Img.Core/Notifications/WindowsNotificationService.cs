using System;
using System.Reflection;
using Microsoft.Win32;

namespace Smiles2Img.Notifications;

/// <summary>
/// Provides native Windows 10/11 Toast Notifications via WinRT (Windows.UI.Notifications)
/// with graceful fallback on older Windows versions or when notifications are disabled.
/// </summary>
public static class WindowsNotificationService
{
    private const string Aumid = "func_SMILES2IMG";
    private static bool _initialized;
    private static object? _notifier;
    private static MethodInfo? _showMethod;
    private static Type? _docType;
    private static Type? _toastType;
    private static MethodInfo? _loadXmlMethod;

    private static void EnsureInitialized()
    {
        if (_initialized) return;
        _initialized = true;

        try
        {
            // Register AUMID in HKCU if not already present
            using (var key = Registry.CurrentUser.CreateSubKey(@"Software\Classes\AppUserModelId\" + Aumid))
            {
                if (key != null && key.GetValue("DisplayName") == null)
                {
                    key.SetValue("DisplayName", "func_SMILES2IMG", RegistryValueKind.String);
                }
            }

            _docType = Type.GetType("Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType=WindowsRuntime");
            var toastMgrType = Type.GetType("Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType=WindowsRuntime");
            _toastType = Type.GetType("Windows.UI.Notifications.ToastNotification, Windows.UI.Notifications, ContentType=WindowsRuntime");

            if (_docType != null && toastMgrType != null && _toastType != null)
            {
                _loadXmlMethod = _docType.GetMethod("LoadXml", new[] { typeof(string) });
                var createNotifier = toastMgrType.GetMethod("CreateToastNotifier", new[] { typeof(string) });
                if (createNotifier != null)
                {
                    _notifier = createNotifier.Invoke(null, new object[] { Aumid });
                    if (_notifier != null)
                    {
                        _showMethod = _notifier.GetType().GetMethod("Show");
                    }
                }
            }
        }
        catch
        {
            // Gracefully ignore WinRT init errors
        }
    }

    /// <summary>
    /// Displays a native Windows Toast Notification.
    /// </summary>
    /// <param name="title">Heading / title of the toast.</param>
    /// <param name="message">Main body message.</param>
    /// <returns>True if the notification was successfully dispatched to Windows, false otherwise.</returns>
    public static bool ShowToast(string title, string message)
    {
        try
        {
            EnsureInitialized();
            if (_notifier != null && _showMethod != null && _docType != null && _toastType != null && _loadXmlMethod != null)
            {
                var doc = Activator.CreateInstance(_docType);
                string safeTitle = EscapeXml(title ?? "func_SMILES2IMG");
                string safeMessage = EscapeXml(message ?? "");
                string xml = $"<toast duration='short'><visual><binding template='ToastGeneric'><text>{safeTitle}</text><text>{safeMessage}</text></binding></visual></toast>";
                _loadXmlMethod.Invoke(doc, new object[] { xml });
                var toast = Activator.CreateInstance(_toastType, new object[] { doc });
                _showMethod.Invoke(_notifier, new object[] { toast });
                return true;
            }
        }
        catch
        {
            // Fallback gracefully on notification delivery failure
        }
        return false;
    }

    private static string EscapeXml(string text)
    {
        if (string.IsNullOrEmpty(text)) return "";
        return text
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;")
            .Replace("'", "&apos;");
    }
}
