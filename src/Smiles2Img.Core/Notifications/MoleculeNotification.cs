using System;
using Smiles2Img.Editor;

namespace Smiles2Img.Notifications;

/// <summary>
/// Unified notification manager for molecular copy/extraction actions across Excel, Word, and PowerPoint.
/// Sends a native Windows Toast Notification (Action Center / system banner) and displays
/// a non-blocking in-app floating toast near the mouse cursor without stealing focus.
/// </summary>
public static class MoleculeNotification
{
    /// <summary>
    /// Notify successful clipboard copy of a molecular property.
    /// </summary>
    /// <param name="label">Title of the property (e.g. "SMILES", "InChI", "Substance Name")</param>
    /// <param name="value">The copied string value</param>
    public static void NotifyCopied(string label, string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return;

        // 1. Fire native Windows Notification
        string displayVal = Truncate(value, 80);
        WindowsNotificationService.ShowToast("func_SMILES2IMG", $"✓ Copied {label} to clipboard: {displayVal}");

        // 2. Display lightweight non-blocking in-app toast near cursor
        CopiedToast.Show(label, value, isWarning: false);
    }

    /// <summary>
    /// Notify user when a structure or property cannot be resolved, without blocking modal popups.
    /// </summary>
    /// <param name="label">Context or property title</param>
    /// <param name="message">Warning / notice message</param>
    public static void NotifyWarning(string label, string message)
    {
        if (string.IsNullOrWhiteSpace(message)) return;

        // 1. Fire native Windows Notification
        WindowsNotificationService.ShowToast("func_SMILES2IMG", $"⚠ {label}: {message}");

        // 2. Display lightweight in-app toast near cursor
        CopiedToast.Show(label, message, isWarning: true);
    }

    private static string Truncate(string text, int maxLen)
    {
        if (string.IsNullOrWhiteSpace(text)) return "";
        text = text.Replace("\r", "").Replace("\n", " ");
        if (text.Length <= maxLen) return text;
        return text.Substring(0, maxLen - 3) + "...";
    }
}
