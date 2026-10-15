using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Web.Script.Serialization;
using System.Xml.Linq;

namespace Smiles2Img;

public sealed class MolstarMetadata
{
    public string Id { get; set; } = "";
    public string Pdb { get; set; } = "";
    public string Mode { get; set; } = "image"; // "image" or "3dmodel"
    public bool Embed { get; set; } = true;
    public string Background { get; set; } = "transparent";
}

public static class MolstarSessionManager
{
    public const string AltTextPrefix = "MOLSTAR:";
    public const string CustomXmlNamespace = "urn:smiles2img:molstar-session";

    private static readonly JavaScriptSerializer Serializer = new() { MaxJsonLength = int.MaxValue };

    private static string GetCacheDirectory()
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Smiles2Img", "MolstarSessions");
        Directory.CreateDirectory(dir);
        return dir;
    }

    public static string FormatAltText(MolstarMetadata meta)
    {
        return AltTextPrefix + Serializer.Serialize(new Dictionary<string, object>
        {
            ["id"] = meta.Id,
            ["pdb"] = meta.Pdb ?? "",
            ["mode"] = meta.Mode ?? "image",
            ["embed"] = meta.Embed,
            ["bg"] = meta.Background ?? "transparent"
        });
    }

    public static MolstarMetadata? ParseAltText(params string?[] texts)
    {
        foreach (var text in texts)
        {
            if (string.IsNullOrWhiteSpace(text)) continue;
            var trimmed = text!.Trim();
            int idx = trimmed.IndexOf(AltTextPrefix, StringComparison.OrdinalIgnoreCase);
            if (idx < 0) continue;

            var jsonPart = trimmed.Substring(idx + AltTextPrefix.Length).Trim();
            try
            {
                var dict = Serializer.Deserialize<Dictionary<string, object>>(jsonPart);
                if (dict != null)
                {
                    return new MolstarMetadata
                    {
                        Id = dict.TryGetValue("id", out var id) && id is string sId ? sId : "",
                        Pdb = dict.TryGetValue("pdb", out var pdb) && pdb is string sPdb ? sPdb : "",
                        Mode = dict.TryGetValue("mode", out var mode) && mode is string sMode ? sMode : "image",
                        Embed = dict.TryGetValue("embed", out var embed) && embed is bool bEmbed ? bEmbed : true,
                        Background = dict.TryGetValue("bg", out var bg) && bg is string sBg ? sBg : "transparent"
                    };
                }
            }
            catch { }
        }
        return null;
    }

    public static byte[] CompressGzip(string text)
    {
        if (string.IsNullOrEmpty(text)) return Array.Empty<byte>();
        var bytes = Encoding.UTF8.GetBytes(text);
        using var ms = new MemoryStream();
        using (var gz = new GZipStream(ms, CompressionLevel.Optimal, true))
        {
            gz.Write(bytes, 0, bytes.Length);
        }
        return ms.ToArray();
    }

    public static string DecompressGzip(byte[] gzipBytes)
    {
        if (gzipBytes == null || gzipBytes.Length == 0) return "";
        using var ms = new MemoryStream(gzipBytes);
        using var gz = new GZipStream(ms, CompressionMode.Decompress);
        using var outMs = new MemoryStream();
        gz.CopyTo(outMs);
        return Encoding.UTF8.GetString(outMs.ToArray());
    }

    public static void SaveLocalCache(string sessionId, string snapshotJson)
    {
        if (string.IsNullOrWhiteSpace(sessionId) || string.IsNullOrWhiteSpace(snapshotJson)) return;
        try
        {
            var path = Path.Combine(GetCacheDirectory(), $"{sessionId}.molj.gz");
            var gz = CompressGzip(snapshotJson);
            File.WriteAllBytes(path, gz);
        }
        catch { }
    }

    public static string? LoadLocalCache(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId)) return null;
        try
        {
            var path = Path.Combine(GetCacheDirectory(), $"{sessionId}.molj.gz");
            if (File.Exists(path))
            {
                var gz = File.ReadAllBytes(path);
                return DecompressGzip(gz);
            }
        }
        catch { }
        return null;
    }

    /// <summary>
    /// Stores the gzip+base64 session payload into the Office Document's CustomXMLParts.
    /// Works seamlessly across PowerPoint, Excel, and Word.
    /// </summary>
    public static bool SaveToCustomXmlParts(object? document, string sessionId, string pdbId, string snapshotJson)
    {
        if (document == null || string.IsNullOrWhiteSpace(sessionId) || string.IsNullOrWhiteSpace(snapshotJson))
            return false;

        try
        {
            dynamic doc = document;
            dynamic? xmlParts = doc.CustomXMLParts;
            if (xmlParts == null) return false;

            var gz = CompressGzip(snapshotJson);
            var b64 = Convert.ToBase64String(gz);
            var xml = $"<molstar-session xmlns=\"{CustomXmlNamespace}\" id=\"{sessionId}\" pdb=\"{pdbId}\">{b64}</molstar-session>";

            // Remove any existing part with the same sessionId
            try
            {
                dynamic parts = xmlParts.SelectByNamespace(CustomXmlNamespace);
                if (parts != null)
                {
                    for (int i = parts.Count; i >= 1; i--)
                    {
                        dynamic p = parts.Item(i);
                        if (p != null)
                        {
                            string? pXml = Convert.ToString((object?)p.XML);
                            if (!string.IsNullOrEmpty(pXml) && pXml!.Contains($"id=\"{sessionId}\""))
                            {
                                p.Delete();
                            }
                        }
                    }
                }
            }
            catch { }

            xmlParts.Add(xml);
            return true;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine("Failed to save to CustomXMLParts: " + ex.Message);
            return false;
        }
    }

    /// <summary>
    /// Loads and decompresses the session snapshot from the Office Document's CustomXMLParts.
    /// </summary>
    public static string? LoadFromCustomXmlParts(object? document, string sessionId)
    {
        if (document == null || string.IsNullOrWhiteSpace(sessionId))
            return null;

        try
        {
            dynamic doc = document;
            dynamic? xmlParts = doc.CustomXMLParts;
            if (xmlParts == null) return null;

            dynamic parts = xmlParts.SelectByNamespace(CustomXmlNamespace);
            if (parts == null) return null;

            for (int i = 1; i <= parts.Count; i++)
            {
                dynamic p = parts.Item(i);
                if (p == null) continue;
                string? xml = Convert.ToString((object?)p.XML);
                if (!string.IsNullOrEmpty(xml) && xml!.Contains($"id=\"{sessionId}\""))
                {
                    var xdoc = XDocument.Parse(xml!);
                    var b64 = xdoc.Root?.Value;
                    if (!string.IsNullOrWhiteSpace(b64))
                    {
                        var gz = Convert.FromBase64String(b64!.Trim());
                        return DecompressGzip(gz);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine("Failed to load from CustomXMLParts: " + ex.Message);
        }
        return null;
    }
}
