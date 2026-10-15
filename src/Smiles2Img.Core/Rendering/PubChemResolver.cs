using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Web.Script.Serialization;

namespace Smiles2Img.Rendering;

/// <summary>
/// Queries PubChem and RCSB PDB REST APIs to resolve compound names, IUPAC names,
/// and macromolecular titles, with in-memory and persistent disk caching.
/// </summary>
public static class PubChemResolver
{
    private static readonly ConcurrentDictionary<string, (string CompoundName, string IupacName)> SmilesCache = new(StringComparer.OrdinalIgnoreCase);
    private static readonly ConcurrentDictionary<string, (string Title, double? WeightKDa)> PdbCache = new(StringComparer.OrdinalIgnoreCase);
    private static readonly JavaScriptSerializer Serializer = new() { MaxJsonLength = int.MaxValue };
    private static readonly object FileLock = new();
    private static bool _cacheLoaded;

    private static string GetCacheFilePath()
    {
        var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Smiles2Img");
        try { Directory.CreateDirectory(dir); } catch { }
        return Path.Combine(dir, "pubchem_cache.json");
    }

    private static void EnsureCacheLoaded()
    {
        if (_cacheLoaded) return;
        lock (FileLock)
        {
            if (_cacheLoaded) return;
            _cacheLoaded = true;
            try
            {
                var path = GetCacheFilePath();
                if (File.Exists(path))
                {
                    var json = File.ReadAllText(path, Encoding.UTF8);
                    var dict = Serializer.Deserialize<Dictionary<string, string[]>>(json);
                    if (dict != null)
                    {
                        foreach (var kvp in dict)
                        {
                            if (kvp.Value != null && kvp.Value.Length >= 2)
                                SmilesCache[kvp.Key] = (kvp.Value[0] ?? "", kvp.Value[1] ?? "");
                        }
                    }
                }
            }
            catch { }
        }
    }

    private static void SaveCache()
    {
        try
        {
            lock (FileLock)
            {
                var dict = new Dictionary<string, string[]>();
                foreach (var kvp in SmilesCache)
                {
                    dict[kvp.Key] = new[] { kvp.Value.CompoundName, kvp.Value.IupacName };
                }
                var json = Serializer.Serialize(dict);
                File.WriteAllText(GetCacheFilePath(), json, Encoding.UTF8);
            }
        }
        catch { }
    }

    /// <summary>
    /// Resolves compound common title and formal IUPAC name from PubChem given a SMILES string.
    /// Falls back cleanly to empty strings if offline or compound not found.
    /// </summary>
    public static (string CompoundName, string IupacName) ResolveCompound(string? smiles, int timeoutMs = 2500)
    {
        if (string.IsNullOrWhiteSpace(smiles)) return ("", "");
        var s = smiles!.Trim();

        EnsureCacheLoaded();
        if (SmilesCache.TryGetValue(s, out var cached))
            return cached;

        try
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            var request = (HttpWebRequest)WebRequest.Create("https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/smiles/property/Title,IUPACName/JSON");
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            request.Timeout = timeoutMs;
            request.ReadWriteTimeout = timeoutMs;

            var postData = "smiles=" + Uri.EscapeDataString(s);
            var bytes = Encoding.UTF8.GetBytes(postData);
            request.ContentLength = bytes.Length;

            using (var stream = request.GetRequestStream())
            {
                stream.Write(bytes, 0, bytes.Length);
            }

            using var response = (HttpWebResponse)request.GetResponse();
            using var reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8);
            var json = reader.ReadToEnd();

            var data = Serializer.Deserialize<Dictionary<string, object>>(json);
            if (data != null && data.TryGetValue("PropertyTable", out var propTableObj) && propTableObj is Dictionary<string, object> propTable)
            {
                if (propTable.TryGetValue("Properties", out var propsObj) && propsObj is System.Collections.ArrayList props && props.Count > 0)
                {
                    if (props[0] is Dictionary<string, object> firstProp)
                    {
                        var title = firstProp.TryGetValue("Title", out var t) && t != null ? t.ToString()!.Trim() : "";
                        var iupac = firstProp.TryGetValue("IUPACName", out var u) && u != null ? u.ToString()!.Trim() : "";

                        var result = (title, iupac);
                        SmilesCache[s] = result;
                        SaveCache();
                        return result;
                    }
                }
            }
        }
        catch
        {
            // Network error, timeout, or compound not in PubChem
        }

        return ("", "");
    }

    /// <summary>
    /// Resolves macromolecule title and molecular weight in kDa from RCSB PDB given a PDB ID.
    /// </summary>
    public static (string Title, double? WeightKDa) ResolvePdb(string? pdbId, int timeoutMs = 2500)
    {
        if (string.IsNullOrWhiteSpace(pdbId)) return ("", null);
        var id = pdbId!.Trim().ToUpperInvariant();

        if (PdbCache.TryGetValue(id, out var cached))
            return cached;

        try
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            var url = $"https://data.rcsb.org/rest/v1/core/entry/{Uri.EscapeDataString(id)}";
            var request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "GET";
            request.Timeout = timeoutMs;
            request.ReadWriteTimeout = timeoutMs;

            using var response = (HttpWebResponse)request.GetResponse();
            using var reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8);
            var json = reader.ReadToEnd();

            var data = Serializer.Deserialize<Dictionary<string, object>>(json);
            if (data != null)
            {
                string title = "";
                double? mw = null;

                if (data.TryGetValue("struct", out var structObj) && structObj is Dictionary<string, object> structDict)
                {
                    if (structDict.TryGetValue("title", out var t) && t != null)
                        title = t.ToString()!.Trim();
                }

                if (data.TryGetValue("rcsb_entry_info", out var infoObj) && infoObj is Dictionary<string, object> infoDict)
                {
                    if (infoDict.TryGetValue("molecular_weight", out var w) && w != null && double.TryParse(w.ToString(), out var mwVal))
                        mw = mwVal;
                }

                var res = (title, mw);
                PdbCache[id] = res;
                return res;
            }
        }
        catch { }

        return ("", null);
    }
}
