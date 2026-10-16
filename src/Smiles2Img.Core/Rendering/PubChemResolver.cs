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

    private static readonly Dictionary<string, (string CompoundName, string IupacName)> BuiltInSmilesToCompound = new(StringComparer.OrdinalIgnoreCase)
    {
        ["CCO"] = ("Ethanol", "ethanol"),
        ["OCC"] = ("Ethanol", "ethanol"),
        ["CO"] = ("Methanol", "methanol"),
        ["OC"] = ("Methanol", "methanol"),
        ["O"] = ("Water", "oxidane"),
        ["CC(C)O"] = ("Isopropanol", "propan-2-ol"),
        ["CC(=O)C"] = ("Acetone", "propan-2-one"),
        ["CC(=O)O"] = ("Acetic Acid", "acetic acid"),
        ["C(=O)O"] = ("Formic Acid", "formic acid"),
        ["c1ccccc1"] = ("Benzene", "benzene"),
        ["Cc1ccccc1"] = ("Toluene", "toluene"),
        ["Oc1ccccc1"] = ("Phenol", "phenol"),
        ["Nc1ccccc1"] = ("Aniline", "aniline"),
        ["OC(=O)c1ccccc1"] = ("Benzoic Acid", "benzoic acid"),
        ["Oc1ccccc1C(=O)O"] = ("Salicylic Acid", "2-hydroxybenzoic acid"),
        ["CC(=O)Oc1ccccc1C(=O)O"] = ("Aspirin", "2-acetyloxybenzoic acid"),
        ["Cn1cnc2c1c(=O)n(C)c(=O)n2C"] = ("Caffeine", "1,3,7-trimethylpurine-2,6-dione"),
        ["CC(=O)Nc1ccc(O)cc1"] = ("Acetaminophen", "N-(4-hydroxyphenyl)acetamide"),
        ["CC(C)Cc1ccc(C(C)C(=O)O)cc1"] = ("Ibuprofen", "2-[4-(2-methylpropyl)phenyl]propanoic acid"),
        ["COc1ccc2cc([C@H](C)C(=O)O)ccc2c1"] = ("Naproxen", "(2S)-2-(6-methoxynaphthalen-2-yl)propanoic acid"),
        ["c1ccncc1"] = ("Pyridine", "pyridine"),
        ["C1CCOC1"] = ("Tetrahydrofuran", "oxolane"),
        ["CS(=O)C"] = ("Dimethyl Sulfoxide", "methylsulfinylmethane"),
        ["CN(C)C=O"] = ("Dimethylformamide", "N,N-dimethylformamide"),
        ["CC#N"] = ("Acetonitrile", "acetonitrile"),
        ["ClCCl"] = ("Dichloromethane", "dichloromethane"),
        ["ClC(Cl)Cl"] = ("Chloroform", "trichloromethane"),
        ["CCOC(=O)C"] = ("Ethyl Acetate", "ethyl acetate"),
        ["CCOCC"] = ("Diethyl Ether", "ethoxyethane"),
        ["CCCCCC"] = ("Hexane", "hexane"),
        ["C1CCCCC1"] = ("Cyclohexane", "cyclohexane"),
        ["NCC(=O)O"] = ("Glycine", "2-aminoacetic acid"),
        ["C[C@H](N)C(=O)O"] = ("Alanine", "(2S)-2-aminopropanoic acid"),
        ["C[C@H](O)C(=O)O"] = ("L-Lactic Acid", "(2S)-2-hydroxypropanoic acid"),
        ["OC(=O)CC(O)(C(=O)O)CC(=O)O"] = ("Citric Acid", "2-hydroxypropane-1,2,3-tricarboxylic acid"),
        ["OC[C@H](O)[C@H]1OC(=O)C(O)=C1O"] = ("Ascorbic Acid", "(5R)-[(1S)-1,2-dihydroxyethyl]-3,4-dihydroxyfuran-2(5H)-one"),
        ["OC[C@H]1OC(O)[C@H](O)[C@@H](O)[C@@H]1O"] = ("D-Glucose", "(3R,4S,5S,6R)-6-(hydroxymethyl)oxane-2,3,4,5-tetrol"),
        ["NCCc1ccc(O)c(O)c1"] = ("Dopamine", "4-(2-aminoethyl)benzene-1,2-diol"),
        ["NCCc1c[nH]c2ccc(O)cc12"] = ("Serotonin", "3-(2-aminoethyl)-1H-indol-5-ol"),
        ["CNC[C@H](O)c1ccc(O)c(O)c1"] = ("Epinephrine", "4-[(1R)-1-hydroxy-2-(methylamino)ethyl]benzene-1,2-diol")
    };

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

        if (BuiltInSmilesToCompound.TryGetValue(s, out var builtIn))
            return builtIn;

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

    private static readonly ConcurrentDictionary<string, string> NameToSmilesCache = new(StringComparer.OrdinalIgnoreCase);
    private static bool _nameCacheLoaded;

    private static readonly Dictionary<string, string> BuiltInCommonChemicals = new(StringComparer.OrdinalIgnoreCase)
    {
        // Common Drugs & Bioactives
        ["aspirin"] = "CC(=O)Oc1ccccc1C(=O)O",
        ["acetylsalicylic acid"] = "CC(=O)Oc1ccccc1C(=O)O",
        ["2-acetoxybenzoic acid"] = "CC(=O)Oc1ccccc1C(=O)O",
        ["50-78-2"] = "CC(=O)Oc1ccccc1C(=O)O",
        ["caffeine"] = "Cn1cnc2c1c(=O)n(C)c(=O)n2C",
        ["58-08-2"] = "Cn1cnc2c1c(=O)n(C)c(=O)n2C",
        ["paracetamol"] = "CC(=O)Nc1ccc(O)cc1",
        ["acetaminophen"] = "CC(=O)Nc1ccc(O)cc1",
        ["tylenol"] = "CC(=O)Nc1ccc(O)cc1",
        ["103-90-2"] = "CC(=O)Nc1ccc(O)cc1",
        ["ibuprofen"] = "CC(C)Cc1ccc(C(C)C(=O)O)cc1",
        ["15687-27-1"] = "CC(C)Cc1ccc(C(C)C(=O)O)cc1",
        ["naproxen"] = "COc1ccc2cc([C@H](C)C(=O)O)ccc2c1",
        ["22204-53-1"] = "COc1ccc2cc([C@H](C)C(=O)O)ccc2c1",
        ["penicillin g"] = "CC1([C@@H](N2[C@H](S1)[C@@H](C2=O)NC(=O)Cc3ccccc3)C(=O)O)C",
        ["benzylpenicillin"] = "CC1([C@@H](N2[C@H](S1)[C@@H](C2=O)NC(=O)Cc3ccccc3)C(=O)O)C",
        ["61-33-6"] = "CC1([C@@H](N2[C@H](S1)[C@@H](C2=O)NC(=O)Cc3ccccc3)C(=O)O)C",
        ["amoxicillin"] = "CC1([C@@H](N2[C@H](S1)[C@@H](C2=O)NC(=O)[C@@H](c3ccc(O)cc3)N)C(=O)O)C",
        ["26787-78-0"] = "CC1([C@@H](N2[C@H](S1)[C@@H](C2=O)NC(=O)[C@@H](c3ccc(O)cc3)N)C(=O)O)C",
        ["artemisinin"] = "C[C@@H]1CC[C@H]2[C@@H](C)C(=O)O[C@@H]3O[C@@]4(C)CC[C@@H]1[C@]23OO4",
        ["63968-64-9"] = "C[C@@H]1CC[C@H]2[C@@H](C)C(=O)O[C@@H]3O[C@@]4(C)CC[C@@H]1[C@]23OO4",
        ["paclitaxel"] = "CC1=C2[C@@]([C@]([C@H]([C@@H]3[C@]4([C@H](OC4)C[C@@H]([C@]3(C(=O)[C@@H]2OC(=O)C)C)O)OC(=O)C)OC(=O)c5ccccc5)(C[C@@H]1OC(=O)[C@H](O)[C@@H](NC(=O)c6ccccc6)c7ccccc7)O)(C)C",
        ["taxol"] = "CC1=C2[C@@]([C@]([C@H]([C@@H]3[C@]4([C@H](OC4)C[C@@H]([C@]3(C(=O)[C@@H]2OC(=O)C)C)O)OC(=O)C)OC(=O)c5ccccc5)(C[C@@H]1OC(=O)[C@H](O)[C@@H](NC(=O)c6ccccc6)c7ccccc7)O)(C)C",
        ["33069-62-4"] = "CC1=C2[C@@]([C@]([C@H]([C@@H]3[C@]4([C@H](OC4)C[C@@H]([C@]3(C(=O)[C@@H]2OC(=O)C)C)O)OC(=O)C)OC(=O)c5ccccc5)(C[C@@H]1OC(=O)[C@H](O)[C@@H](NC(=O)c6ccccc6)c7ccccc7)O)(C)C",
        ["dopamine"] = "NCCc1ccc(O)c(O)c1",
        ["51-61-6"] = "NCCc1ccc(O)c(O)c1",
        ["serotonin"] = "NCCc1c[nH]c2ccc(O)cc12",
        ["50-67-9"] = "NCCc1c[nH]c2ccc(O)cc12",
        ["adrenaline"] = "CNC[C@H](O)c1ccc(O)c(O)c1",
        ["epinephrine"] = "CNC[C@H](O)c1ccc(O)c(O)c1",
        ["51-43-4"] = "CNC[C@H](O)c1ccc(O)c(O)c1",
        ["cholesterol"] = "C[C@H](CCCC(C)C)[C@H]1CC[C@@H]2[C@@]1(CC[C@H]3[C@H]2CC=C4[C@@]3(CC[C@@H](C4)O)C)C",
        ["57-88-5"] = "C[C@H](CCCC(C)C)[C@H]1CC[C@@H]2[C@@]1(CC[C@H]3[C@H]2CC=C4[C@@]3(CC[C@@H](C4)O)C)C",

        // Common Solvents & Laboratory Reagents
        ["water"] = "O",
        ["7732-18-5"] = "O",
        ["ethanol"] = "CCO",
        ["ethyl alcohol"] = "CCO",
        ["64-17-5"] = "CCO",
        ["methanol"] = "CO",
        ["methyl alcohol"] = "CO",
        ["67-56-1"] = "CO",
        ["isopropanol"] = "CC(C)O",
        ["2-propanol"] = "CC(C)O",
        ["67-63-0"] = "CC(C)O",
        ["acetone"] = "CC(=O)C",
        ["propanone"] = "CC(=O)C",
        ["67-64-1"] = "CC(=O)C",
        ["acetic acid"] = "CC(=O)O",
        ["ethanoic acid"] = "CC(=O)O",
        ["64-19-7"] = "CC(=O)O",
        ["formic acid"] = "C(=O)O",
        ["64-18-6"] = "C(=O)O",
        ["benzene"] = "c1ccccc1",
        ["71-43-2"] = "c1ccccc1",
        ["toluene"] = "Cc1ccccc1",
        ["108-88-3"] = "Cc1ccccc1",
        ["phenol"] = "Oc1ccccc1",
        ["108-95-2"] = "Oc1ccccc1",
        ["aniline"] = "Nc1ccccc1",
        ["62-53-3"] = "Nc1ccccc1",
        ["benzoic acid"] = "OC(=O)c1ccccc1",
        ["65-85-0"] = "OC(=O)c1ccccc1",
        ["salicylic acid"] = "Oc1ccccc1C(=O)O",
        ["69-72-7"] = "Oc1ccccc1C(=O)O",
        ["pyridine"] = "c1ccncc1",
        ["110-86-1"] = "c1ccncc1",
        ["thf"] = "C1CCOC1",
        ["tetrahydrofuran"] = "C1CCOC1",
        ["109-99-9"] = "C1CCOC1",
        ["dmso"] = "CS(=O)C",
        ["dimethyl sulfoxide"] = "CS(=O)C",
        ["67-68-5"] = "CS(=O)C",
        ["dmf"] = "CN(C)C=O",
        ["dimethylformamide"] = "CN(C)C=O",
        ["68-12-2"] = "CN(C)C=O",
        ["acetonitrile"] = "CC#N",
        ["75-05-8"] = "CC#N",
        ["dichloromethane"] = "ClCCl",
        ["dcm"] = "ClCCl",
        ["75-09-2"] = "ClCCl",
        ["chloroform"] = "ClC(Cl)Cl",
        ["67-66-3"] = "ClC(Cl)Cl",
        ["ethyl acetate"] = "CCOC(=O)C",
        ["141-78-6"] = "CCOC(=O)C",
        ["diethyl ether"] = "CCOCC",
        ["60-29-7"] = "CCOCC",
        ["hexane"] = "CCCCCC",
        ["110-54-3"] = "CCCCCC",
        ["cyclohexane"] = "C1CCCCC1",
        ["110-82-7"] = "C1CCCCC1",

        // Carbohydrates, Amino Acids & Metabolites
        ["glucose"] = "OC[C@H]1OC(O)[C@H](O)[C@@H](O)[C@@H]1O",
        ["d-glucose"] = "OC[C@H]1OC(O)[C@H](O)[C@@H](O)[C@@H]1O",
        ["50-99-7"] = "OC[C@H]1OC(O)[C@H](O)[C@@H](O)[C@@H]1O",
        ["lactic acid"] = "C[C@H](O)C(=O)O",
        ["50-21-5"] = "C[C@H](O)C(=O)O",
        ["citric acid"] = "OC(=O)CC(O)(C(=O)O)CC(=O)O",
        ["77-92-9"] = "OC(=O)CC(O)(C(=O)O)CC(=O)O",
        ["ascorbic acid"] = "OC[C@H](O)[C@H]1OC(=O)C(O)=C1O",
        ["vitamin c"] = "OC[C@H](O)[C@H]1OC(=O)C(O)=C1O",
        ["50-81-7"] = "OC[C@H](O)[C@H]1OC(=O)C(O)=C1O",
        ["glycine"] = "NCC(=O)O",
        ["56-40-6"] = "NCC(=O)O",
        ["alanine"] = "C[C@H](N)C(=O)O",
        ["56-41-7"] = "C[C@H](N)C(=O)O"
    };

    private static string GetNameCacheFilePath()
    {
        var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Smiles2Img");
        try { Directory.CreateDirectory(dir); } catch { }
        return Path.Combine(dir, "pubchem_name2smiles_cache.json");
    }

    private static void EnsureNameCacheLoaded()
    {
        if (_nameCacheLoaded) return;
        lock (FileLock)
        {
            if (_nameCacheLoaded) return;
            _nameCacheLoaded = true;
            try
            {
                var path = GetNameCacheFilePath();
                if (File.Exists(path))
                {
                    var json = File.ReadAllText(path, Encoding.UTF8);
                    var dict = Serializer.Deserialize<Dictionary<string, string>>(json);
                    if (dict != null)
                    {
                        foreach (var kvp in dict)
                        {
                            if (!string.IsNullOrWhiteSpace(kvp.Value))
                                NameToSmilesCache[kvp.Key] = kvp.Value;
                        }
                    }
                }
            }
            catch { }
        }
    }

    private static void SaveNameCache()
    {
        try
        {
            lock (FileLock)
            {
                var dict = new Dictionary<string, string>();
                foreach (var kvp in NameToSmilesCache)
                {
                    dict[kvp.Key] = kvp.Value;
                }
                var json = Serializer.Serialize(dict);
                File.WriteAllText(GetNameCacheFilePath(), json, Encoding.UTF8);
            }
        }
        catch { }
    }

    /// <summary>
    /// Resolves a chemical name, IUPAC name, or CAS registry number to a canonical SMILES string.
    /// Checks 0ms built-in dictionary first, then persistent local cache, then PubChem REST API.
    /// </summary>
    public static string? ResolveNameToSmiles(string? nameOrCas, int timeoutMs = 3000)
    {
        if (string.IsNullOrWhiteSpace(nameOrCas)) return null;
        var q = nameOrCas!.Trim();

        if (BuiltInCommonChemicals.TryGetValue(q, out var builtInSmiles))
            return builtInSmiles;

        EnsureNameCacheLoaded();
        if (NameToSmilesCache.TryGetValue(q, out var cachedSmiles))
            return cachedSmiles;

        try
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            var url = $"https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/{Uri.EscapeDataString(q)}/property/IsomericSMILES,CanonicalSMILES,ConnectivitySMILES/JSON";
            var request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "GET";
            request.Timeout = timeoutMs;
            request.ReadWriteTimeout = timeoutMs;

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
                        string? smiles = null;
                        string[] keys = ["IsomericSMILES", "CanonicalSMILES", "ConnectivitySMILES", "SMILES"];
                        foreach (var k in keys)
                        {
                            if (firstProp.TryGetValue(k, out var val) && val != null && !string.IsNullOrWhiteSpace(val.ToString()))
                            {
                                smiles = val.ToString()!.Trim();
                                break;
                            }
                        }

                        if (!string.IsNullOrEmpty(smiles))
                        {
                            NameToSmilesCache[q] = smiles!;
                            SaveNameCache();
                            return smiles;
                        }
                    }
                }
            }
        }
        catch { }

        return null;
    }
}
