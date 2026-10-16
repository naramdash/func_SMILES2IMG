using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using com.epam.indigo;

namespace Smiles2Img.Rendering;

public sealed class SdfRecord
{
    public string Name { get; set; } = "";
    public string Smiles { get; set; } = "";
    public string Formula { get; set; } = "";
    public double MolecularWeight { get; set; }
    public double ExactMass { get; set; }
    public Dictionary<string, string> Properties { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}

public static class SdfManager
{
    public static List<SdfRecord> ReadSdfFile(string filePath, int maxRecords = 2000)
    {
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            return [];

        var bytes = File.ReadAllBytes(filePath);
        return ReadSdfBytes(bytes, maxRecords);
    }

    public static List<SdfRecord> ReadSdfBytes(byte[] sdfBytes, int maxRecords = 2000)
    {
        var results = new List<SdfRecord>();
        if (sdfBytes == null || sdfBytes.Length == 0)
            return results;

        NativeIndigo.EnsureInitialized();
        using var indigo = new Indigo();
        using var reader = indigo.loadBuffer(sdfBytes);

        int count = 0;
        foreach (IndigoObject item in indigo.iterateSDF(reader))
        {
            if (count >= maxRecords)
            {
                item.Dispose();
                break;
            }

            using (item)
            {
                try
                {
                    if (item.countAtoms() <= 0)
                        continue;

                    var rec = new SdfRecord();
                    try { rec.Name = item.name()?.Trim() ?? ""; } catch { }
                    try { rec.Smiles = item.canonicalSmiles() ?? item.smiles() ?? ""; }
                    catch
                    {
                        try { rec.Smiles = item.smiles() ?? ""; } catch { }
                    }

                    if (string.IsNullOrWhiteSpace(rec.Smiles))
                        continue;

                    try { rec.Formula = item.grossFormula()?.Replace(" ", "") ?? ""; } catch { }
                    try { rec.MolecularWeight = Math.Round((double)item.molecularWeight(), 2); } catch { }
                    try { rec.ExactMass = Math.Round((double)item.monoisotopicMass(), 4); } catch { }

                    try
                    {
                        foreach (IndigoObject prop in item.iterateProperties())
                        {
                            using (prop)
                            {
                                var pName = prop.name()?.Trim() ?? "";
                                var pVal = prop.rawData()?.Trim() ?? "";
                                if (!string.IsNullOrEmpty(pName))
                                {
                                    rec.Properties[pName] = pVal;
                                    if (string.IsNullOrWhiteSpace(rec.Name) &&
                                        (pName.Equals("NAME", StringComparison.OrdinalIgnoreCase) ||
                                         pName.Equals("PUBCHEM_IUPAC_TRADITIONAL_NAME", StringComparison.OrdinalIgnoreCase) ||
                                         pName.Equals("TITLE", StringComparison.OrdinalIgnoreCase) ||
                                         pName.Equals("GENERIC_NAME", StringComparison.OrdinalIgnoreCase)))
                                    {
                                        rec.Name = pVal;
                                    }
                                }
                            }
                        }
                    }
                    catch { }

                    results.Add(rec);
                    count++;
                }
                catch { }
            }
        }

        return results;
    }

    public static void WriteSdfFile(string filePath, IEnumerable<SdfRecord> records)
    {
        var bytes = WriteSdfBytes(records);
        File.WriteAllBytes(filePath, bytes);
    }

    public static byte[] WriteSdfBytes(IEnumerable<SdfRecord> records)
    {
        NativeIndigo.EnsureInitialized();
        using var indigo = new Indigo();
        using var buf = indigo.writeBuffer();
        using var saver = indigo.createSaver(buf, "sdf");

        if (records != null)
        {
            foreach (var rec in records)
            {
                if (rec == null || string.IsNullOrWhiteSpace(rec.Smiles))
                    continue;

                try
                {
                    var smiles = MoleculeRenderer.TryConvertToSmiles(rec.Smiles) ?? rec.Smiles.Trim();
                    using var mol = indigo.loadMolecule(smiles);
                    if (mol.countAtoms() <= 0)
                        continue;

                    try { mol.layout(); } catch { }

                    if (!string.IsNullOrWhiteSpace(rec.Name))
                    {
                        try { mol.setName(rec.Name.Trim()); } catch { }
                    }

                    var canSmiles = "";
                    try { canSmiles = mol.canonicalSmiles() ?? smiles; } catch { canSmiles = smiles; }
                    var formula = !string.IsNullOrWhiteSpace(rec.Formula)
                        ? rec.Formula
                        : (mol.grossFormula()?.Replace(" ", "") ?? "");
                    var mw = rec.MolecularWeight > 0
                        ? rec.MolecularWeight
                        : Math.Round((double)mol.molecularWeight(), 2);

                    if (!string.IsNullOrWhiteSpace(rec.Name) && !rec.Properties.ContainsKey("Name"))
                        mol.setProperty("Name", rec.Name.Trim());

                    mol.setProperty("SMILES", canSmiles);
                    if (!string.IsNullOrEmpty(formula))
                        mol.setProperty("Formula", formula);
                    if (mw > 0)
                        mol.setProperty("MW", mw.ToString("F2", CultureInfo.InvariantCulture));

                    foreach (var kv in rec.Properties)
                    {
                        if (!string.IsNullOrWhiteSpace(kv.Key) && kv.Value != null)
                        {
                            try { mol.setProperty(kv.Key.Trim(), kv.Value); } catch { }
                        }
                    }

                    saver.append(mol);
                }
                catch { }
            }
        }

        saver.close();
        return buf.toBuffer();
    }
}
