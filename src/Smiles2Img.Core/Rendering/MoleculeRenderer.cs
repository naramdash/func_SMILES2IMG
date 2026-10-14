using System;
using System.Collections.Generic;
using System.Linq;
using com.epam.indigo;

namespace Smiles2Img.Rendering;

[Serializable]
public sealed class InvalidSmilesException(string message) : Exception(message);

public sealed class MoleculeInfo
{
    public string CompoundName { get; set; } = "";
    public string IupacName { get; set; } = "";
    public string Formula { get; set; } = "";
    public double MolecularWeight { get; set; }
    public double ExactMass { get; set; }
    public string Smiles { get; set; } = "";
    public bool IsReaction { get; set; }

    public string InChI { get; set; } = "";
    public string InChIKey { get; set; } = "";

    public string FormatClipboardSummary()
    {
        if (IsReaction)
            return $"Reaction SMILES: {Smiles}";

        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(CompoundName))
            parts.Add($"Name: {CompoundName}");
        if (!string.IsNullOrWhiteSpace(IupacName) && !string.Equals(IupacName, CompoundName, StringComparison.OrdinalIgnoreCase))
            parts.Add($"IUPAC: {IupacName}");
        parts.Add($"Formula: {Formula}");
        parts.Add($"MW: {MolecularWeight:F2} g/mol");
        if (ExactMass > 0)
            parts.Add($"Exact Mass: {ExactMass:F4}");
        parts.Add($"SMILES: {Smiles}");
        if (!string.IsNullOrWhiteSpace(InChI))
            parts.Add($"InChI: {InChI}");
        if (!string.IsNullOrWhiteSpace(InChIKey))
            parts.Add($"InChIKey: {InChIKey}");

        return string.Join(" | ", parts);
    }

    public string FormatDisplayMessage()
    {
        if (IsReaction)
            return $"[Reaction]\n\n• Reaction SMILES: {Smiles}";

        var sb = new System.Text.StringBuilder();
        if (!string.IsNullOrWhiteSpace(CompoundName))
            sb.AppendLine($"• Substance Name: {CompoundName}");
        if (!string.IsNullOrWhiteSpace(IupacName) && !string.Equals(IupacName, CompoundName, StringComparison.OrdinalIgnoreCase))
            sb.AppendLine($"• IUPAC Name: {IupacName}");
        sb.AppendLine($"• Formula: {Formula}");
        sb.AppendLine($"• Molecular Weight (MW): {MolecularWeight:F2} g/mol");
        if (ExactMass > 0)
            sb.AppendLine($"• Exact Mass: {ExactMass:F4}");
        sb.AppendLine($"• Canonical SMILES: {Smiles}");
        if (!string.IsNullOrWhiteSpace(InChI))
            sb.AppendLine($"• InChI: {InChI}");
        if (!string.IsNullOrWhiteSpace(InChIKey))
            sb.Append($"• InChIKey: {InChIKey}");
        return sb.ToString().TrimEnd();
    }
}

public static class MoleculeRenderer
{
    private const int RingTopology = 10;
    private static readonly char[] TransformSeparators = [' ', ',', '|', ';'];

    public static string Normalize(string smiles)
    {
        var input = smiles?.Trim();
        if (string.IsNullOrEmpty(input) || input!.Length > 2000)
            throw new InvalidSmilesException("Enter a SMILES string containing 1 to 2000 characters.");

        // Fast syntax validation (bracket and parenthesis balancing)
        var branches = 0;
        var inAtom = false;
        foreach (var character in input)
        {
            if (character == '[' && !inAtom) inAtom = true;
            else if (character == ']' && inAtom) inAtom = false;
            else if (character is '[' or ']') throw Invalid();
            else if (!inAtom && character == '(') branches++;
            else if (!inAtom && character == ')' && --branches < 0) throw Invalid();
        }

        if (inAtom || branches != 0) throw Invalid();
        return input;
    }

    public static byte[] Render(string smiles, bool color = true, string? transform = null, string? background = null)
        => Render(smiles, new RenderOptions { Color = color, Transform = transform, Background = background });

    public static byte[] Render(string smiles, RenderOptions options)
    {
        options ??= new RenderOptions();
        using var indigo = new Indigo();
        using var renderer = new IndigoRenderer(indigo);

        indigo.setOption("render-output-format", "png");
        indigo.setOption("render-bond-length", 80.0);
        indigo.setOption("render-margins", 30, 30);
        indigo.setOption("render-relative-thickness", 1.5);
        indigo.setOption("aromaticity-model", "generic");
        indigo.setOption("smart-layout", "true");
        indigo.setOption("layout-orientation", "horizontal");
        indigo.setOption("render-coloring", options.Color);

        ApplyBackground(indigo, options.Background);

        if (!string.IsNullOrWhiteSpace(options.BaseColor))
            ApplyBaseColor(indigo, options.BaseColor);

        if (options.AtomNumbers)
        {
            indigo.setOption("render-atom-ids-visible", "true");
            indigo.setOption("render-atom-bond-ids-from-one", "true");
        }

        indigo.setOption("render-stereo-style", options.StereoLabels ? "ext" : "none");

        using var molecule = Parse(indigo, Normalize(smiles));

        if (options.UnfoldHydrogens)
            molecule.unfoldHydrogens();
        else
            ExposeRingJunctionChiralHydrogens(molecule);

        if (options.AllCarbons)
            indigo.setOption("render-label-mode", "all");
        else if (molecule.countAtoms() <= 4)
            indigo.setOption("render-label-mode", "terminal-hetero");
        else
            indigo.setOption("render-label-mode", "hetero");

        molecule.dearomatize();
        molecule.layout();
        ApplyTransform(molecule, options.Transform);

        if (!options.AllCarbons)
            LabelHeteroMethyls(molecule);

        return renderer.renderToBuffer(molecule);
    }

    public static bool IsKet(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return false;
        var trimmed = text!.Trim();
        return trimmed.StartsWith("KET:", StringComparison.OrdinalIgnoreCase) ||
               (trimmed.StartsWith("{") && (trimmed.Contains("\"root\"") || trimmed.Contains("\"nodes\"")));
    }

    public static string CleanKet(string? ket)
    {
        if (string.IsNullOrWhiteSpace(ket)) return "";
        var trimmed = ket!.Trim();
        if (trimmed.StartsWith("KET:", StringComparison.OrdinalIgnoreCase))
            return trimmed.Substring(4).Trim();
        return trimmed;
    }

    public static string? TryConvertToSmiles(string? input)
    {
        if (string.IsNullOrWhiteSpace(input)) return null;
        var trimmed = input!.Trim();

        if (trimmed.StartsWith("SMILES:", StringComparison.OrdinalIgnoreCase))
            return trimmed.Substring(7).Trim();

        if (IsKet(trimmed))
        {
            var info = TryGetMoleculeInfo(trimmed);
            if (!string.IsNullOrEmpty(info?.Smiles))
                return info!.Smiles;
            return null;
        }

        return trimmed;
    }

    public static string? TryConvertToKet(string? input)
    {
        if (string.IsNullOrWhiteSpace(input)) return null;
        var trimmed = input!.Trim();

        if (trimmed.StartsWith("KET:", StringComparison.OrdinalIgnoreCase))
            trimmed = trimmed.Substring(4).Trim();

        if (IsKet(trimmed))
            return StructureMetadata.CompactKet(trimmed);

        if (trimmed.StartsWith("SMILES:", StringComparison.OrdinalIgnoreCase))
            trimmed = trimmed.Substring(7).Trim();

        try
        {
            using var indigo = new Indigo();
            IndigoObject? obj = null;
            try
            {
                obj = indigo.loadMolecule(trimmed);
            }
            catch
            {
                try { obj = indigo.loadReaction(trimmed); } catch { }
            }

            if (obj != null)
            {
                using (obj)
                {
                    var json = obj.json();
                    if (!string.IsNullOrWhiteSpace(json))
                        return StructureMetadata.CompactKet(json);
                }
            }
        }
        catch { }

        return null;
    }

    public static string? TryConvertToInchi(string? input)
    {
        if (string.IsNullOrWhiteSpace(input)) return null;
        var trimmed = input!.Trim();

        if (trimmed.StartsWith("KET:", StringComparison.OrdinalIgnoreCase))
            trimmed = trimmed.Substring(4).Trim();
        else if (trimmed.StartsWith("SMILES:", StringComparison.OrdinalIgnoreCase))
            trimmed = trimmed.Substring(7).Trim();

        try
        {
            using var indigo = new Indigo();
            using var inchi = new IndigoInchi(indigo);
            using var mol = indigo.loadMolecule(trimmed);
            return inchi.getInchi(mol);
        }
        catch
        {
            var smiles = TryConvertToSmiles(trimmed);
            if (!string.IsNullOrEmpty(smiles) && !string.Equals(smiles, trimmed, StringComparison.Ordinal))
            {
                try
                {
                    using var indigo = new Indigo();
                    using var inchi = new IndigoInchi(indigo);
                    using var mol = indigo.loadMolecule(smiles);
                    return inchi.getInchi(mol);
                }
                catch { }
            }
        }

        return null;
    }

    public static string? TryConvertToInchiKey(string? input)
    {
        var inchiStr = TryConvertToInchi(input);
        if (string.IsNullOrEmpty(inchiStr)) return null;

        try
        {
            using var indigo = new Indigo();
            using var inchi = new IndigoInchi(indigo);
            return inchi.getInchiKey(inchiStr);
        }
        catch { }

        return null;
    }

    public static byte[]? RenderKet(string ket, RenderOptions? options = null)
    {
        if (string.IsNullOrWhiteSpace(ket)) return null;
        var cleanKet = CleanKet(ket);
        options ??= new RenderOptions();
        using var indigo = new Indigo();
        using var renderer = new IndigoRenderer(indigo);

        indigo.setOption("render-output-format", "png");
        indigo.setOption("render-bond-length", 80.0);
        indigo.setOption("render-margins", 30, 30);
        indigo.setOption("render-relative-thickness", 1.5);
        indigo.setOption("aromaticity-model", "generic");
        indigo.setOption("render-coloring", options.Color);

        ApplyBackground(indigo, options.Background);

        if (!string.IsNullOrWhiteSpace(options.BaseColor))
            ApplyBaseColor(indigo, options.BaseColor);

        IndigoObject? obj = null;
        try
        {
            obj = indigo.loadMolecule(cleanKet);
        }
        catch
        {
            try { obj = indigo.loadReaction(cleanKet); } catch { }
        }

        if (obj == null) return null;

        using (obj)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(options.Transform))
                    ApplyTransform(obj, options.Transform);
            }
            catch { }

            return renderer.renderToBuffer(obj);
        }
    }

    public static MoleculeInfo? TryGetMoleculeInfo(string? structureOrSmiles)
    {
        if (string.IsNullOrWhiteSpace(structureOrSmiles)) return null;
        var input = structureOrSmiles!.Trim();
        string? fallbackSmiles = null;

        var smilesIdx = input.IndexOf("\nSMILES:", StringComparison.OrdinalIgnoreCase);
        if (smilesIdx >= 0)
        {
            fallbackSmiles = input.Substring(smilesIdx + 8).Trim();
            input = input.Substring(0, smilesIdx).Trim();
        }

        if (input.StartsWith("KET:", StringComparison.OrdinalIgnoreCase))
            input = input.Substring(4).Trim();
        else if (input.StartsWith("SMILES:", StringComparison.OrdinalIgnoreCase))
            input = input.Substring(7).Trim();

        if (string.IsNullOrWhiteSpace(input))
            input = fallbackSmiles ?? "";

        if (string.IsNullOrWhiteSpace(input)) return null;

        var result = TryGetMoleculeInfoCore(input);
        if (result == null && !string.IsNullOrWhiteSpace(fallbackSmiles))
            result = TryGetMoleculeInfoCore(fallbackSmiles!);
        return result;
    }

    private static MoleculeInfo? TryGetMoleculeInfoCore(string input)
    {
        try
        {
            using var indigo = new Indigo();
            try
            {
                using var mol = indigo.loadMolecule(input);
                if (mol.countAtoms() > 0)
                {
                    var rawFormula = "";
                    try { rawFormula = mol.grossFormula()?.Replace(" ", "") ?? ""; } catch { }
                    double mw = 0;
                    try { mw = mol.molecularWeight(); } catch { }
                    double exact = 0;
                    try { exact = mol.monoisotopicMass(); } catch { }
                    var canSmiles = "";
                    try { canSmiles = mol.canonicalSmiles() ?? ""; } catch { try { canSmiles = mol.smiles() ?? ""; } catch { } }

                    string title = "";
                    string iupac = "";
                    if (!string.IsNullOrWhiteSpace(canSmiles))
                    {
                        var resolved = PubChemResolver.ResolveCompound(canSmiles);
                        title = resolved.CompoundName;
                        iupac = resolved.IupacName;
                    }

                    string canInchi = "";
                    string canInchiKey = "";
                    try
                    {
                        using var inchi = new IndigoInchi(indigo);
                        canInchi = inchi.getInchi(mol) ?? "";
                        if (!string.IsNullOrEmpty(canInchi))
                            canInchiKey = inchi.getInchiKey(canInchi) ?? "";
                    }
                    catch { }

                    return new MoleculeInfo
                    {
                        CompoundName = title,
                        IupacName = iupac,
                        Formula = rawFormula,
                        MolecularWeight = mw,
                        ExactMass = exact,
                        Smiles = canSmiles,
                        InChI = canInchi,
                        InChIKey = canInchiKey,
                        IsReaction = false
                    };
                }
            }
            catch { }

            try
            {
                using var rxn = indigo.loadReaction(input);
                var rxnSmiles = "";
                try { rxnSmiles = rxn.smiles() ?? ""; } catch { }
                if (!string.IsNullOrEmpty(rxnSmiles))
                {
                    return new MoleculeInfo
                    {
                        Smiles = rxnSmiles,
                        IsReaction = true
                    };
                }
            }
            catch { }
        }
        catch { }

        return null;
    }

    public static bool TryRenderStructure(string? rawInput, RenderOptions? options, out byte[]? pngBytes, out string? ket, out string? smiles)
    {
        pngBytes = null;
        ket = null;
        smiles = null;

        if (string.IsNullOrWhiteSpace(rawInput)) return false;
        var input = rawInput!.Trim();
        string? embeddedSmiles = null;

        var smilesIdx = input.IndexOf("\nSMILES:", StringComparison.OrdinalIgnoreCase);
        if (smilesIdx >= 0)
        {
            embeddedSmiles = input.Substring(smilesIdx + 8).Trim();
            input = input.Substring(0, smilesIdx).Trim();
        }

        if (input.StartsWith("KET:", StringComparison.OrdinalIgnoreCase))
            input = input.Substring(4).Trim();
        else if (input.StartsWith("SMILES:", StringComparison.OrdinalIgnoreCase))
            input = input.Substring(7).Trim();

        if (string.IsNullOrWhiteSpace(input)) return false;

        options ??= RenderOptions.Default;

        // 1. Check if input is KET JSON
        if (input.StartsWith("{", StringComparison.Ordinal) && input.Contains("\"root\""))
        {
            try
            {
                pngBytes = RenderKet(input, options);
                if (pngBytes != null && pngBytes.Length > 0)
                {
                    ket = input;
                    smiles = !string.IsNullOrWhiteSpace(embeddedSmiles)
                        ? embeddedSmiles
                        : (TryGetMoleculeInfo(input)?.Smiles ?? "");
                    return true;
                }
            }
            catch { }
        }

        // 2. Try standard SMILES rendering
        try
        {
            pngBytes = Render(input, options);
            if (pngBytes != null && pngBytes.Length > 0)
            {
                smiles = Normalize(input);
                return true;
            }
        }
        catch { }

        // 3. Try Reaction SMILES or Molfile via RenderKet
        try
        {
            pngBytes = RenderKet(input, options);
            if (pngBytes != null && pngBytes.Length > 0)
            {
                var info = TryGetMoleculeInfo(input);
                smiles = info?.Smiles ?? input;
                return true;
            }
        }
        catch { }

        return false;
    }

    // Expose explicit H only on stereocenters whose non-H bonds are all part of rings
    // (e.g. ring junctions, bridgeheads). This prevents Indigo from putting wedge bonds
    // directly onto ring bonds while keeping standard skeletal formulas everywhere else.
    private static void ExposeRingJunctionChiralHydrogens(IndigoObject molecule)
    {
        if (molecule.countStereocenters() == 0) return;

        var chainBonds = new Dictionary<int, int>();
        foreach (IndigoObject atom in molecule.iterateAtoms()) chainBonds[atom.index()] = 0;
        foreach (IndigoObject bond in molecule.iterateBonds())
        {
            if (bond.topology() != RingTopology)
            {
                chainBonds[bond.source().index()]++;
                chainBonds[bond.destination().index()]++;
            }
        }

        var junctions = new HashSet<int>();
        foreach (IndigoObject sc in molecule.iterateStereocenters())
        {
            var idx = sc.index();
            if (chainBonds[idx] == 0 && sc.countHydrogens() > 0)
                junctions.Add(idx);
        }

        if (junctions.Count == 0) return;

        molecule.unfoldHydrogens();
        var toRemove = new List<int>();
        foreach (IndigoObject atom in molecule.iterateAtoms())
        {
            if (atom.symbol() == "H")
            {
                var keep = atom.iterateNeighbors().Cast<IndigoObject>().Any(n => junctions.Contains(n.index()));
                if (!keep) toRemove.Add(atom.index());
            }
        }

        toRemove.Sort();
        for (var i = toRemove.Count - 1; i >= 0; i--)
            molecule.getAtom(toRemove[i]).remove();
    }

    private static void ApplyTransform(IndigoObject molecule, string? transform)
    {
        if (string.IsNullOrWhiteSpace(transform)) return;
        var tokens = transform!.ToLowerInvariant().Split(TransformSeparators, StringSplitOptions.RemoveEmptyEntries);
        foreach (var token in tokens)
        {
            switch (token)
            {
                case "flip" or "flipx":
                    foreach (IndigoObject a in molecule.iterateAtoms())
                    {
                        var p = a.xyz();
                        a.setXYZ(-p[0], p[1], p[2]);
                    }
                    break;
                case "flipy":
                    foreach (IndigoObject a in molecule.iterateAtoms())
                    {
                        var p = a.xyz();
                        a.setXYZ(p[0], -p[1], p[2]);
                    }
                    break;
                case "rot90":
                    foreach (IndigoObject a in molecule.iterateAtoms())
                    {
                        var p = a.xyz();
                        a.setXYZ(p[1], -p[0], p[2]);
                    }
                    break;
                case "rot180":
                    foreach (IndigoObject a in molecule.iterateAtoms())
                    {
                        var p = a.xyz();
                        a.setXYZ(-p[0], -p[1], p[2]);
                    }
                    break;
                case "rot270":
                    foreach (IndigoObject a in molecule.iterateAtoms())
                    {
                        var p = a.xyz();
                        a.setXYZ(-p[1], p[0], p[2]);
                    }
                    break;
            }
        }
    }

    private static void LabelHeteroMethyls(IndigoObject molecule)
    {
        if (molecule.countAtoms() <= 4) return;
        foreach (IndigoObject atom in molecule.iterateAtoms())
        {
            if (atom.symbol() == "C" && atom.degree() == 1)
            {
                var neighbor = atom.iterateNeighbors().Cast<IndigoObject>().FirstOrDefault();
                if (neighbor != null)
                {
                    var sym = neighbor.symbol();
                    if (sym is "N" or "O" or "S")
                    {
                        var cPos = atom.xyz();
                        var nPos = neighbor.xyz();
                        atom.resetAtom(cPos[0] < nPos[0] - 0.2f ? "H3C" : "CH3");
                    }
                }
            }
        }
    }

    private static void ApplyBackground(Indigo indigo, string? background)
    {
        if (ColorHelper.TryGetRgb(background, out var r, out var g, out var b))
            indigo.setOption("render-background-color", r, g, b);
        else
            indigo.setOption("render-background-color", -1.0, -1.0, -1.0); // transparent
    }

    private static void ApplyBaseColor(Indigo indigo, string? baseColor)
    {
        if (ColorHelper.TryGetRgb(baseColor, out var r, out var g, out var b))
            indigo.setOption("render-base-color", r, g, b);
    }

    private static IndigoObject Parse(Indigo indigo, string input)
    {
        IndigoObject molecule;
        try { molecule = indigo.loadMolecule(input); }
        catch (IndigoException ex) { throw new InvalidSmilesException("Indigo: " + ex.Message); }
        if (molecule.countAtoms() == 0)
        {
            molecule.Dispose();
            throw Invalid();
        }
        return molecule;
    }

    private static InvalidSmilesException Invalid() => new("Invalid SMILES string.");
}
