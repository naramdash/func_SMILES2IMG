using System;
using System.Collections.Generic;
using System.Linq;
using com.epam.indigo;

namespace Smiles2Img.Rendering;

[Serializable]
internal sealed class InvalidSmilesException(string message) : Exception(message);

internal static class MoleculeRenderer
{
    private const int RingTopology = 10;
    private static readonly char[] TransformSeparators = [' ', ',', '|', ';'];

    internal static string Normalize(string smiles)
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

    internal static byte[] Render(string smiles, bool color = true, string? transform = null, string? background = null)
        => Render(smiles, new RenderOptions { Color = color, Transform = transform, Background = background });

    internal static byte[] Render(string smiles, RenderOptions options)
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
