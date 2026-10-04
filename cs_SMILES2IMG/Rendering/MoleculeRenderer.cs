using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using com.epam.indigo;

namespace Smiles2Img.Rendering;

internal sealed class InvalidSmilesException(string message) : Exception(message);

internal static class MoleculeRenderer
{
    private const int RingTopology = 10;

    internal static string Normalize(string smiles)
    {
        var input = smiles?.Trim();
        if (string.IsNullOrEmpty(input) || input!.Length > 2000)
            throw new InvalidSmilesException("Enter a SMILES string containing 1 to 2000 characters.");
        var branches = 0;
        var inAtom = false;
        foreach (var character in input)
        {
            if (character == '[' && !inAtom) inAtom = true;
            else if (character == ']' && inAtom) inAtom = false;
            else if (character == '[' || character == ']') throw Invalid();
            else if (!inAtom && character == '(') branches++;
            else if (!inAtom && character == ')' && --branches < 0) throw Invalid();
        }
        if (inAtom || branches != 0) throw Invalid();
        return input;
    }

    internal static byte[] Render(string smiles, bool color = true, string? transform = null, string? background = null)
    {
        using var indigo = new Indigo();
        using var renderer = new IndigoRenderer(indigo);
        indigo.setOption("render-output-format", "png");
        indigo.setOption("render-bond-length", 80.0);
        indigo.setOption("render-margins", 30, 30);
        indigo.setOption("render-relative-thickness", 1.5);
        ApplyBackground(indigo, background);
        indigo.setOption("render-coloring", color);
        indigo.setOption("render-stereo-style", "none");
        indigo.setOption("aromaticity-model", "generic");
        indigo.setOption("smart-layout", "true");
        indigo.setOption("layout-orientation", "horizontal");

        using var molecule = Parse(indigo, Normalize(smiles));
        ExposeRingJunctionChiralHydrogens(molecule);
        if (molecule.countAtoms() <= 4)
            indigo.setOption("render-label-mode", "terminal-hetero");
        else
            indigo.setOption("render-label-mode", "hetero");

        molecule.dearomatize();
        molecule.layout();
        ApplyTransform(molecule, transform);
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
            {
                junctions.Add(idx);
            }
        }

        if (junctions.Count == 0) return;

        molecule.unfoldHydrogens();
        var toRemove = new List<int>();
        foreach (IndigoObject atom in molecule.iterateAtoms())
        {
            if (atom.symbol() == "H")
            {
                var keep = false;
                foreach (IndigoObject neighbor in atom.iterateNeighbors())
                {
                    if (junctions.Contains(neighbor.index())) { keep = true; break; }
                }
                if (!keep) toRemove.Add(atom.index());
            }
        }

        toRemove.Sort();
        for (var i = toRemove.Count - 1; i >= 0; i--)
        {
            molecule.getAtom(toRemove[i]).remove();
        }
    }

    private static void ApplyTransform(IndigoObject molecule, string? transform)
    {
        if (string.IsNullOrWhiteSpace(transform)) return;
        var tokens = transform!.ToLowerInvariant().Split(new[] { ' ', ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (var token in tokens)
        {
            switch (token)
            {
                case "flip":
                case "flipx":
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
                    if (sym == "N" || sym == "O" || sym == "S")
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
        if (string.IsNullOrWhiteSpace(background))
        {
            indigo.setOption("render-background-color", 1.0, 1.0, 1.0);
            return;
        }

        var bg = background!.Trim().ToLowerInvariant();
        if (bg is "transparent" or "trans" or "none" or "clear" or "nobg")
        {
            indigo.setOption("render-background-color", -1.0, -1.0, -1.0);
            return;
        }

        try
        {
            var c = ColorTranslator.FromHtml(background);
            indigo.setOption("render-background-color", c.R / 255.0, c.G / 255.0, c.B / 255.0);
        }
        catch
        {
            indigo.setOption("render-background-color", 1.0, 1.0, 1.0);
        }
    }

    private static IndigoObject Parse(Indigo indigo, string input)
    {
        IndigoObject molecule;
        try { molecule = indigo.loadMolecule(input); }
        catch (IndigoException) { throw Invalid(); }
        if (molecule.countAtoms() == 0)
        {
            molecule.Dispose();
            throw Invalid();
        }
        return molecule;
    }

    private static InvalidSmilesException Invalid() => new("Invalid SMILES string.");
}
