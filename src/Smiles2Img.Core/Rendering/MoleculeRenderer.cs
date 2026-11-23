using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using com.epam.indigo;

namespace Smiles2Img.Rendering;

public static class MoleculeRenderer
{
    private const int MaxStructureInputLength = 16384;
    private static readonly char[] TransformSeparators = [' ', ',', '|', ';'];
    private static readonly ConcurrentDictionary<string, MoleculeInfo?> _infoCache = new(StringComparer.Ordinal);

    public static bool IsInchi(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return false;
        var trimmed = text!.Trim();
        return trimmed.StartsWith("InChI=", StringComparison.OrdinalIgnoreCase) ||
               trimmed.StartsWith("1S/", StringComparison.OrdinalIgnoreCase) ||
               trimmed.StartsWith("1/", StringComparison.OrdinalIgnoreCase);
    }

    public static string NormalizeInchi(string input)
    {
        var trimmed = input.Trim();
        if (trimmed.StartsWith("InChI=", StringComparison.OrdinalIgnoreCase))
            return trimmed;
        if (trimmed.StartsWith("1S/", StringComparison.OrdinalIgnoreCase) || trimmed.StartsWith("1/", StringComparison.OrdinalIgnoreCase))
            return "InChI=" + trimmed;
        return trimmed;
    }

    public static string Normalize(string smiles)
    {
        var input = smiles?.Trim();
        if (!string.IsNullOrEmpty(input) && input!.StartsWith("SMILES:", StringComparison.OrdinalIgnoreCase))
            input = input.Substring(7).Trim();

        if (string.IsNullOrEmpty(input) || input!.Length > MaxStructureInputLength)
            throw new InvalidSmilesException($"Enter a SMILES or InChI string containing 1 to {MaxStructureInputLength} characters.");

        if (IsInchi(input))
            return NormalizeInchi(input!);

        if (!HasValidSmilesBrackets(input!))
            throw Invalid();

        return input!;
    }

    private static bool IsSingleLineSmilesCandidate(string input)
        => !IsInchi(input) && !IsKet(input) && input.IndexOf('\n') < 0 && input.IndexOf('\r') < 0;

    private static bool HasValidSmilesBrackets(string input)
    {
        if (string.IsNullOrEmpty(input) || input.Length > MaxStructureInputLength)
            return false;

        // Fast syntax validation (bracket and parenthesis balancing)
        var branches = 0;
        var inAtom = false;
        var prev = '\0';
        foreach (var character in input)
        {
            if (character == '[' && !inAtom)
            {
                inAtom = true;
            }
            else if (character == ']' && inAtom)
            {
                inAtom = false;
            }
            else if (character is '[' or ']')
            {
                return false;
            }
            else if (!inAtom && character == '(')
            {
                branches++;
            }
            else if (!inAtom && character == ')')
            {
                if (prev == '(' || --branches < 0)
                    return false;
            }

            prev = character;
        }

        return !inAtom && branches == 0;
    }

    private static string? GetBadValenceError(IndigoObject mol)
    {
        try
        {
            var err = mol.checkBadValence();
            return string.IsNullOrWhiteSpace(err) ? null : err.Trim();
        }
        catch
        {
            return null;
        }
    }

    public static byte[] Render(string smiles, bool color = true, string? transform = null, string? background = null)
        => Render(smiles, new RenderOptions { Color = color, Transform = transform, Background = background });

    public static byte[] Render(string smiles, RenderOptions options)
        => RenderCore(smiles, options, "png");

    private static void ConfigureBaseRendererOptions(Indigo indigo, RenderOptions options, string outputFormat)
    {
        indigo.setOption("render-output-format", outputFormat);
        indigo.setOption("render-bond-length", 80.0);
        indigo.setOption("render-margins", 30, 30);
        indigo.setOption("render-relative-thickness", 1.5);
        indigo.setOption("aromaticity-model", "generic");
        indigo.setOption("render-coloring", options.Color);

        ApplyBackground(indigo, options.Background);

        if (!string.IsNullOrWhiteSpace(options.BaseColor))
            ApplyBaseColor(indigo, options.BaseColor);
    }

    private static string StripStructurePrefixes(string? rawInput, out string? embeddedSmiles)
    {
        embeddedSmiles = null;
        if (string.IsNullOrWhiteSpace(rawInput)) return "";

        var input = rawInput!.Trim();
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

        return input;
    }

    private static byte[] RenderCore(string smiles, RenderOptions options, string outputFormat)
    {
        NativeIndigo.EnsureInitialized();
        options ??= new RenderOptions();
        using var indigo = new Indigo();
        using var renderer = new IndigoRenderer(indigo);

        ConfigureBaseRendererOptions(indigo, options, outputFormat);
        indigo.setOption("smart-layout", "true");
        indigo.setOption("layout-orientation", "horizontal");

        if (options.AtomNumbers)
        {
            indigo.setOption("render-atom-ids-visible", "true");
            indigo.setOption("render-atom-bond-ids-from-one", "true");
        }

        indigo.setOption("render-stereo-style", options.StereoLabels ? "ext" : "none");

        using var molecule = Parse(indigo, Normalize(smiles));

        if (!string.IsNullOrWhiteSpace(options.HighlightQuery))
            ApplySubstructureHighlight(indigo, molecule, options.HighlightQuery);

        if (options.UnfoldHydrogens)
            molecule.unfoldHydrogens();

        if (options.AllCarbons)
            indigo.setOption("render-label-mode", "all");
        else if (molecule.countAtoms() <= 4)
            indigo.setOption("render-label-mode", "terminal-hetero");
        else
            indigo.setOption("render-label-mode", "hetero");

        molecule.dearomatize();
        molecule.layout();
        ApplyTransform(molecule, options.Transform);

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
        {
            trimmed = trimmed.Substring(7).Trim();
            if (string.IsNullOrEmpty(trimmed)) return null;
        }

        if (IsKet(trimmed))
        {
            var info = TryGetMoleculeInfo(trimmed, resolvePubChem: false);
            if (!string.IsNullOrEmpty(info?.Smiles))
                return info!.Smiles;
            return null;
        }

        if (IsInchi(trimmed))
        {
            NativeIndigo.EnsureInitialized();
            try
            {
                using var indigo = new Indigo();
                using var inchi = new IndigoInchi(indigo);
                using var mol = inchi.loadMolecule(NormalizeInchi(trimmed));
                if (mol.countAtoms() <= 0) return null;
                return mol.canonicalSmiles() ?? mol.smiles();
            }
            catch { }
            return null;
        }

        if (IsSingleLineSmilesCandidate(trimmed) && !HasValidSmilesBrackets(trimmed))
            return null;

        NativeIndigo.EnsureInitialized();
        try
        {
            using var indigo = new Indigo();
            using var mol = indigo.loadMolecule(trimmed);
            if (mol.countAtoms() <= 0 || GetBadValenceError(mol) != null) return null;
            return mol.canonicalSmiles() ?? mol.smiles();
        }
        catch { }

        return null;
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

        if (IsSingleLineSmilesCandidate(trimmed) && !HasValidSmilesBrackets(trimmed))
            return null;

        NativeIndigo.EnsureInitialized();
        try
        {
            using var indigo = new Indigo();
            IndigoObject? obj = null;
            var isReaction = false;
            try
            {
                if (IsInchi(trimmed))
                {
                    using var inchi = new IndigoInchi(indigo);
                    obj = inchi.loadMolecule(NormalizeInchi(trimmed));
                }
                else
                {
                    obj = indigo.loadMolecule(trimmed);
                }
            }
            catch
            {
                try { obj = indigo.loadReaction(trimmed); isReaction = true; } catch { }
            }

            if (obj != null)
            {
                using (obj)
                {
                    if (!isReaction && !IsKet(trimmed) && GetBadValenceError(obj) != null)
                        return null;

                    try
                    {
                        if (!obj.hasCoord())
                            obj.layout();
                    }
                    catch
                    {
                        try { obj.layout(); } catch { }
                    }

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

        if (IsInchi(trimmed))
        {
            var normalized = NormalizeInchi(trimmed);
            NativeIndigo.EnsureInitialized();
            try
            {
                using var indigo = new Indigo();
                using var inchi = new IndigoInchi(indigo);
                using var mol = inchi.loadMolecule(normalized);
                return mol.countAtoms() > 0 ? normalized : null;
            }
            catch
            {
                return null;
            }
        }

        if (trimmed.StartsWith("KET:", StringComparison.OrdinalIgnoreCase))
            trimmed = trimmed.Substring(4).Trim();
        else if (trimmed.StartsWith("SMILES:", StringComparison.OrdinalIgnoreCase))
            trimmed = trimmed.Substring(7).Trim();

        if (IsSingleLineSmilesCandidate(trimmed) && !HasValidSmilesBrackets(trimmed))
            return null;

        NativeIndigo.EnsureInitialized();
        try
        {
            using var indigo = new Indigo();
            using var inchi = new IndigoInchi(indigo);
            using var mol = indigo.loadMolecule(trimmed);
            if (mol.countAtoms() <= 0 || GetBadValenceError(mol) != null) return null;
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

        NativeIndigo.EnsureInitialized();
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
        => RenderKetCore(ket, options, "png");

    private static byte[]? RenderKetCore(string ket, RenderOptions? options, string outputFormat)
    {
        if (string.IsNullOrWhiteSpace(ket)) return null;
        var cleanKet = CleanKet(ket);
        NativeIndigo.EnsureInitialized();
        options ??= new RenderOptions();
        using var indigo = new Indigo();
        using var renderer = new IndigoRenderer(indigo);

        ConfigureBaseRendererOptions(indigo, options, outputFormat);

        if (options.AtomNumbers)
        {
            indigo.setOption("render-atom-ids-visible", "true");
            indigo.setOption("render-atom-bond-ids-from-one", "true");
        }

        indigo.setOption("render-stereo-style", options.StereoLabels ? "ext" : "none");

        if (options.AllCarbons)
            indigo.setOption("render-label-mode", "all");

        IndigoObject? obj = null;
        var isReaction = false;
        try
        {
            obj = indigo.loadMolecule(cleanKet);
        }
        catch
        {
            try
            {
                obj = indigo.loadReaction(cleanKet);
                isReaction = true;
            }
            catch { }
        }

        if (obj == null) return null;

        using (obj)
        {
            if (options.UnfoldHydrogens)
            {
                try { obj.unfoldHydrogens(); } catch { }
            }

            if (!isReaction && !string.IsNullOrWhiteSpace(options.HighlightQuery))
            {
                ApplySubstructureHighlight(indigo, obj, options.HighlightQuery);
                try { obj.dearomatize(); } catch { }
            }

            try
            {
                if (!string.IsNullOrWhiteSpace(options.Transform))
                    ApplyTransform(obj, options.Transform);
            }
            catch { }

            return renderer.renderToBuffer(obj);
        }
    }

    public static string? RenderSvg(string? rawInput, RenderOptions? options = null)
    {
        var input = StripStructurePrefixes(rawInput, out _);
        if (string.IsNullOrWhiteSpace(input)) return null;
        options ??= RenderOptions.Default;

        if (input.StartsWith("{", StringComparison.Ordinal) && input.Contains("\"root\""))
        {
            try
            {
                var bytes = RenderKetCore(input, options, "svg");
                if (bytes != null && bytes.Length > 0)
                    return Encoding.UTF8.GetString(bytes);
            }
            catch { }
        }

        try
        {
            var bytes = RenderCore(input, options, "svg");
            if (bytes != null && bytes.Length > 0)
                return Encoding.UTF8.GetString(bytes);
        }
        catch { }

        try
        {
            var bytes = RenderKetCore(input, options, "svg");
            if (bytes != null && bytes.Length > 0)
                return Encoding.UTF8.GetString(bytes);
        }
        catch { }

        return null;
    }

    public static bool? IsSubstructureMatch(string? targetInput, string? queryInput)
    {
        if (string.IsNullOrWhiteSpace(targetInput) || string.IsNullOrWhiteSpace(queryInput))
            return null;

        NativeIndigo.EnsureInitialized();
        try
        {
            using var indigo = new Indigo();
            indigo.setOption("aromaticity-model", "generic");
            using var targetMol = TryLoadMoleculeForAnalysis(indigo, targetInput!);
            if (targetMol == null) return null;

            try { targetMol.dearomatize(); } catch { }
            try { targetMol.aromatize(); } catch { }

            using var queryObj = TryLoadQueryForMatching(indigo, queryInput!.Trim());
            if (queryObj == null) return null;

            using var matcher = indigo.substructureMatcher(targetMol);
            using var match = matcher.match(queryObj);
            return match != null;
        }
        catch
        {
            return null;
        }
    }

    private static IndigoObject? TryLoadMoleculeForAnalysis(Indigo indigo, string rawInput)
    {
        var input = StripStructurePrefixes(rawInput, out _);
        if (string.IsNullOrWhiteSpace(input)) return null;
        if (IsSingleLineSmilesCandidate(input) && !HasValidSmilesBrackets(input)) return null;

        try
        {
            if (IsInchi(input))
            {
                using var inchi = new IndigoInchi(indigo);
                var m = inchi.loadMolecule(NormalizeInchi(input));
                if (m.countAtoms() > 0) return m;
                m.Dispose();
                return null;
            }
            else
            {
                var m = indigo.loadMolecule(input);
                if (m.countAtoms() > 0)
                {
                    if (!IsKet(input) && GetBadValenceError(m) != null)
                    {
                        m.Dispose();
                        return null;
                    }
                    return m;
                }
                m.Dispose();
                return null;
            }
        }
        catch
        {
            return null;
        }
    }

    private static IndigoObject? TryLoadQueryForMatching(Indigo indigo, string query)
    {
        try
        {
            return indigo.loadSmarts(query);
        }
        catch
        {
            try
            {
                var q = indigo.loadQueryMolecule(query);
                try { q.aromatize(); } catch { }
                return q;
            }
            catch
            {
                return null;
            }
        }
    }

    private static void ApplySubstructureHighlight(Indigo indigo, IndigoObject molecule, string? query)
    {
        if (string.IsNullOrWhiteSpace(query)) return;
        try
        {
            indigo.setOption("aromaticity-model", "generic");
            indigo.setOption("render-highlight-color-enabled", true);
            indigo.setOption("render-highlight-color", 0.92f, 0.18f, 0.18f);
            indigo.setOption("render-highlight-thickness-enabled", true);

            try { molecule.dearomatize(); } catch { }
            try { molecule.aromatize(); } catch { }
            using var queryObj = TryLoadQueryForMatching(indigo, query!.Trim());
            if (queryObj == null) return;

            using var matcher = indigo.substructureMatcher(molecule);
            foreach (IndigoObject m in matcher.iterateMatches(queryObj))
            {
                using (m)
                {
                    foreach (IndigoObject qa in queryObj.iterateAtoms())
                    {
                        try { m.mapAtom(qa)?.highlight(); } catch { }
                    }
                    foreach (IndigoObject qb in queryObj.iterateBonds())
                    {
                        try { m.mapBond(qb)?.highlight(); } catch { }
                    }
                }
            }
        }
        catch { }
    }

    public static MoleculeInfo? TryGetMoleculeInfo(string? structureOrSmiles, bool resolvePubChem = false)
    {
        var input = StripStructurePrefixes(structureOrSmiles, out var fallbackSmiles);
        if (string.IsNullOrWhiteSpace(input))
            input = fallbackSmiles ?? "";

        if (string.IsNullOrWhiteSpace(input)) return null;

        var result = TryGetMoleculeInfoCore(input, resolvePubChem);
        if (result == null && !string.IsNullOrWhiteSpace(fallbackSmiles))
            result = TryGetMoleculeInfoCore(fallbackSmiles!, resolvePubChem);
        return result;
    }

    private static MoleculeInfo? TryGetMoleculeInfoCore(string input, bool resolvePubChem)
    {
        if (!resolvePubChem && _infoCache.TryGetValue(input, out var cached))
            return cached;

        var computed = ComputeMoleculeInfoCore(input, resolvePubChem);
        if (!resolvePubChem)
        {
            if (_infoCache.Count > 2048) _infoCache.Clear();
            _infoCache[input] = computed;
        }
        return computed;
    }

    private static MoleculeInfo? ComputeMoleculeInfoCore(string input, bool resolvePubChem)
    {
        if (IsSingleLineSmilesCandidate(input) && !HasValidSmilesBrackets(input))
            return null;

        NativeIndigo.EnsureInitialized();
        try
        {
            using var indigo = new Indigo();
            indigo.setOption("aromaticity-model", "generic");
            try
            {
                IndigoObject? mol = null;
                IndigoInchi? inchi = null;
                try
                {
                    if (IsInchi(input))
                    {
                        inchi = new IndigoInchi(indigo);
                        mol = inchi.loadMolecule(NormalizeInchi(input));
                    }
                    else
                    {
                        mol = indigo.loadMolecule(input);
                    }
                }
                catch { }

                bool hasAtoms = false;
                try { hasAtoms = mol != null && mol.countAtoms() > 0; } catch { }
                if (!hasAtoms)
                {
                    try { mol?.Dispose(); } catch { }
                    try { inchi?.Dispose(); } catch { }
                }
                else if (mol != null)
                {
                    using (mol)
                    {
                        if (!IsKet(input) && GetBadValenceError(mol) != null)
                        {
                            try { inchi?.Dispose(); } catch { }
                            return null;
                        }

                        var hasQueryOrDummyAtoms = false;
                        foreach (IndigoObject atom in mol.iterateAtoms())
                        {
                            using (atom)
                            {
                                if (atom.isPseudoatom() || atom.isRSite() || atom.isTemplateAtom())
                                {
                                    hasQueryOrDummyAtoms = true;
                                    break;
                                }
                            }
                        }

                        var rawFormula = "";
                        double mw = 0;
                        double exact = 0;
                        int heavyAtoms = 0;
                        int rings = 0;
                        double logP = 0;
                        bool hasLogP = false;
                        double tpsa = 0;
                        bool hasTpsa = false;
                        double tpsaWithSP = 0;
                        bool hasTpsaWithSP = false;
                        int aromaticRings = 0;
                        double fsp3 = 0.0;
                        var canSmiles = "";

                        if (hasQueryOrDummyAtoms)
                        {
                            try { canSmiles = mol.canonicalSmiles() ?? mol.smiles() ?? ""; }
                            catch { try { canSmiles = mol.smiles() ?? ""; } catch { } }
                        }
                        else
                        {
                            try { rawFormula = mol.grossFormula()?.Replace(" ", "") ?? ""; } catch { }
                            try { mw = Math.Round(mol.molecularWeight(), 4, MidpointRounding.AwayFromZero); } catch { }
                            try { exact = mol.monoisotopicMass(); } catch { }
                            try { heavyAtoms = mol.countHeavyAtoms(); } catch { }
                            try { rings = mol.countSSSR(); } catch { }

                            using var analysisMol = mol.clone();
                            try { analysisMol.dearomatize(); } catch { }
                            try { analysisMol.aromatize(); } catch { }

                            try
                            {
                                // Bypass Indigo.Net checkResult(double) negative-return bug via direct IndigoLib + GetLastIndigoError.
                                indigo.setSessionID();
                                IndigoLib.indigoSetErrorMessage("");
                                double rawLogP = IndigoLib.indigoLogP(analysisMol.self);
                                string err = NativeIndigo.GetLastIndigoError();
                                if (string.IsNullOrEmpty(err) && !double.IsNaN(rawLogP) && !double.IsInfinity(rawLogP))
                                {
                                    logP = Math.Round(rawLogP, 2, MidpointRounding.AwayFromZero);
                                    if (logP == 0) logP = 0.0;
                                    hasLogP = true;
                                }
                            }
                            catch { }

                            try { tpsa = Math.Round((double)analysisMol.tpsa(false), 2, MidpointRounding.AwayFromZero); hasTpsa = true; } catch { }
                            try { tpsaWithSP = Math.Round((double)analysisMol.tpsa(true), 2, MidpointRounding.AwayFromZero); hasTpsaWithSP = true; } catch { }

                            try
                            {
                                foreach (IndigoObject ring in analysisMol.iterateSSSR())
                                {
                                    using (ring)
                                    {
                                        bool allAromatic = true;
                                        int bondCount = 0;
                                        foreach (IndigoObject b in ring.iterateBonds())
                                        {
                                            using (b)
                                            {
                                                bondCount++;
                                                if (b.bondOrder() != 4)
                                                {
                                                    allAromatic = false;
                                                    break;
                                                }
                                            }
                                        }
                                        if (bondCount > 0 && allAromatic)
                                            aromaticRings++;
                                    }
                                }

                                int totalCarbons = 0;
                                int sp3Carbons = 0;
                                foreach (IndigoObject atom in analysisMol.iterateAtoms())
                                {
                                    using (atom)
                                    {
                                        if (!atom.isPseudoatom() && !atom.isRSite() && !atom.isTemplateAtom() && atom.atomicNumber() == 6)
                                        {
                                            totalCarbons++;
                                            if (string.Equals(atom.getHybridizationStr(), "SP3", StringComparison.OrdinalIgnoreCase))
                                                sp3Carbons++;
                                        }
                                    }
                                }
                                if (totalCarbons > 0)
                                    fsp3 = Math.Round((double)sp3Carbons / totalCarbons, 2, MidpointRounding.AwayFromZero);
                            }
                            catch { }

                            try { canSmiles = analysisMol.canonicalSmiles() ?? mol.canonicalSmiles() ?? ""; }
                            catch { try { canSmiles = mol.canonicalSmiles() ?? mol.smiles() ?? ""; } catch { } }
                        }

                        string title = "";
                        string iupac = "";
                        if (!hasQueryOrDummyAtoms && resolvePubChem && !string.IsNullOrWhiteSpace(canSmiles))
                        {
                            var resolved = PubChemResolver.ResolveCompound(canSmiles);
                            title = resolved.CompoundName;
                            iupac = resolved.IupacName;
                        }

                        string canInchi = "";
                        string canInchiKey = "";
                        try
                        {
                            if (IsInchi(input))
                            {
                                canInchi = NormalizeInchi(input);
                                inchi ??= new IndigoInchi(indigo);
                                canInchiKey = inchi.getInchiKey(canInchi) ?? "";
                            }
                            else if (!hasQueryOrDummyAtoms)
                            {
                                inchi ??= new IndigoInchi(indigo);
                                canInchi = inchi.getInchi(mol) ?? "";
                                if (!string.IsNullOrEmpty(canInchi))
                                    canInchiKey = inchi.getInchiKey(canInchi) ?? "";
                            }
                        }
                        catch { }
                        finally
                        {
                            inchi?.Dispose();
                        }

                        return new MoleculeInfo
                        {
                            CompoundName = title,
                            IupacName = iupac,
                            HasQueryOrDummyAtoms = hasQueryOrDummyAtoms,
                            Formula = rawFormula,
                            MolecularWeight = mw,
                            ExactMass = exact,
                            LogP = logP,
                            HasLogP = hasLogP,
                            Tpsa = tpsa,
                            HasTpsa = hasTpsa,
                            TpsaWithSP = tpsaWithSP,
                            HasTpsaWithSP = hasTpsaWithSP,
                            HeavyAtomCount = heavyAtoms,
                            RingCount = rings,
                            AromaticRingCount = aromaticRings,
                            Fsp3 = fsp3,
                            Smiles = canSmiles,
                            InChI = canInchi,
                            InChIKey = canInchiKey,
                            IsReaction = false
                        };
                    }
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

        var input = StripStructurePrefixes(rawInput, out var embeddedSmiles);
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
                smiles = IsInchi(input) ? (TryConvertToSmiles(input) ?? NormalizeInchi(input)) : Normalize(input);
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

    internal static void ApplyTransform(IndigoObject molecule, string? transform)
    {
        if (string.IsNullOrWhiteSpace(transform)) return;
        var tokens = transform!.ToLowerInvariant().Split(TransformSeparators, StringSplitOptions.RemoveEmptyEntries);
        foreach (var token in tokens)
        {
            switch (token)
            {
                case "rot90":
                    foreach (IndigoObject a in molecule.iterateAtoms())
                    {
                        using (a)
                        {
                            var p = a.xyz();
                            a.setXYZ(p[1], -p[0], p[2]);
                        }
                    }
                    break;
                case "rot180":
                    foreach (IndigoObject a in molecule.iterateAtoms())
                    {
                        using (a)
                        {
                            var p = a.xyz();
                            a.setXYZ(-p[0], -p[1], p[2]);
                        }
                    }
                    break;
                case "rot270":
                    foreach (IndigoObject a in molecule.iterateAtoms())
                    {
                        using (a)
                        {
                            var p = a.xyz();
                            a.setXYZ(-p[1], p[0], p[2]);
                        }
                    }
                    break;
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
        try
        {
            if (IsInchi(input))
            {
                using var inchi = new IndigoInchi(indigo);
                molecule = inchi.loadMolecule(NormalizeInchi(input));
            }
            else
            {
                molecule = indigo.loadMolecule(input);
            }
        }
        catch (IndigoException ex) { throw new InvalidSmilesException("Indigo: " + ex.Message); }
        if (molecule.countAtoms() == 0)
        {
            molecule.Dispose();
            throw Invalid();
        }
        var badValence = GetBadValenceError(molecule);
        if (badValence != null)
        {
            molecule.Dispose();
            throw new InvalidSmilesException("Indigo: " + badValence);
        }
        return molecule;
    }

    private static InvalidSmilesException Invalid() => new("Invalid SMILES or InChI string.");
}
