using System;
using System.Runtime.CompilerServices;
using ExcelDna.Integration;
using ExcelDna.Logging;
using Smiles2Img.Rendering;

[assembly: InternalsVisibleTo("Smiles2Img.Tests")]

namespace Smiles2Img;

public static class Functions
{
    [ExcelFunction(Name = "SMILES2IMG", Category = "SMILES",
        Description = "[Excel 2024 / Microsoft 365] Renders a high-resolution molecular structure image into the cell. On Excel 2016/2019/2021 use SMILES2IMG.FLOAT.", IsMacroType = true)]
    public static object Img(
        [ExcelArgument(Name = "smiles", Description = "A SMILES string, Ketcher (KET) string, or cell reference (e.g., A2, \"CCO\").")] string smiles,
        [ExcelArgument(Name = "background", Description = "Optional: \"trans\" (default), \"white\", CSS color (e.g. \"yellow\"), or hex (#RRGGBB).")] object? background = null,
        [ExcelArgument(Name = "style", Description = "Optional: \"color\" (default), \"bw\", \"num\", \"stereo\", \"h\", \"all-c\", \"white\", combined (e.g. \"bw num\" or \"bw|stereo\").")] object? style = null,
        [ExcelArgument(Name = "transform", Description = "Optional: orientation (flip, flipy, rot90, rot180, rot270).")] object? transform = null)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "SMILES2IMG";
        if (XlCall.Excel(XlCall.xlfCaller) is not ExcelReference caller ||
            caller.RowFirst != caller.RowLast || caller.ColumnFirst != caller.ColumnLast)
            return ExcelError.ExcelErrorValue;

        if (CellImages.NativeImagesUnsupported) return CellImages.UnsupportedMessage;

        try
        {
            var options = ParseOptions(background, style, transform);

            if (MoleculeRenderer.IsKet(smiles))
            {
                var cleanKet = MoleculeRenderer.CleanKet(smiles);
                if (CellImages.Find(caller, cleanKet, options) is { } cachedKet) return cachedKet;

                var pngKet = MoleculeRenderer.RenderKet(cleanKet, options);
                if (pngKet == null) return ExcelError.ExcelErrorValue;

                CellImageUpdates.Queue(caller, cleanKet, options, pngKet);
                return "";
            }

            var input = MoleculeRenderer.Normalize(smiles);
            if (CellImages.Find(caller, input, options) is { } image) return image;

            var png = MoleculeRenderer.Render(input, options);
            CellImageUpdates.Queue(caller, input, options, png);
            return "";
        }
        catch (InvalidSmilesException ex)
        {
            LogDisplay.WriteLine("SMILES2IMG: " + ex.Message);
            return ExcelError.ExcelErrorValue;
        }
        catch (Exception error)
        {
            LogDisplay.WriteLine("SMILES2IMG: " + error.Message);
            return ExcelError.ExcelErrorNA;
        }
    }

    [ExcelFunction(Name = "KET2IMG", Category = "SMILES",
        Description = "[Excel 2024 / Microsoft 365] Renders a Ketcher (KET) structure/reaction image into the cell. On Excel 2016/2019/2021 use KET2IMG.FLOAT.", IsMacroType = true)]
    public static object KetImg(
        [ExcelArgument(Name = "ket", Description = "A Ketcher JSON string or cell reference containing one (e.g., A2, \"{...}\").")] string ket,
        [ExcelArgument(Name = "background", Description = "Optional: \"trans\" (default), \"white\", CSS color (e.g. \"yellow\"), or hex (#RRGGBB).")] object? background = null,
        [ExcelArgument(Name = "style", Description = "Optional: \"color\" (default), \"bw\", \"white\", combined (e.g. \"bw\").")] object? style = null,
        [ExcelArgument(Name = "transform", Description = "Optional: orientation (flip, flipy, rot90, rot180, rot270).")] object? transform = null)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "KET2IMG";
        if (XlCall.Excel(XlCall.xlfCaller) is not ExcelReference caller ||
            caller.RowFirst != caller.RowLast || caller.ColumnFirst != caller.ColumnLast)
            return ExcelError.ExcelErrorValue;

        if (CellImages.NativeImagesUnsupported) return CellImages.UnsupportedMessage;

        try
        {
            if (string.IsNullOrWhiteSpace(ket)) return ExcelError.ExcelErrorValue;
            var options = ParseOptions(background, style, transform);
            var cleanKet = MoleculeRenderer.CleanKet(ket);
            if (CellImages.Find(caller, cleanKet, options) is { } image) return image;

            var png = MoleculeRenderer.RenderKet(cleanKet, options);
            if (png == null) return ExcelError.ExcelErrorValue;

            CellImageUpdates.Queue(caller, cleanKet, options, png);
            return "";
        }
        catch (Exception error)
        {
            LogDisplay.WriteLine("KET2IMG: " + error.Message);
            return ExcelError.ExcelErrorValue;
        }
    }

    [ExcelFunction(Name = "SMILES2NAME", Category = "SMILES",
        Description = "Resolves chemical substance name from PubChem for a SMILES string, KET, or cell reference.", IsMacroType = false)]
    public static object SmilesName(
        [ExcelArgument(Name = "smiles", Description = "A SMILES string, KET string, or cell reference.")] string smiles)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "SMILES2NAME";
        if (string.IsNullOrWhiteSpace(smiles)) return ExcelError.ExcelErrorValue;
        try
        {
            var info = MoleculeRenderer.TryGetMoleculeInfo(smiles, resolvePubChem: true);
            if (info != null && !string.IsNullOrWhiteSpace(info.CompoundName))
                return info.CompoundName;
            return ExcelError.ExcelErrorNA;
        }
        catch
        {
            return ExcelError.ExcelErrorValue;
        }
    }

    [ExcelFunction(Name = "SMILES2MW", Category = "SMILES",
        Description = "Calculates molecular weight (MW) in g/mol for a SMILES string, KET, or cell reference.", IsMacroType = false)]
    public static object SmilesMW(
        [ExcelArgument(Name = "smiles", Description = "A SMILES string, KET string, or cell reference.")] string smiles)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "SMILES2MW";
        if (string.IsNullOrWhiteSpace(smiles)) return ExcelError.ExcelErrorValue;
        try
        {
            var info = MoleculeRenderer.TryGetMoleculeInfo(smiles, resolvePubChem: false);
            if (info != null && !info.IsReaction && info.MolecularWeight > 0)
                return Math.Round(info.MolecularWeight, 2);
            return ExcelError.ExcelErrorNA;
        }
        catch
        {
            return ExcelError.ExcelErrorValue;
        }
    }

    [ExcelFunction(Name = "SMILES2FORMULA", Category = "SMILES",
        Description = "Calculates Hill chemical formula for a SMILES string, KET, or cell reference.", IsMacroType = false)]
    public static object SmilesFormula(
        [ExcelArgument(Name = "smiles", Description = "A SMILES string, KET string, or cell reference.")] string smiles)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "SMILES2FORMULA";
        if (string.IsNullOrWhiteSpace(smiles)) return ExcelError.ExcelErrorValue;
        try
        {
            var info = MoleculeRenderer.TryGetMoleculeInfo(smiles, resolvePubChem: false);
            if (info != null && !info.IsReaction && !string.IsNullOrWhiteSpace(info.Formula))
                return info.Formula;
            return ExcelError.ExcelErrorNA;
        }
        catch
        {
            return ExcelError.ExcelErrorValue;
        }
    }

    [ExcelFunction(Name = "KET2NAME", Category = "SMILES",
        Description = "Resolves chemical substance name from PubChem for a Ketcher (KET) structure.", IsMacroType = false)]
    public static object KetName(
        [ExcelArgument(Name = "ket", Description = "A Ketcher JSON string or cell reference.")] string ket)
        => SmilesName(ket);

    [ExcelFunction(Name = "KET2MW", Category = "SMILES",
        Description = "Calculates molecular weight (MW) in g/mol for a Ketcher (KET) structure.", IsMacroType = false)]
    public static object KetMW(
        [ExcelArgument(Name = "ket", Description = "A Ketcher JSON string or cell reference.")] string ket)
        => SmilesMW(ket);

    [ExcelFunction(Name = "KET2FORMULA", Category = "SMILES",
        Description = "Calculates Hill chemical formula for a Ketcher (KET) structure.", IsMacroType = false)]
    public static object KetFormula(
        [ExcelArgument(Name = "ket", Description = "A Ketcher JSON string or cell reference.")] string ket)
        => SmilesFormula(ket);

    [ExcelFunction(Name = "SMILES2INCHI", Category = "SMILES",
        Description = "Converts a SMILES string, KET, or cell reference to an IUPAC InChI string.", IsMacroType = false)]
    public static object SmilesInchi(
        [ExcelArgument(Name = "smiles", Description = "A SMILES string, KET string, or cell reference.")] string smiles)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "SMILES2INCHI";
        if (string.IsNullOrWhiteSpace(smiles)) return ExcelError.ExcelErrorValue;
        try
        {
            var inchi = MoleculeRenderer.TryConvertToInchi(smiles);
            if (!string.IsNullOrEmpty(inchi))
                return inchi!;
            return ExcelError.ExcelErrorNA;
        }
        catch
        {
            return ExcelError.ExcelErrorValue;
        }
    }

    [ExcelFunction(Name = "SMILES2INCHIKEY", Category = "SMILES",
        Description = "Calculates hashed IUPAC InChIKey (27 characters) for a SMILES string, KET, or cell reference.", IsMacroType = false)]
    public static object SmilesInchiKey(
        [ExcelArgument(Name = "smiles", Description = "A SMILES string, KET string, or cell reference.")] string smiles)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "SMILES2INCHIKEY";
        if (string.IsNullOrWhiteSpace(smiles)) return ExcelError.ExcelErrorValue;
        try
        {
            var key = MoleculeRenderer.TryConvertToInchiKey(smiles);
            if (!string.IsNullOrEmpty(key))
                return key!;
            return ExcelError.ExcelErrorNA;
        }
        catch
        {
            return ExcelError.ExcelErrorValue;
        }
    }

    [ExcelFunction(Name = "KET2INCHI", Category = "SMILES",
        Description = "Converts a Ketcher (KET) structure to an IUPAC InChI string.", IsMacroType = false)]
    public static object KetInchi(
        [ExcelArgument(Name = "ket", Description = "A Ketcher JSON string or cell reference.")] string ket)
        => SmilesInchi(ket);

    [ExcelFunction(Name = "KET2INCHIKEY", Category = "SMILES",
        Description = "Calculates hashed IUPAC InChIKey for a Ketcher (KET) structure.", IsMacroType = false)]
    public static object KetInchiKey(
        [ExcelArgument(Name = "ket", Description = "A Ketcher JSON string or cell reference.")] string ket)
        => SmilesInchiKey(ket);

    [ExcelFunction(Name = "INCHI2IMG", Category = "SMILES",
        Description = "[Excel 2024 / Microsoft 365] Renders an InChI molecular structure image into the cell. On Excel 2016/2019/2021 use INCHI2IMG.FLOAT.", IsMacroType = true)]
    public static object InchiImg(
        [ExcelArgument(Name = "inchi", Description = "An IUPAC InChI string or cell reference (e.g., A2, \"InChI=1S/...\").")] string inchi,
        [ExcelArgument(Name = "background", Description = "Optional: \"trans\" (default), \"white\", CSS color (e.g. \"yellow\"), or hex (#RRGGBB).")] object? background = null,
        [ExcelArgument(Name = "style", Description = "Optional: \"color\" (default), \"bw\", \"num\", \"stereo\", \"h\", \"all-c\", \"white\", combined.")] object? style = null,
        [ExcelArgument(Name = "transform", Description = "Optional: orientation (flip, flipy, rot90, rot180, rot270).")] object? transform = null)
        => Img(inchi, background, style, transform);

    [ExcelFunction(Name = "INCHI2SMILES", Category = "SMILES",
        Description = "Converts an IUPAC InChI string or cell reference to canonical SMILES.", IsMacroType = false)]
    public static object InchiSmiles(
        [ExcelArgument(Name = "inchi", Description = "An IUPAC InChI string or cell reference.")] string inchi)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "INCHI2SMILES";
        if (string.IsNullOrWhiteSpace(inchi)) return ExcelError.ExcelErrorValue;
        try
        {
            var smiles = MoleculeRenderer.TryConvertToSmiles(inchi);
            if (!string.IsNullOrEmpty(smiles))
                return smiles!;
            return ExcelError.ExcelErrorNA;
        }
        catch
        {
            return ExcelError.ExcelErrorValue;
        }
    }

    [ExcelFunction(Name = "INCHI2INCHIKEY", Category = "SMILES",
        Description = "Calculates hashed 27-character IUPAC InChIKey from an InChI string or cell reference.", IsMacroType = false)]
    public static object InchiKey(
        [ExcelArgument(Name = "inchi", Description = "An IUPAC InChI string or cell reference.")] string inchi)
        => SmilesInchiKey(inchi);

    [ExcelFunction(Name = "INCHI2NAME", Category = "SMILES",
        Description = "Resolves chemical substance name from PubChem for an IUPAC InChI string.", IsMacroType = false)]
    public static object InchiName(
        [ExcelArgument(Name = "inchi", Description = "An IUPAC InChI string or cell reference.")] string inchi)
        => SmilesName(inchi);

    [ExcelFunction(Name = "INCHI2MW", Category = "SMILES",
        Description = "Calculates molecular weight (MW) in g/mol for an IUPAC InChI string.", IsMacroType = false)]
    public static object InchiMW(
        [ExcelArgument(Name = "inchi", Description = "An IUPAC InChI string or cell reference.")] string inchi)
        => SmilesMW(inchi);

    [ExcelFunction(Name = "INCHI2FORMULA", Category = "SMILES",
        Description = "Calculates Hill chemical formula for an IUPAC InChI string.", IsMacroType = false)]
    public static object InchiFormula(
        [ExcelArgument(Name = "inchi", Description = "An IUPAC InChI string or cell reference.")] string inchi)
        => SmilesFormula(inchi);

    [ExcelFunction(Name = "SMILES2EXACTMASS", Category = "SMILES",
        Description = "Calculates monoisotopic exact mass (4 decimal places) for a SMILES, InChI, or KET structure.", IsMacroType = false)]
    public static object SmilesExactMass(
        [ExcelArgument(Name = "smiles", Description = "A SMILES string, InChI, KET string, or cell reference.")] string smiles)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "SMILES2EXACTMASS";
        if (string.IsNullOrWhiteSpace(smiles)) return ExcelError.ExcelErrorValue;
        try
        {
            var info = MoleculeRenderer.TryGetMoleculeInfo(smiles, resolvePubChem: false);
            if (info != null && !info.IsReaction && info.ExactMass > 0)
                return Math.Round(info.ExactMass, 4);
            return ExcelError.ExcelErrorNA;
        }
        catch
        {
            return ExcelError.ExcelErrorValue;
        }
    }

    [ExcelFunction(Name = "INCHI2EXACTMASS", Category = "SMILES",
        Description = "Calculates monoisotopic exact mass (4 decimal places) for an IUPAC InChI string.", IsMacroType = false)]
    public static object InchiExactMass(
        [ExcelArgument(Name = "inchi", Description = "An IUPAC InChI string or cell reference.")] string inchi)
        => SmilesExactMass(inchi);

    [ExcelFunction(Name = "KET2EXACTMASS", Category = "SMILES",
        Description = "Calculates monoisotopic exact mass (4 decimal places) for a Ketcher (KET) structure.", IsMacroType = false)]
    public static object KetExactMass(
        [ExcelArgument(Name = "ket", Description = "A Ketcher JSON string or cell reference.")] string ket)
        => SmilesExactMass(ket);

    [ExcelFunction(Name = "SMILES2LOGP", Category = "SMILES",
        Description = "Calculates Crippen LogP (octanol-water partition coefficient) for a SMILES, InChI, or KET structure.", IsMacroType = false)]
    public static object SmilesLogP(
        [ExcelArgument(Name = "smiles", Description = "A SMILES string, InChI, KET string, or cell reference.")] string smiles)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "SMILES2LOGP";
        if (string.IsNullOrWhiteSpace(smiles)) return ExcelError.ExcelErrorValue;
        try
        {
            var info = MoleculeRenderer.TryGetMoleculeInfo(smiles, resolvePubChem: false);
            if (info != null && !info.IsReaction)
                return Math.Round(info.LogP, 2);
            return ExcelError.ExcelErrorNA;
        }
        catch
        {
            return ExcelError.ExcelErrorValue;
        }
    }

    [ExcelFunction(Name = "SMILES2TPSA", Category = "SMILES",
        Description = "Calculates Topological Polar Surface Area (TPSA in Å²) for a SMILES, InChI, or KET structure.", IsMacroType = false)]
    public static object SmilesTpsa(
        [ExcelArgument(Name = "smiles", Description = "A SMILES string, InChI, KET string, or cell reference.")] string smiles)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "SMILES2TPSA";
        if (string.IsNullOrWhiteSpace(smiles)) return ExcelError.ExcelErrorValue;
        try
        {
            var info = MoleculeRenderer.TryGetMoleculeInfo(smiles, resolvePubChem: false);
            if (info != null && !info.IsReaction)
                return Math.Round(info.Tpsa, 2);
            return ExcelError.ExcelErrorNA;
        }
        catch
        {
            return ExcelError.ExcelErrorValue;
        }
    }

    [ExcelFunction(Name = "SMILES2HBD", Category = "SMILES",
        Description = "Counts Hydrogen Bond Donors (HBD) for a SMILES, InChI, or KET structure.", IsMacroType = false)]
    public static object SmilesHbd(
        [ExcelArgument(Name = "smiles", Description = "A SMILES string, InChI, KET string, or cell reference.")] string smiles)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "SMILES2HBD";
        if (string.IsNullOrWhiteSpace(smiles)) return ExcelError.ExcelErrorValue;
        try
        {
            var info = MoleculeRenderer.TryGetMoleculeInfo(smiles, resolvePubChem: false);
            if (info != null && !info.IsReaction)
                return info.HBondDonors;
            return ExcelError.ExcelErrorNA;
        }
        catch
        {
            return ExcelError.ExcelErrorValue;
        }
    }

    [ExcelFunction(Name = "SMILES2HBA", Category = "SMILES",
        Description = "Counts Hydrogen Bond Acceptors (HBA) for a SMILES, InChI, or KET structure.", IsMacroType = false)]
    public static object SmilesHba(
        [ExcelArgument(Name = "smiles", Description = "A SMILES string, InChI, KET string, or cell reference.")] string smiles)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "SMILES2HBA";
        if (string.IsNullOrWhiteSpace(smiles)) return ExcelError.ExcelErrorValue;
        try
        {
            var info = MoleculeRenderer.TryGetMoleculeInfo(smiles, resolvePubChem: false);
            if (info != null && !info.IsReaction)
                return info.HBondAcceptors;
            return ExcelError.ExcelErrorNA;
        }
        catch
        {
            return ExcelError.ExcelErrorValue;
        }
    }

    [ExcelFunction(Name = "SMILES2ROTBONDS", Category = "SMILES",
        Description = "Counts rotatable bonds for a SMILES, InChI, or KET structure.", IsMacroType = false)]
    public static object SmilesRotBonds(
        [ExcelArgument(Name = "smiles", Description = "A SMILES string, InChI, KET string, or cell reference.")] string smiles)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "SMILES2ROTBONDS";
        if (string.IsNullOrWhiteSpace(smiles)) return ExcelError.ExcelErrorValue;
        try
        {
            var info = MoleculeRenderer.TryGetMoleculeInfo(smiles, resolvePubChem: false);
            if (info != null && !info.IsReaction)
                return info.RotatableBonds;
            return ExcelError.ExcelErrorNA;
        }
        catch
        {
            return ExcelError.ExcelErrorValue;
        }
    }

    [ExcelFunction(Name = "SMILES2LIPINSKI", Category = "SMILES",
        Description = "Evaluates Lipinski's Rule of 5 (MW <= 500, LogP <= 5, HBD <= 5, HBA <= 10) and returns \"PASS\" or \"FAIL\".", IsMacroType = false)]
    public static object SmilesLipinski(
        [ExcelArgument(Name = "smiles", Description = "A SMILES string, InChI, KET string, or cell reference.")] string smiles)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "SMILES2LIPINSKI";
        if (string.IsNullOrWhiteSpace(smiles)) return ExcelError.ExcelErrorValue;
        try
        {
            var info = MoleculeRenderer.TryGetMoleculeInfo(smiles, resolvePubChem: false);
            if (info != null && !info.IsReaction)
                return info.PassesLipinski ? "PASS" : "FAIL";
            return ExcelError.ExcelErrorNA;
        }
        catch
        {
            return ExcelError.ExcelErrorValue;
        }
    }

    [ExcelFunction(Name = "SMILES.MATCH", Category = "SMILES",
        Description = "Checks if a target molecule contains a substructure query (SMILES or SMARTS). Returns TRUE or FALSE.", IsMacroType = false)]
    public static object SmilesMatch(
        [ExcelArgument(Name = "target", Description = "Target SMILES, InChI, or KET molecule.")] string target,
        [ExcelArgument(Name = "query", Description = "Substructure query in SMILES or SMARTS format (e.g. \"c1ccccc1\", \"C(=O)O\").")] string query)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "SMILES.MATCH";
        if (string.IsNullOrWhiteSpace(target) || string.IsNullOrWhiteSpace(query))
            return ExcelError.ExcelErrorValue;
        try
        {
            var matched = MoleculeRenderer.IsSubstructureMatch(target, query);
            if (matched.HasValue)
                return matched.Value;
            return ExcelError.ExcelErrorValue;
        }
        catch
        {
            return ExcelError.ExcelErrorValue;
        }
    }

    [ExcelFunction(Name = "SMILES.SIMILARITY", Category = "SMILES",
        Description = "Calculates chemical fingerprint similarity (0.0 to 1.0) between two molecules using Tanimoto or Tversky metric.", IsMacroType = false)]
    public static object SmilesSimilarity(
        [ExcelArgument(Name = "mol1", Description = "First molecule (SMILES, InChI, or KET).")] string mol1,
        [ExcelArgument(Name = "mol2", Description = "Second molecule (SMILES, InChI, or KET).")] string mol2,
        [ExcelArgument(Name = "metric", Description = "Optional: \"tanimoto\" (default), \"tversky\", or \"euclid-sub\".")] object? metric = null)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "SMILES.SIMILARITY";
        if (string.IsNullOrWhiteSpace(mol1) || string.IsNullOrWhiteSpace(mol2))
            return ExcelError.ExcelErrorValue;
        try
        {
            string metricStr = metric is null or ExcelMissing or ExcelEmpty ? "tanimoto" : (metric.ToString() ?? "tanimoto");
            var sim = MoleculeRenderer.CalculateSimilarity(mol1, mol2, metricStr);
            if (sim.HasValue)
                return sim.Value;
            return ExcelError.ExcelErrorValue;
        }
        catch
        {
            return ExcelError.ExcelErrorValue;
        }
    }

    [ExcelFunction(Name = "NAME2SMILES", Category = "SMILES",
        Description = "Resolves a chemical common name, IUPAC name, or CAS Registry Number (e.g., \"Aspirin\", \"50-78-2\") to canonical SMILES.", IsMacroType = false)]
    public static object NameSmiles(
        [ExcelArgument(Name = "name", Description = "Chemical name or CAS number (e.g. \"Aspirin\", \"Caffeine\", \"50-78-2\").")] string name)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "NAME2SMILES";
        if (string.IsNullOrWhiteSpace(name)) return ExcelError.ExcelErrorValue;
        try
        {
            var smiles = PubChemResolver.ResolveNameToSmiles(name);
            if (!string.IsNullOrWhiteSpace(smiles))
                return smiles!;
            return ExcelError.ExcelErrorNA;
        }
        catch
        {
            return ExcelError.ExcelErrorValue;
        }
    }

    [ExcelFunction(Name = "NAME2IMG", Category = "SMILES",
        Description = "[Excel 2024 / Microsoft 365] Resolves a chemical name or CAS number and renders its molecular structure into the cell.", IsMacroType = true)]
    public static object NameImg(
        [ExcelArgument(Name = "name", Description = "Chemical name or CAS number (e.g. \"Aspirin\", \"50-78-2\").")] string name,
        [ExcelArgument(Name = "background", Description = "Optional: \"trans\" (default), \"white\", CSS color, or hex (#RRGGBB).")] object? background = null,
        [ExcelArgument(Name = "style", Description = "Optional: \"color\" (default), \"bw\", \"num\", \"stereo\", \"h\", \"all-c\", \"white\", \"hl:<query>\".")] object? style = null,
        [ExcelArgument(Name = "transform", Description = "Optional: orientation (flip, flipy, rot90, rot180, rot270).")] object? transform = null)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "NAME2IMG";
        if (string.IsNullOrWhiteSpace(name)) return ExcelError.ExcelErrorValue;
        var smiles = PubChemResolver.ResolveNameToSmiles(name);
        if (string.IsNullOrWhiteSpace(smiles)) return ExcelError.ExcelErrorNA;
        return Img(smiles!, background, style, transform);
    }

    /// <summary>
    /// Delegates parsing of Excel formula arguments to RenderOptions.
    /// </summary>
    internal static RenderOptions ParseOptions(object? bgArg, object? styleArg, object? transformArg)
        => RenderOptions.Parse(bgArg, styleArg, transformArg);

    [ExcelCommand(MenuName = "SMILES", MenuText = "Show diagnostics")]
    public static void ShowDiagnostics() => LogDisplay.Show();
}
