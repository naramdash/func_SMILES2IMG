using System;
using System.Runtime.CompilerServices;
using ExcelDna.Integration;
using ExcelDna.Logging;
using Smiles2Img.Rendering;

[assembly: InternalsVisibleTo("Smiles2Img.Tests")]

namespace Smiles2Img;

public static class Functions
{
    internal const string ArgSmilesImgDesc = "SMILES, reaction (>>), InChI, or KET JSON string/cell (e.g., A2, \"CCO\"). #VALUE! on invalid syntax.";
    internal const string ArgKetImgDesc = "Strict Ketcher KET JSON string/cell (molecule or reaction). Returns #VALUE! if input is not KET JSON.";
    internal const string ArgInchiImgDesc = "Strict IUPAC InChI string/cell (\"InChI=1S/...\"). Returns #VALUE! if input is not InChI.";
    internal const string ArgNameDesc = "Compound name, IUPAC name, or CAS number (e.g., \"Aspirin\", \"50-78-2\"). #N/A if unknown or offline.";
    internal const string ArgBgDesc = "Optional background: \"trans\" (default), \"white\", CSS color (\"yellow\"), or hex (\"#1E293B\").";
    internal const string ArgStyleDesc = "Optional tokens (space/pipe): \"color\" (default), \"bw\", \"white\" (dark ink), \"ink=#HEX\", \"num\", \"stereo\", \"h\", \"all-c\", \"hl:<SMARTS>\".";
    internal const string ArgTransformDesc = "Optional planar clockwise rotation: \"rot90\", \"rot180\", or \"rot270\" (preserves R/S stereochemistry).";
    internal const string ArgSmilesScalarDesc = "SMILES string/cell (auto-detects InChI or single-molecule KET). #N/A on queries (*), reactions (>>), or invalid SMILES.";
    internal const string ArgInchiScalarDesc = "Strict IUPAC InChI string/cell (\"InChI=1S/...\"). #VALUE! if not InChI; #N/A if unparseable.";
    internal const string ArgKetScalarDesc = "Strict single-molecule KET JSON string/cell. #VALUE! if not KET; #N/A on reactions or queries.";
    internal const string ArgTpsaIncludeSpDesc = "Optional: FALSE (default, Ertl N/O only = RDKit) or TRUE/\"sp\"/\"pubchem\" (include S & P = PubChem).";
    internal const string ArgSmartsQueryDesc = "Substructure query in SMILES or SMARTS (e.g., \"c1ccccc1\", \"[CX3](=O)[OX2H1]\"). #VALUE! on bad syntax.";

    [ExcelFunction(Name = "SMILES2IMG", Category = "SMILES",
        Description = "[Excel 365/2024+] In-cell 2D structure/reaction image (offline Indigo). Accepts SMILES, >>, InChI, or KET. Use .FLOAT on Excel 2016-2021; #VALUE! on bad syntax.", IsMacroType = true)]
    public static object Img(
        [ExcelArgument(Name = "smiles", Description = ArgSmilesImgDesc)] string smiles,
        [ExcelArgument(Name = "background", Description = ArgBgDesc)] object? background = null,
        [ExcelArgument(Name = "style", Description = ArgStyleDesc)] object? style = null,
        [ExcelArgument(Name = "transform", Description = ArgTransformDesc)] object? transform = null)
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
        Description = "[Excel 365/2024+] In-cell 2D image from strict Ketcher KET JSON (molecule or reaction, offline). Use .FLOAT on Excel 2016-2021; #VALUE! if not KET.", IsMacroType = true)]
    public static object KetImg(
        [ExcelArgument(Name = "ket", Description = ArgKetImgDesc)] string ket,
        [ExcelArgument(Name = "background", Description = ArgBgDesc)] object? background = null,
        [ExcelArgument(Name = "style", Description = ArgStyleDesc)] object? style = null,
        [ExcelArgument(Name = "transform", Description = ArgTransformDesc)] object? transform = null)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "KET2IMG";
        if (string.IsNullOrWhiteSpace(ket) || !MoleculeRenderer.IsKet(ket)) return ExcelError.ExcelErrorValue;
        if (XlCall.Excel(XlCall.xlfCaller) is not ExcelReference caller ||
            caller.RowFirst != caller.RowLast || caller.ColumnFirst != caller.ColumnLast)
            return ExcelError.ExcelErrorValue;

        if (CellImages.NativeImagesUnsupported) return CellImages.UnsupportedMessage;

        try
        {
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

    private static object EvaluateMoleculeProperty(
        string smiles,
        string wizardLabel,
        Func<MoleculeInfo, object?> selector,
        bool resolvePubChem = false,
        bool allowReaction = false)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return wizardLabel;
        if (string.IsNullOrWhiteSpace(smiles)) return ExcelError.ExcelErrorValue;
        try
        {
            var info = MoleculeRenderer.TryGetMoleculeInfo(smiles, resolvePubChem);
            if (info != null && !info.HasQueryOrDummyAtoms && (allowReaction || !info.IsReaction))
            {
                var val = selector(info);
                if (val != null) return val;
            }
            return ExcelError.ExcelErrorNA;
        }
        catch
        {
            return ExcelError.ExcelErrorValue;
        }
    }

    [ExcelFunction(Name = "SMILES2NAME", Category = "SMILES",
        Description = "Substance name (built-in common compounds offline; else PubChem API with 30d cache). Returns #N/A if offline, unknown, query, or reaction.", IsMacroType = false)]
    public static object SmilesName(
        [ExcelArgument(Name = "smiles", Description = ArgSmilesScalarDesc)] string smiles)
        => EvaluateMoleculeProperty(smiles, "SMILES2NAME",
            info => !string.IsNullOrWhiteSpace(info.SubstanceName) ? info.SubstanceName : null,
            resolvePubChem: true, allowReaction: false);

    [ExcelFunction(Name = "SMILES2MW", Category = "SMILES",
        Description = "Average molecular weight in g/mol (number, 2 dec, offline Indigo). Isotopes & salts/counterions included as drawn. #N/A on queries/reactions.", IsMacroType = false)]
    public static object SmilesMW(
        [ExcelArgument(Name = "smiles", Description = ArgSmilesScalarDesc)] string smiles)
        => EvaluateMoleculeProperty(smiles, "SMILES2MW",
            info => info.MolecularWeight > 0 ? Math.Round(info.MolecularWeight, 2, MidpointRounding.AwayFromZero) : (object?)null);

    [ExcelFunction(Name = "SMILES2FORMULA", Category = "SMILES",
        Description = "Hill chemical formula (text, offline Indigo). Salts/counterions included as drawn; isotopes folded to element symbol. #N/A on queries/reactions.", IsMacroType = false)]
    public static object SmilesFormula(
        [ExcelArgument(Name = "smiles", Description = ArgSmilesScalarDesc)] string smiles)
        => EvaluateMoleculeProperty(smiles, "SMILES2FORMULA",
            info => !string.IsNullOrWhiteSpace(info.Formula) ? info.Formula : null);

    [ExcelFunction(Name = "KET2NAME", Category = "SMILES",
        Description = "Substance name for strict single-molecule KET (common names offline; else PubChem API, 30d cache). #VALUE! if not KET; #N/A on reactions.", IsMacroType = false)]
    public static object KetName(
        [ExcelArgument(Name = "ket", Description = ArgKetScalarDesc)] string ket)
        => EvaluateKetFunction(ket, "KET2NAME", SmilesName);

    [ExcelFunction(Name = "KET2MW", Category = "SMILES",
        Description = "Average molecular weight in g/mol (number, 2 dec, offline Indigo) for strict single-molecule KET. #VALUE! if not KET; #N/A on reactions.", IsMacroType = false)]
    public static object KetMW(
        [ExcelArgument(Name = "ket", Description = ArgKetScalarDesc)] string ket)
        => EvaluateKetFunction(ket, "KET2MW", SmilesMW);

    [ExcelFunction(Name = "KET2FORMULA", Category = "SMILES",
        Description = "Hill chemical formula (text, offline Indigo) for strict single-molecule KET JSON. #VALUE! if not KET; #N/A on reactions.", IsMacroType = false)]
    public static object KetFormula(
        [ExcelArgument(Name = "ket", Description = ArgKetScalarDesc)] string ket)
        => EvaluateKetFunction(ket, "KET2FORMULA", SmilesFormula);

    [ExcelFunction(Name = "SMILES2INCHI", Category = "SMILES",
        Description = "Standard IUPAC InChI (1S text, offline). Auto-detects SMILES/InChI/KET; #N/A on queries/reactions or invalid structures.", IsMacroType = false)]
    public static object SmilesInchi(
        [ExcelArgument(Name = "smiles", Description = ArgSmilesScalarDesc)] string smiles)
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
        Description = "27-char Standard InChIKey (text, offline) for deduplication (COUNTIF/XLOOKUP). Auto-detects SMILES/InChI/KET; #N/A on queries/reactions.", IsMacroType = false)]
    public static object SmilesInchiKey(
        [ExcelArgument(Name = "smiles", Description = ArgSmilesScalarDesc)] string smiles)
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
        Description = "Standard IUPAC InChI (1S text, offline) for strict single-molecule KET JSON. #VALUE! if not KET; #N/A on reactions.", IsMacroType = false)]
    public static object KetInchi(
        [ExcelArgument(Name = "ket", Description = ArgKetScalarDesc)] string ket)
        => EvaluateKetFunction(ket, "KET2INCHI", SmilesInchi);

    [ExcelFunction(Name = "KET2INCHIKEY", Category = "SMILES",
        Description = "27-char Standard InChIKey (text, offline) from strict single-molecule KET JSON. #VALUE! if not KET; #N/A on reactions.", IsMacroType = false)]
    public static object KetInchiKey(
        [ExcelArgument(Name = "ket", Description = ArgKetScalarDesc)] string ket)
        => EvaluateKetFunction(ket, "KET2INCHIKEY", SmilesInchiKey);

    private static object EvaluateInchiFunction(string inchi, string wizardLabel, Func<string, object> evaluator)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return wizardLabel;
        if (string.IsNullOrWhiteSpace(inchi) || !MoleculeRenderer.IsInchi(inchi)) return ExcelError.ExcelErrorValue;
        return evaluator(inchi);
    }

    private static object EvaluateKetFunction(string ket, string wizardLabel, Func<string, object> evaluator)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return wizardLabel;
        if (string.IsNullOrWhiteSpace(ket) || !MoleculeRenderer.IsKet(ket)) return ExcelError.ExcelErrorValue;
        return evaluator(ket);
    }

    [ExcelFunction(Name = "INCHI2IMG", Category = "SMILES",
        Description = "[Excel 365/2024+] In-cell 2D image from strict IUPAC InChI (offline Indigo). Use .FLOAT on Excel 2016-2021; #VALUE! if input is not InChI.", IsMacroType = true)]
    public static object InchiImg(
        [ExcelArgument(Name = "inchi", Description = ArgInchiImgDesc)] string inchi,
        [ExcelArgument(Name = "background", Description = ArgBgDesc)] object? background = null,
        [ExcelArgument(Name = "style", Description = ArgStyleDesc)] object? style = null,
        [ExcelArgument(Name = "transform", Description = ArgTransformDesc)] object? transform = null)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "INCHI2IMG";
        if (string.IsNullOrWhiteSpace(inchi) || !MoleculeRenderer.IsInchi(inchi)) return ExcelError.ExcelErrorValue;
        return Img(inchi, background, style, transform);
    }

    [ExcelFunction(Name = "INCHI2SMILES", Category = "SMILES",
        Description = "Canonical SMILES (text, offline) from strict IUPAC InChI. Note: InChI=1S mobile-H may pick alternate amide/enol tautomer; #VALUE! if not InChI.", IsMacroType = false)]
    public static object InchiSmiles(
        [ExcelArgument(Name = "inchi", Description = ArgInchiScalarDesc)] string inchi)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "INCHI2SMILES";
        if (string.IsNullOrWhiteSpace(inchi) || !MoleculeRenderer.IsInchi(inchi)) return ExcelError.ExcelErrorValue;
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
        Description = "27-char Standard InChIKey (text, offline) from strict IUPAC InChI for deduplication. #VALUE! if not InChI.", IsMacroType = false)]
    public static object InchiKey(
        [ExcelArgument(Name = "inchi", Description = ArgInchiScalarDesc)] string inchi)
        => EvaluateInchiFunction(inchi, "INCHI2INCHIKEY", SmilesInchiKey);

    [ExcelFunction(Name = "INCHI2NAME", Category = "SMILES",
        Description = "Substance name for strict IUPAC InChI (common names offline; else PubChem API, 30d cache). #VALUE! if not InChI; #N/A if offline/unknown.", IsMacroType = false)]
    public static object InchiName(
        [ExcelArgument(Name = "inchi", Description = ArgInchiScalarDesc)] string inchi)
        => EvaluateInchiFunction(inchi, "INCHI2NAME", SmilesName);

    [ExcelFunction(Name = "INCHI2MW", Category = "SMILES",
        Description = "Average molecular weight in g/mol (number, 2 dec, offline Indigo) for strict InChI. Isotopes & salts included as drawn; #VALUE! if not InChI.", IsMacroType = false)]
    public static object InchiMW(
        [ExcelArgument(Name = "inchi", Description = ArgInchiScalarDesc)] string inchi)
        => EvaluateInchiFunction(inchi, "INCHI2MW", SmilesMW);

    [ExcelFunction(Name = "INCHI2FORMULA", Category = "SMILES",
        Description = "Hill chemical formula (text, offline Indigo) for strict IUPAC InChI. Salts included as drawn; #VALUE! if not InChI.", IsMacroType = false)]
    public static object InchiFormula(
        [ExcelArgument(Name = "inchi", Description = ArgInchiScalarDesc)] string inchi)
        => EvaluateInchiFunction(inchi, "INCHI2FORMULA", SmilesFormula);

    [ExcelFunction(Name = "SMILES2EXACTMASS", Category = "SMILES",
        Description = "Neutral monoisotopic exact mass M in Da (4 dec, offline Indigo; add/sub 1.0073 for [M+H]+/[M-H]-). Isotopes & salts included as drawn.", IsMacroType = false)]
    public static object SmilesExactMass(
        [ExcelArgument(Name = "smiles", Description = ArgSmilesScalarDesc)] string smiles)
        => EvaluateMoleculeProperty(smiles, "SMILES2EXACTMASS",
            info => info.ExactMass > 0 ? Math.Round(info.ExactMass, 4, MidpointRounding.AwayFromZero) : (object?)null);

    [ExcelFunction(Name = "INCHI2EXACTMASS", Category = "SMILES",
        Description = "Neutral monoisotopic exact mass M in Da (4 dec, offline Indigo) for strict InChI. Isotopes & salts included as drawn; #VALUE! if not InChI.", IsMacroType = false)]
    public static object InchiExactMass(
        [ExcelArgument(Name = "inchi", Description = ArgInchiScalarDesc)] string inchi)
        => EvaluateInchiFunction(inchi, "INCHI2EXACTMASS", SmilesExactMass);

    [ExcelFunction(Name = "KET2EXACTMASS", Category = "SMILES",
        Description = "Neutral monoisotopic exact mass M in Da (4 dec, offline Indigo) for strict single-molecule KET. #VALUE! if not KET; #N/A on reactions.", IsMacroType = false)]
    public static object KetExactMass(
        [ExcelArgument(Name = "ket", Description = ArgKetScalarDesc)] string ket)
        => EvaluateKetFunction(ket, "KET2EXACTMASS", SmilesExactMass);

    [ExcelFunction(Name = "KET2SMILES", Category = "SMILES",
        Description = "Canonical SMILES (text, offline) from strict single-molecule KET JSON. Preserves drawn tautomer/stereo; #VALUE! if not KET; #N/A on reactions.", IsMacroType = false)]
    public static object KetSmiles(
        [ExcelArgument(Name = "ket", Description = ArgKetScalarDesc)] string ket)
        => EvaluateKetFunction(ket, "KET2SMILES", k =>
        {
            try
            {
                var smiles = MoleculeRenderer.TryConvertToSmiles(k);
                return !string.IsNullOrEmpty(smiles) ? smiles! : ExcelError.ExcelErrorNA;
            }
            catch
            {
                return ExcelError.ExcelErrorValue;
            }
        });

    [ExcelFunction(Name = "SMILES2LOGP", Category = "SMILES",
        Description = "Predicted LogP (Wildman-Crippen, 2 dec, offline Indigo). Model only (not experimental/XLogP3); sensitive to charge/salts (best for neutral organics).", IsMacroType = false)]
    public static object SmilesLogP(
        [ExcelArgument(Name = "smiles", Description = ArgSmilesScalarDesc)] string smiles)
        => EvaluateMoleculeProperty(smiles, "SMILES2LOGP", info => info.HasLogP ? Math.Round(info.LogP, 2, MidpointRounding.AwayFromZero) : (object?)null);

    [ExcelFunction(Name = "INCHI2LOGP", Category = "SMILES",
        Description = "Predicted LogP (Wildman-Crippen, 2 dec, offline) for strict InChI (not experimental/XLogP3). InChI=1S mobile-H may shift amides; #VALUE! if not InChI.", IsMacroType = false)]
    public static object InchiLogP(
        [ExcelArgument(Name = "inchi", Description = ArgInchiScalarDesc)] string inchi)
        => EvaluateInchiFunction(inchi, "INCHI2LOGP", SmilesLogP);

    [ExcelFunction(Name = "KET2LOGP", Category = "SMILES",
        Description = "Predicted LogP (Wildman-Crippen, 2 dec, offline) for strict single-molecule KET (not experimental/XLogP3). #VALUE! if not KET; #N/A on reactions.", IsMacroType = false)]
    public static object KetLogP(
        [ExcelArgument(Name = "ket", Description = ArgKetScalarDesc)] string ket)
        => EvaluateKetFunction(ket, "KET2LOGP", SmilesLogP);

    private static bool ParseIncludeSp(object? arg)
    {
        if (arg is null or ExcelMissing or ExcelEmpty) return false;
        if (arg is bool b) return b;
        if (arg is double d) return d != 0;
        if (arg is string s)
        {
            var norm = s.Trim().ToLowerInvariant();
            return norm is "true" or "1" or "yes" or "sp" or "s" or "p" or "pubchem" or "all";
        }
        return false;
    }

    [ExcelFunction(Name = "SMILES2TPSA", Category = "SMILES",
        Description = "Ertl TPSA in Å² (2 dec, offline Indigo). Default FALSE=N/O only (RDKit); TRUE=+S/P (PubChem). Sensitive to formal charges; salts included as drawn.", IsMacroType = false)]
    public static object SmilesTpsa(
        [ExcelArgument(Name = "smiles", Description = ArgSmilesScalarDesc)] string smiles,
        [ExcelArgument(Name = "include_sp", Description = ArgTpsaIncludeSpDesc)] object? includeSp = null)
    {
        var sp = ParseIncludeSp(includeSp);
        return EvaluateMoleculeProperty(smiles, "SMILES2TPSA",
            info => sp
                ? (info.HasTpsaWithSP ? Math.Round(info.TpsaWithSP, 2, MidpointRounding.AwayFromZero) : (object?)null)
                : (info.HasTpsa ? Math.Round(info.Tpsa, 2, MidpointRounding.AwayFromZero) : (object?)null));
    }

    [ExcelFunction(Name = "INCHI2TPSA", Category = "SMILES",
        Description = "Ertl TPSA in Å² (2 dec, offline) for strict InChI. FALSE=N/O (RDKit), TRUE=+S/P (PubChem). InChI=1S mobile-H may shift amides; #VALUE! if not InChI.", IsMacroType = false)]
    public static object InchiTpsa(
        [ExcelArgument(Name = "inchi", Description = ArgInchiScalarDesc)] string inchi,
        [ExcelArgument(Name = "include_sp", Description = ArgTpsaIncludeSpDesc)] object? includeSp = null)
        => EvaluateInchiFunction(inchi, "INCHI2TPSA", i => SmilesTpsa(i, includeSp));

    [ExcelFunction(Name = "KET2TPSA", Category = "SMILES",
        Description = "Ertl TPSA in Å² (2 dec, offline) for strict single-molecule KET. FALSE=N/O (RDKit), TRUE=+S/P (PubChem). #VALUE! if not KET; #N/A on reactions.", IsMacroType = false)]
    public static object KetTpsa(
        [ExcelArgument(Name = "ket", Description = ArgKetScalarDesc)] string ket,
        [ExcelArgument(Name = "include_sp", Description = ArgTpsaIncludeSpDesc)] object? includeSp = null)
        => EvaluateKetFunction(ket, "KET2TPSA", k => SmilesTpsa(k, includeSp));

    [ExcelFunction(Name = "SMILES2HEAVY", Category = "SMILES",
        Description = "Heavy (non-hydrogen) atom count (integer, offline Indigo). Evaluates salts/counterions as drawn; #N/A on queries/reactions.", IsMacroType = false)]
    public static object SmilesHeavy(
        [ExcelArgument(Name = "smiles", Description = ArgSmilesScalarDesc)] string smiles)
        => EvaluateMoleculeProperty(smiles, "SMILES2HEAVY", info => info.HeavyAtomCount);

    [ExcelFunction(Name = "INCHI2HEAVY", Category = "SMILES",
        Description = "Heavy (non-hydrogen) atom count (integer, offline Indigo) for strict IUPAC InChI. Evaluates salts as drawn; #VALUE! if not InChI.", IsMacroType = false)]
    public static object InchiHeavy(
        [ExcelArgument(Name = "inchi", Description = ArgInchiScalarDesc)] string inchi)
        => EvaluateInchiFunction(inchi, "INCHI2HEAVY", SmilesHeavy);

    [ExcelFunction(Name = "KET2HEAVY", Category = "SMILES",
        Description = "Heavy (non-hydrogen) atom count (integer, offline Indigo) for strict single-molecule KET JSON. #VALUE! if not KET; #N/A on reactions.", IsMacroType = false)]
    public static object KetHeavy(
        [ExcelArgument(Name = "ket", Description = ArgKetScalarDesc)] string ket)
        => EvaluateKetFunction(ket, "KET2HEAVY", SmilesHeavy);

    [ExcelFunction(Name = "SMILES2RINGS", Category = "SMILES",
        Description = "Total ring count (SSSR integer, offline Indigo). Evaluates all fragments/salts as drawn; #N/A on queries/reactions.", IsMacroType = false)]
    public static object SmilesRings(
        [ExcelArgument(Name = "smiles", Description = ArgSmilesScalarDesc)] string smiles)
        => EvaluateMoleculeProperty(smiles, "SMILES2RINGS", info => info.RingCount);

    [ExcelFunction(Name = "INCHI2RINGS", Category = "SMILES",
        Description = "Total ring count (SSSR integer, offline Indigo) for strict IUPAC InChI. #VALUE! if not InChI; #N/A if unparseable.", IsMacroType = false)]
    public static object InchiRings(
        [ExcelArgument(Name = "inchi", Description = ArgInchiScalarDesc)] string inchi)
        => EvaluateInchiFunction(inchi, "INCHI2RINGS", SmilesRings);

    [ExcelFunction(Name = "KET2RINGS", Category = "SMILES",
        Description = "Total ring count (SSSR integer, offline Indigo) for strict single-molecule KET JSON. #VALUE! if not KET; #N/A on reactions.", IsMacroType = false)]
    public static object KetRings(
        [ExcelArgument(Name = "ket", Description = ArgKetScalarDesc)] string ket)
        => EvaluateKetFunction(ket, "KET2RINGS", SmilesRings);

    [ExcelFunction(Name = "SMILES2AROMATIC", Category = "SMILES",
        Description = "Aromatic ring count (SSSR integer, offline) via Indigo generic aromaticity model. May differ from RDKit/Daylight on non-benzenoids (e.g. Azulene).", IsMacroType = false)]
    public static object SmilesAromatic(
        [ExcelArgument(Name = "smiles", Description = ArgSmilesScalarDesc)] string smiles)
        => EvaluateMoleculeProperty(smiles, "SMILES2AROMATIC", info => info.AromaticRingCount);

    [ExcelFunction(Name = "INCHI2AROMATIC", Category = "SMILES",
        Description = "Aromatic ring count (SSSR integer, offline; Indigo generic model) for strict IUPAC InChI. #VALUE! if not InChI.", IsMacroType = false)]
    public static object InchiAromatic(
        [ExcelArgument(Name = "inchi", Description = ArgInchiScalarDesc)] string inchi)
        => EvaluateInchiFunction(inchi, "INCHI2AROMATIC", SmilesAromatic);

    [ExcelFunction(Name = "KET2AROMATIC", Category = "SMILES",
        Description = "Aromatic ring count (SSSR integer, offline; Indigo generic model) for strict single-molecule KET. #VALUE! if not KET; #N/A on reactions.", IsMacroType = false)]
    public static object KetAromatic(
        [ExcelArgument(Name = "ket", Description = ArgKetScalarDesc)] string ket)
        => EvaluateKetFunction(ket, "KET2AROMATIC", SmilesAromatic);

    [ExcelFunction(Name = "SMILES2FSP3", Category = "SMILES",
        Description = "Fraction of sp3 carbons (sp3 C / total C, 0.00-1.00, 2 dec, offline Indigo). Returns 0.00 for carbon-free molecules (e.g. H2O, NaCl).", IsMacroType = false)]
    public static object SmilesFsp3(
        [ExcelArgument(Name = "smiles", Description = ArgSmilesScalarDesc)] string smiles)
        => EvaluateMoleculeProperty(smiles, "SMILES2FSP3", info => Math.Round(info.Fsp3, 2, MidpointRounding.AwayFromZero));

    [ExcelFunction(Name = "INCHI2FSP3", Category = "SMILES",
        Description = "Fraction of sp3 carbons (sp3 C / total C, 0.00-1.00, 2 dec, offline) for strict InChI. Returns 0.00 if carbon-free; #VALUE! if not InChI.", IsMacroType = false)]
    public static object InchiFsp3(
        [ExcelArgument(Name = "inchi", Description = ArgInchiScalarDesc)] string inchi)
        => EvaluateInchiFunction(inchi, "INCHI2FSP3", SmilesFsp3);

    [ExcelFunction(Name = "KET2FSP3", Category = "SMILES",
        Description = "Fraction of sp3 carbons (sp3 C / total C, 0.00-1.00, 2 dec, offline) for strict single-molecule KET. Returns 0.00 if carbon-free; #VALUE! if not KET.", IsMacroType = false)]
    public static object KetFsp3(
        [ExcelArgument(Name = "ket", Description = ArgKetScalarDesc)] string ket)
        => EvaluateKetFunction(ket, "KET2FSP3", SmilesFsp3);

    [ExcelFunction(Name = "SMILES.MATCH", Category = "SMILES",
        Description = "Offline SMARTS/SMILES substructure match (Indigo) -> TRUE/FALSE. Auto-detects SMILES/InChI/KET target; #VALUE! on bad syntax or reaction.", IsMacroType = false)]
    public static object SmilesMatch(
        [ExcelArgument(Name = "target", Description = "Target SMILES, InChI, or single-molecule KET string/cell. Returns #VALUE! on invalid syntax or reaction.")] string target,
        [ExcelArgument(Name = "query", Description = ArgSmartsQueryDesc)] string query)
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

    [ExcelFunction(Name = "INCHI.MATCH", Category = "SMILES",
        Description = "Offline SMARTS/SMILES substructure match -> TRUE/FALSE for strict IUPAC InChI target. #VALUE! if not InChI or bad SMARTS syntax.", IsMacroType = false)]
    public static object InchiMatch(
        [ExcelArgument(Name = "inchi", Description = "Strict IUPAC InChI target string/cell (\"InChI=1S/...\"). Returns #VALUE! if input is not InChI.")] string inchi,
        [ExcelArgument(Name = "query", Description = ArgSmartsQueryDesc)] string query)
        => EvaluateInchiFunction(inchi, "INCHI.MATCH", i => SmilesMatch(i, query));

    [ExcelFunction(Name = "KET.MATCH", Category = "SMILES",
        Description = "Offline SMARTS/SMILES substructure match -> TRUE/FALSE for strict single-molecule KET target. #VALUE! if not KET, reaction, or bad SMARTS.", IsMacroType = false)]
    public static object KetMatch(
        [ExcelArgument(Name = "ket", Description = "Strict single-molecule KET JSON target string/cell. Returns #VALUE! if not KET or if reaction.")] string ket,
        [ExcelArgument(Name = "query", Description = ArgSmartsQueryDesc)] string query)
        => EvaluateKetFunction(ket, "KET.MATCH", k => SmilesMatch(k, query));

    [ExcelFunction(Name = "NAME2SMILES", Category = "SMILES",
        Description = "Resolves name/CAS to canonical SMILES. 104 built-in offline + CAS checksum (0ms); else PubChem API (30d local cache, #N/A if offline/unknown).", IsMacroType = false)]
    public static object NameSmiles(
        [ExcelArgument(Name = "name", Description = ArgNameDesc)] string name)
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
        Description = "[Excel 365/2024+] In-cell 2D image from name/CAS. 104 built-in offline + CAS check; else PubChem (30d cache, #N/A if offline). Use .FLOAT on Excel 2016-2021.", IsMacroType = true)]
    public static object NameImg(
        [ExcelArgument(Name = "name", Description = ArgNameDesc)] string name,
        [ExcelArgument(Name = "background", Description = ArgBgDesc)] object? background = null,
        [ExcelArgument(Name = "style", Description = ArgStyleDesc)] object? style = null,
        [ExcelArgument(Name = "transform", Description = ArgTransformDesc)] object? transform = null)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return "NAME2IMG";
        if (XlCall.Excel(XlCall.xlfCaller) is not ExcelReference caller ||
            caller.RowFirst != caller.RowLast || caller.ColumnFirst != caller.ColumnLast)
            return ExcelError.ExcelErrorValue;
        if (CellImages.NativeImagesUnsupported) return CellImages.UnsupportedMessage;
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
