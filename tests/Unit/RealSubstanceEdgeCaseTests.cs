using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using Smiles2Img;
using Smiles2Img.Editor;
using Smiles2Img.Rendering;
using static Smiles2Img.Tests.TestAssert;

namespace Smiles2Img.Tests.Unit;

internal static class RealSubstanceEdgeCaseTests
{
    private sealed class SubstanceCase
    {
        public string Name { get; }
        public string Smiles { get; }
        public string Formula { get; }
        public double Mw { get; }
        public double ExactMass { get; }
        public double LogP { get; }
        public double TpsaNoSp { get; }
        public double TpsaWithSp { get; }
        public int Heavy { get; }
        public int Rings { get; }
        public int Aromatic { get; }
        public double Fsp3 { get; }
        public string? InchiKey { get; }

        public SubstanceCase(
            string name,
            string smiles,
            string formula,
            double mw,
            double exactMass,
            double logP,
            double tpsaNoSp,
            double tpsaWithSp,
            int heavy,
            int rings,
            int aromatic,
            double fsp3,
            string? inchiKey = null)
        {
            Name = name;
            Smiles = smiles;
            Formula = formula;
            Mw = mw;
            ExactMass = exactMass;
            LogP = logP;
            TpsaNoSp = tpsaNoSp;
            TpsaWithSp = tpsaWithSp;
            Heavy = heavy;
            Rings = rings;
            Aromatic = aromatic;
            Fsp3 = fsp3;
            InchiKey = inchiKey;
        }
    }

    private static readonly SubstanceCase[] BenchmarkMatrix =
    [
        // FDA drugs & natural products
        new("Aspirin", "CC(=O)Oc1ccccc1C(=O)O", "C9H8O4", 180.16, 180.0423, 1.31, 63.60, 63.60, 13, 1, 1, 0.11, "BSYNRYMUTXBXSQ-UHFFFAOYSA-N"),
        new("Caffeine", "Cn1cnc2c1c(=O)n(C)c(=O)n2C", "C8H10N4O2", 194.19, 194.0804, -1.03, 61.82, 61.82, 14, 2, 2, 0.38, "RYYVLZVUVIJVGH-UHFFFAOYSA-N"),
        new("Paracetamol", "CC(=O)Nc1ccc(O)cc1", "C8H9NO2", 151.17, 151.0633, 1.35, 49.33, 49.33, 11, 1, 1, 0.13, "RZVAJINKPMORJF-UHFFFAOYSA-N"),
        new("(S)-Ibuprofen", "CC(C)Cc1ccc(cc1)[C@@H](C)C(=O)O", "C13H18O2", 206.29, 206.1307, 3.21, 37.30, 37.30, 15, 1, 1, 0.46, "HEFNNWSXXWATRW-SNVBAGLBSA-N"),
        new("(R)-Ibuprofen", "CC(C)Cc1ccc([C@H](C)C(=O)O)cc1", "C13H18O2", 206.29, 206.1307, 3.21, 37.30, 37.30, 15, 1, 1, 0.46, "HEFNNWSXXWATRW-JTQLQIEISA-N"),
        new("(S)-Naproxen", "COc1ccc2cc([C@H](C)C(=O)O)ccc2c1", "C14H14O3", 230.26, 230.0943, 3.04, 46.53, 46.53, 17, 2, 2, 0.21, "CMWTZPSULFXXJA-VIFPVBQESA-N"),
        new("Nicotine", "CN1CCC[C@H]1c2cccnc2", "C10H14N2", 162.24, 162.1157, 1.85, 16.13, 16.13, 12, 2, 1, 0.50, "SNICXCGAKADSCV-JTQLQIEISA-N"),
        new("Dopamine", "NCCc1ccc(O)c(O)c1", "C8H11NO2", 153.18, 153.0790, 0.73, 66.48, 66.48, 11, 1, 1, 0.25, "VYFYYTLLBUKUHU-UHFFFAOYSA-N"),
        new("Serotonin", "NCCc1c[nH]c2ccc(O)cc12", "C10H12N2O", 176.22, 176.0950, 1.51, 62.04, 62.04, 13, 2, 2, 0.20, "QZAYGJVTTNCVMB-UHFFFAOYSA-N"),
        new("Epinephrine", "CNC[C@H](O)c1ccc(O)c(O)c1", "C9H13NO3", 183.21, 183.0895, 0.35, 72.72, 72.72, 13, 1, 1, 0.33, "UCTWMZQNUQWSLP-VIFPVBQESA-N"),
        new("Atorvastatin", "CC(C)c1c(C(=O)Nc2ccccc2)c(c(n1CC[C@@H](C[C@@H](CC(=O)O)O)O)c3ccc(F)cc3)c4ccccc4", "C33H35FN2O5", 558.65, 558.2530, 6.45, 111.79, 111.79, 41, 4, 4, 0.27, "XUKUURHRXDUEBC-SVBPBHIXSA-N"),
        new("Penicillin G", "CC1([C@@H](N2[C@H](S1)[C@@H](C2=O)NC(=O)Cc3ccccc3)C(=O)O)C", "C16H18N2O4S", 334.39, 334.0987, 0.99, 86.71, 112.01, 23, 3, 1, 0.44, "JGSARLDLIJGVTE-MBNYWOFBSA-N"),
        new("Amoxicillin", "CC1([C@@H](N2[C@H](S1)[C@@H](C2=O)NC(=O)[C@@H](c3ccc(O)cc3)N)C(=O)O)C", "C16H19N3O5S", 365.40, 365.1045, 0.02, 132.96, 158.26, 25, 3, 1, 0.44, "LSQZJLSUYDQPKJ-NJBDSQKTSA-N"),
        new("Sulfamethoxazole", "Cc1cc(NS(=O)(=O)c2ccc(N)cc2)no1", "C10H11N3O3S", 253.28, 253.0521, 1.37, 98.20, 106.58, 17, 2, 2, 0.10, "JLKIGFTWXXRPMT-UHFFFAOYSA-N"),
        new("Morphine", "CN1CC[C@]23[C@@H]4Oc5c3c(C[C@@H]1[C@@H]2C=C[C@@H]4O)ccc5O", "C17H19NO3", 285.34, 285.1365, 1.33, 52.93, 52.93, 21, 5, 1, 0.53, "BQJCRHHNABKAKU-ORWBNGGBSA-N"),
        new("Artemisinin", "C[C@@H]1CC[C@H]2[C@@H](C)C(=O)O[C@@H]3O[C@@]4(C)CC[C@@H]1[C@]23OO4", "C15H22O5", 282.34, 282.1467, 2.39, 53.99, 53.99, 20, 4, 0, 0.93, "BLUAFEHZUWYNDE-GZEPNZDPSA-N"),
        new("Cholesterol", "C[C@H](CCCC(C)C)[C@H]1CC[C@@H]2[C@@]1(CC[C@H]3[C@H]2CC=C4[C@@]3(CC[C@@H](C4)O)C)C", "C27H46O", 386.66, 386.3549, 7.39, 20.23, 20.23, 28, 4, 0, 0.93, "HVYWMOMLDIMFJA-DPAQBDIFSA-N"),
        new("Paclitaxel", "CC1=C2[C@@]([C@]([C@H]([C@@H]3[C@]4([C@H](OC4)C[C@@H]([C@]3(C(=O)[C@@H]2OC(=O)C)C)O)OC(=O)C)OC(=O)c5ccccc5)(C[C@@H]1OC(=O)[C@H](O)[C@@H](NC(=O)c6ccccc6)c7ccccc7)O)(C)C", "C47H51NO14", 853.92, 853.3310, 3.74, 221.29, 221.29, 62, 7, 3, 0.45, "RCINICONZNJXQF-MZXODVADSA-N"),
        new("Zidovudine (AZT)", "Cc1cn([C@H]2C[C@H](N=[N+]=[N-])[C@@H](CO)O2)c(=O)[nH]c1=O", "C10H13N5O4", 267.25, 267.0968, -0.20, 96.70, 96.70, 19, 2, 1, 0.60),
        new("Bortezomib", "CC(C)C[C@H](NC(=O)[C@H](Cc1ccccc1)NC(=O)c1cnccn1)B(O)O", "C19H25BN4O4", 384.24, 384.1969, 0.49, 124.44, 124.44, 28, 2, 2, 0.37, "GXJABQQUPOEUTA-RDJZCZTQSA-N"),
        new("Cyclosporine A", @"CC[C@@H]1NC(=O)[C@H]([C@H](O)[C@H](C)C\C=C\C)N(C)C(=O)[C@H](C(C)C)N(C)C(=O)[C@H](CC(C)C)N(C)C(=O)[C@H](CC(C)C)N(C)C(=O)[C@H](C)NC(=O)[C@H](C)NC(=O)[C@H](CC(C)C)N(C)C(=O)[C@@H](NC(=O)[C@H](CC(C)C)N(C)C(=O)CN(C)C1=O)C(C)C", "C62H111N11O12", 1202.64, 1201.8414, 3.27, 278.80, 278.80, 85, 1, 0, 0.79, "PMATZTZNYRCHOR-RCAXAPJVSA-N"),

        // Sulfur / Phosphorus TPSA & nucleotide triphosphate
        new("Glyphosate", "OC(=O)CNCP(=O)(O)O", "C3H8NO5P", 169.07, 169.0140, -1.20, 106.86, 116.67, 10, 0, 0, 0.67, "XDDAORKBJWWYJS-UHFFFAOYSA-N"),
        new("ATP", "Nc1ncnc2c1ncn2[C@@H]1O[C@H](COP(=O)(O)OP(=O)(O)OP(=O)(O)O)[C@@H](O)[C@H]1O", "C10H16N5O13P3", 507.18, 506.9957, -1.63, 279.05, 308.50, 31, 3, 2, 0.50, "ZKHQWZAMYRWXGA-KQYNXXCUSA-N"),
        new("Thiophene", "c1ccsc1", "C4H4S", 84.14, 84.0034, 1.75, 0.0, 28.24, 5, 1, 1, 0.00, "YTPLMLYBLZKORZ-UHFFFAOYSA-N"),
        new("DMSO", "CS(=O)C", "C2H6OS", 78.13, 78.0139, -0.01, 17.07, 36.28, 4, 0, 0, 1.00),
        new("Phosphoric Acid", "OP(=O)(O)O", "H3O4P", 97.99, 97.9769, -0.93, 77.76, 87.57, 5, 0, 0, 0.00, "NBIIXXVUZAFLBC-UHFFFAOYSA-N"),

        // Isotopes (deuterated & 13C-labeled)
        new("DMSO-d6", "[2H]C([2H])([2H])S(=O)C([2H])([2H])[2H]", "C2H6OS", 84.17, 84.0516, -0.01, 17.07, 36.28, 4, 0, 0, 1.00),
        new("13C-Ethanol", "[13CH3]CO", "C2H6O", 47.06, 47.0452, 0.00, 20.23, 20.23, 3, 0, 0, 1.00),

        // Quaternary ammonium, zwitterion, multi-component salts & nitro ester
        new("Acetylcholine", "CC(=O)OCC[N+](C)(C)C", "C7H16NO2", 146.21, 146.1181, 0.26, 26.30, 26.30, 10, 0, 0, 0.86),
        new("Glycine Zwitterion", "[NH3+]CC(=O)[O-]", "C2H5NO2", 75.07, 75.0320, -3.02, 67.77, 67.77, 5, 0, 0, 0.50),
        new("Metformin HCl", "CN(C)C(=N)NC(=N)N.Cl", "C4H12ClN5", 165.63, 165.0781, -0.61, 88.99, 88.99, 10, 0, 0, 0.50),
        new("Sodium Acetate", "[Na+].CC(=O)[O-]", "C2H3NaO2", 82.03, 82.0031, -1.62, 40.13, 40.13, 5, 0, 0, 0.50),
        new("Sodium Chloride", "[Na+].[Cl-]", "ClNa", 58.44, 57.9586, -3.38, 0.0, 0.0, 2, 0, 0, 0.00, "FAPWRFPIFSIZLT-UHFFFAOYSA-M"),
        new("Nitroglycerin", "O=[N+]([O-])OCC(CO[N+](=O)[O-])O[N+](=O)[O-]", "C3H5N3O9", 227.09, 227.0026, -1.02, 157.11, 157.11, 15, 0, 0, 1.00),

        // Organometallics, metalloids & stable free radical
        new("Cisplatin", "[Pt](N)(N)(Cl)Cl", "Cl2H4N2Pt", 298.03, 296.9399, 0.20, 52.04, 52.04, 5, 0, 0, 0.00, "DQLATGHUWYMOKM-UHFFFAOYSA-L"),
        new("Ferrocene", "[Fe+2].[cH-]1cccc1.[cH-]1cccc1", "C10H10Fe", 186.04, 186.0132, 2.81, 0.0, 0.0, 11, 2, 2, 0.00, "KTWOOEGAPBSYNW-UHFFFAOYSA-N"),
        new("Trimethylsilanol", "C[Si](C)(C)O", "C3H10OSi", 90.20, 90.0501, 0.81, 20.23, 20.23, 5, 0, 0, 1.00),
        new("Ebselen", "O=c1c2ccccc2[se]n1-c1ccccc1", "C13H9NOSe", 274.19, 274.9849, 2.05, 22.00, 22.00, 16, 3, 3, 0.00),
        new("TEMPO", "CC1(C)CCCC(C)(C)N1[O]", "C9H18NO", 156.25, 156.1388, 2.38, 3.24, 3.24, 11, 1, 0, 1.00, "QYTDEUPAUMOIOP-UHFFFAOYSA-N"),

        // Small inorganic/allene/alkyne, cages, bridged, spiro & heterocycles
        new("Water", "O", "H2O", 18.02, 18.0106, -0.44, 0.0, 0.0, 1, 0, 0, 0.00, "XLYOFNOQVPJJNP-UHFFFAOYSA-N"),
        new("Methane", "C", "CH4", 16.04, 16.0313, 0.64, 0.0, 0.0, 1, 0, 0, 1.00, "VNWKTOKETHGBQD-UHFFFAOYSA-N"),
        new("Carbon Dioxide", "O=C=O", "CO2", 44.01, 43.9898, -0.55, 34.14, 34.14, 3, 0, 0, 0.00, "CURLTUGMZLYLDI-UHFFFAOYSA-N"),
        new("Acetylene", "C#C", "C2H2", 26.04, 26.0157, 0.33, 0.0, 0.0, 2, 0, 0, 0.00, "HSFWRNGVRCDJHI-UHFFFAOYSA-N"),
        new("Allene", "C=C=C", "C3H4", 40.07, 40.0313, 0.96, 0.0, 0.0, 3, 0, 0, 0.00, "IYABWNGZIDDRAK-UHFFFAOYSA-N"),
        new("Cyclopropane", "C1CC1", "C3H6", 42.08, 42.0470, 1.04, 0.0, 0.0, 3, 1, 0, 1.00, "LVZWSLJZHVFIQJ-UHFFFAOYSA-N"),
        new("Adamantane", "C1C2CC3CC1CC(C2)C3", "C10H16", 136.24, 136.1252, 2.83, 0.0, 0.0, 10, 3, 0, 1.00),
        new("Cubane", "C12C3C4C1C5C2C3C45", "C8H8", 104.15, 104.0626, 0.98, 0.0, 0.0, 8, 5, 0, 1.00, "TXWRERCHRDBNLG-UHFFFAOYSA-N"),
        new("Spiro[4.5]decane", "C1CCC2(C1)CCCCC2", "C10H18", 138.25, 138.1409, 3.51, 0.0, 0.0, 10, 2, 0, 1.00, "CTDQAGUNKPRERK-UHFFFAOYSA-N"),
        new("Naphthalene", "c1ccc2ccccc2c1", "C10H8", 128.17, 128.0626, 2.84, 0.0, 0.0, 10, 2, 2, 0.00, "UFWIBTONFRDIAS-UHFFFAOYSA-N"),
        new("Indole", "c1ccc2[nH]ccc2c1", "C8H7N", 117.15, 117.0578, 2.17, 15.79, 15.79, 9, 2, 2, 0.00, "SIKJAQJRHWYJAI-UHFFFAOYSA-N"),
        new("Pyrrole", "c1cc[nH]c1", "C4H5N", 67.09, 67.0422, 1.01, 15.79, 15.79, 5, 1, 1, 0.00, "KAESVJOAVNADME-UHFFFAOYSA-N"),
        new("Furan", "c1ccoc1", "C4H4O", 68.08, 68.0262, 1.28, 13.14, 13.14, 5, 1, 1, 0.00, "YLQBMQCUIZJEEH-UHFFFAOYSA-N"),
        new("2-Pyridone", "O=c1cccc[nH]1", "C5H5NO", 95.10, 95.0371, 0.37, 32.86, 32.86, 7, 1, 1, 0.00, "UBQKCCHYAOITMY-UHFFFAOYSA-N"),
        new("Uracil", "O=c1cc[nH]c(=O)[nH]1", "C4H4N2O2", 112.09, 112.0273, -0.94, 65.72, 65.72, 8, 1, 1, 0.00, "ISAKRJDGNUQOIC-UHFFFAOYSA-N"),
        new("Adenine", "Nc1ncnc2[nH]cnc12", "C5H5N5", 135.13, 135.0545, -0.06, 80.48, 80.48, 10, 2, 2, 0.00, "GFFGJBXGBJISGV-UHFFFAOYSA-N"),
        new("Azulene", "c1ccc2cccc2cc1", "C10H8", 128.17, 128.0626, 2.84, 0.0, 0.0, 10, 2, 2, 0.00, "CUFNKYGDVFVPHO-UHFFFAOYSA-N"),
        new("Biphenyl", "c1ccc(-c2ccccc2)cc1", "C12H10", 154.21, 154.0783, 3.35, 0.0, 0.0, 12, 2, 2, 0.00, "ZUOUZKKEUPVFJK-UHFFFAOYSA-N")
    ];

    public static void RunAll()
    {
        RunRealSubstanceBenchmarkMatrixTests();
        RunStereochemistryAndNotationParityTests();
        RunSmartsAndAdversarialGuardrailTests();
        RunOfflineCatalogCasQuickInsertAndSdfTests();
        RunFormulaRewriterAndTpsaOptionMatrixTests();
    }

    private static void AssertValidPng(byte[]? bytes, string label)
    {
        Check(bytes != null && bytes.Length > 200, $"{label}: PNG length > 200 B");
        Check(bytes![0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47,
            $"{label}: valid PNG signature");
    }

    private static void AssertValidSvg(string? svg, string label)
    {
        Check(!string.IsNullOrWhiteSpace(svg) && svg!.Contains("<svg"), $"{label}: valid SVG markup");
    }

    private static void AssertApprox(object actualObj, double expected, double tol, string label)
    {
        Check(actualObj is double d && Math.Abs(d - expected) <= tol,
            $"{label}: expected ~{expected.ToString("F4", CultureInfo.InvariantCulture)} (+/-{tol}), got {actualObj}");
    }

    private static void RunRealSubstanceBenchmarkMatrixTests()
    {
        foreach (var c in BenchmarkMatrix)
        {
            var info = MoleculeRenderer.TryGetMoleculeInfo(c.Smiles, resolvePubChem: false);
            Check(info != null, $"{c.Name}: TryGetMoleculeInfo succeeded");
            Check(info!.Formula == c.Formula, $"{c.Name}: Formula expected {c.Formula}, got {info.Formula}");
            Check(Math.Abs(info.MolecularWeight - c.Mw) <= 0.05, $"{c.Name}: MW ~ {c.Mw}, got {info.MolecularWeight}");
            Check(Math.Abs(info.ExactMass - c.ExactMass) <= 0.002, $"{c.Name}: ExactMass ~ {c.ExactMass}, got {info.ExactMass}");
            Check(info.HasLogP && Math.Abs(info.LogP - c.LogP) <= 0.05, $"{c.Name}: LogP ~ {c.LogP}, got {info.LogP}");
            Check(info.HasTpsa && Math.Abs(info.Tpsa - c.TpsaNoSp) <= 0.2, $"{c.Name}: TPSA(noSP) ~ {c.TpsaNoSp}, got {info.Tpsa}");
            Check(info.HasTpsaWithSP && Math.Abs(info.TpsaWithSP - c.TpsaWithSp) <= 0.2, $"{c.Name}: TPSA(+SP) ~ {c.TpsaWithSp}, got {info.TpsaWithSP}");
            Check(info.HeavyAtomCount == c.Heavy, $"{c.Name}: Heavy expected {c.Heavy}, got {info.HeavyAtomCount}");
            Check(info.RingCount == c.Rings, $"{c.Name}: Rings expected {c.Rings}, got {info.RingCount}");
            Check(info.AromaticRingCount == c.Aromatic, $"{c.Name}: Aromatic expected {c.Aromatic}, got {info.AromaticRingCount}");
            Check(Math.Abs(info.Fsp3 - c.Fsp3) <= 0.02, $"{c.Name}: Fsp3 ~ {c.Fsp3}, got {info.Fsp3}");
            Check(!string.IsNullOrEmpty(info.InChIKey) && info.InChIKey.Length == 27, $"{c.Name}: InChIKey length 27, got {info.InChIKey}");

            if (c.InchiKey != null)
                Check(info.InChIKey == c.InchiKey, $"{c.Name}: InChIKey expected {c.InchiKey}, got {info.InChIKey}");

            // Also verify Ebselen exact mass is heavier than average MW (Se monoisotopic 79.9165 vs 78.971)
            if (c.Name == "Ebselen")
                Check(info.ExactMass > info.MolecularWeight, "Ebselen monoisotopic exact mass exceeds average MW (Se isotope profile)");

            // Verify all Excel SMILES2* scalar UDFs match
            Check(Equals(Functions.SmilesFormula(c.Smiles), c.Formula), $"{c.Name}: SMILES2FORMULA");
            AssertApprox(Functions.SmilesMW(c.Smiles), c.Mw, 0.05, $"{c.Name}: SMILES2MW");
            AssertApprox(Functions.SmilesExactMass(c.Smiles), c.ExactMass, 0.002, $"{c.Name}: SMILES2EXACTMASS");
            AssertApprox(Functions.SmilesLogP(c.Smiles), c.LogP, 0.05, $"{c.Name}: SMILES2LOGP");
            AssertApprox(Functions.SmilesTpsa(c.Smiles), c.TpsaNoSp, 0.2, $"{c.Name}: SMILES2TPSA(false)");
            AssertApprox(Functions.SmilesTpsa(c.Smiles, true), c.TpsaWithSp, 0.2, $"{c.Name}: SMILES2TPSA(true)");
            Check(Equals(Functions.SmilesHeavy(c.Smiles), c.Heavy), $"{c.Name}: SMILES2HEAVY");
            Check(Equals(Functions.SmilesRings(c.Smiles), c.Rings), $"{c.Name}: SMILES2RINGS");
            Check(Equals(Functions.SmilesAromatic(c.Smiles), c.Aromatic), $"{c.Name}: SMILES2AROMATIC");
            AssertApprox(Functions.SmilesFsp3(c.Smiles), c.Fsp3, 0.02, $"{c.Name}: SMILES2FSP3");

            AssertValidPng(MoleculeRenderer.Render(c.Smiles, RenderOptions.Default), c.Name);
            AssertValidSvg(MoleculeRenderer.RenderSvg(c.Smiles, RenderOptions.Default), c.Name);
        }

        PassSuite($"Unit: Data-Driven Real-Substance Benchmark Matrix ({BenchmarkMatrix.Length} Diverse Compounds)");
    }

    private static void RunStereochemistryAndNotationParityTests()
    {
        // 1. SMILES atom-ordering & Kekule permutation invariance
        string[][] invariantGroups =
        [
            ["c1ccccc1", "C1=CC=CC=C1", "c1(cccc1) | self" /* handled below */],
        ];
        _ = invariantGroups;
        (string Label, string A, string B, string C)[] permuted =
        [
            ("Benzene", "c1ccccc1", "C1=CC=CC=C1", "C=1C=CC=CC=1"),
            ("Aspirin", "CC(=O)Oc1ccccc1C(=O)O", "O=C(O)c1ccccc1OC(C)=O", "CC(=O)OC1=CC=CC=C1C(=O)O"),
            ("Caffeine", "Cn1cnc2c1c(=O)n(C)c(=O)n2C", "CN1C=NC2=C1C(=O)N(C(=O)N2C)C", "Cn1c(=O)c2c(ncn2C)n(C)c1=O"),
            ("Benzoic Acid", "OC(=O)c1ccccc1", "c1ccc(cc1)C(=O)O", "O=C(O)C1=CC=CC=C1"),
            ("Pyridine", "c1ccncc1", "C1=CC=NC=C1", "n1ccccc1"),
            ("Naphthalene", "c1ccc2ccccc2c1", "C1=CC2=CC=CC=C2C=C1", "c12ccccc1cccc2"),
            ("Indole", "c1ccc2[nH]ccc2c1", "C1=CC=C2C(=C1)C=CN2", "[nH]1ccc2ccccc12"),
            ("Salicylic Acid", "O=C(O)c1ccccc1O", "Oc1ccccc1C(=O)O", "OC(=O)C1=CC=CC=C1O")
        ];

        foreach (var g in permuted)
        {
            var iA = MoleculeRenderer.TryGetMoleculeInfo(g.A)!;
            var iB = MoleculeRenderer.TryGetMoleculeInfo(g.B)!;
            var iC = MoleculeRenderer.TryGetMoleculeInfo(g.C)!;
            Check(iA.InChIKey == iB.InChIKey && iA.InChIKey == iC.InChIKey,
                $"{g.Label}: InChIKey invariant across permutations ({iA.InChIKey} vs {iB.InChIKey} vs {iC.InChIKey})");
            Check(iA.Formula == iB.Formula && iA.Formula == iC.Formula, $"{g.Label}: Formula invariant");
            Check(Math.Abs(iA.LogP - iB.LogP) < 1e-6 && Math.Abs(iA.LogP - iC.LogP) < 1e-6, $"{g.Label}: LogP invariant");
            Check(Math.Abs(iA.Tpsa - iB.Tpsa) < 1e-6 && Math.Abs(iA.Tpsa - iC.Tpsa) < 1e-6, $"{g.Label}: TPSA invariant");
            Check(iA.RingCount == iB.RingCount && iA.AromaticRingCount == iC.AromaticRingCount, $"{g.Label}: Ring/Arom invariant");
        }

        // 2. Enantiomers, Diastereomers, Meso & E/Z Isomers
        var rIbu = MoleculeRenderer.TryGetMoleculeInfo("CC(C)Cc1ccc([C@H](C)C(=O)O)cc1")!;
        var sIbu = MoleculeRenderer.TryGetMoleculeInfo("CC(C)Cc1ccc([C@@H](C)C(=O)O)cc1")!;
        Check(rIbu.InChIKey.Substring(0, 14) == sIbu.InChIKey.Substring(0, 14), "(R)/(S)-Ibuprofen share 14-char connectivity layer");
        Check(rIbu.InChIKey != sIbu.InChIKey, "(R)/(S)-Ibuprofen have distinct stereochemical InChIKey layers");

        var tartRR = MoleculeRenderer.TryGetMoleculeInfo("O=C(O)[C@H](O)[C@@H](O)C(=O)O")!;
        var tartSS = MoleculeRenderer.TryGetMoleculeInfo("O=C(O)[C@@H](O)[C@H](O)C(=O)O")!;
        var mesoA = MoleculeRenderer.TryGetMoleculeInfo("O=C(O)[C@H](O)[C@H](O)C(=O)O")!;
        var mesoB = MoleculeRenderer.TryGetMoleculeInfo("O=C(O)[C@@H](O)[C@@H](O)C(=O)O")!;
        Check(tartRR.InChIKey == "FEWJPZIEWOKRBE-JCYAYHJZSA-N", $"2R,3R-Tartaric IK, got {tartRR.InChIKey}");
        Check(tartSS.InChIKey == "FEWJPZIEWOKRBE-LWMBPPNESA-N", $"2S,3S-Tartaric IK, got {tartSS.InChIKey}");
        Check(mesoA.InChIKey == "FEWJPZIEWOKRBE-XIXRPRMCSA-N" && mesoA.InChIKey == mesoB.InChIKey,
            $"Both meso-Tartaric SMILES yield identical InChIKey FEWJPZIEWOKRBE-XIXRPRMCSA-N");

        var transDfe = MoleculeRenderer.TryGetMoleculeInfo("F/C=C/F")!;
        var cisDfe = MoleculeRenderer.TryGetMoleculeInfo(@"F/C=C\F")!;
        Check(transDfe.InChIKey == "WFLOTYSKFUPZQB-OWOJBTEDSA-N", $"trans-DFE IK, got {transDfe.InChIKey}");
        Check(cisDfe.InChIKey == "WFLOTYSKFUPZQB-UPHRSURJSA-N", $"cis-DFE IK, got {cisDfe.InChIKey}");

        // 3. Full SMILES <-> KET parity on 16 diverse compounds, plus SMILES <-> InChI parity
        // Note: Standard InChI (1S) uses mobile-H tautomers; secondary amides (e.g. Paracetamol, Penicillin G,
        // Atorvastatin, Bortezomib, Paclitaxel, Cyclosporine A) preserve topology/formula/MW/Fsp3/InChIKey
        // across SMILES <-> InChI, while bond-level descriptors (LogP, TPSA) depend on reconstructed tautomer.
        string[] ketAndCoreInchiParityNames =
        [
            "Aspirin", "Caffeine", "Paracetamol", "(S)-Ibuprofen", "Nicotine", "Atorvastatin",
            "Penicillin G", "Artemisinin", "Cholesterol", "Morphine", "Paclitaxel", "ATP",
            "Bortezomib", "Cyclosporine A", "Glyphosate", "Sulfamethoxazole"
        ];
        var nonTautomericLogPTpsaNames = new HashSet<string>(StringComparer.Ordinal)
        {
            "Aspirin", "Caffeine", "(S)-Ibuprofen", "Nicotine", "Artemisinin",
            "Cholesterol", "Morphine", "ATP", "Glyphosate", "Sulfamethoxazole"
        };

        foreach (var name in ketAndCoreInchiParityNames)
        {
            var c = BenchmarkMatrix.First(x => x.Name == name);
            var inchi = MoleculeRenderer.TryConvertToInchi(c.Smiles);
            var ket = MoleculeRenderer.TryConvertToKet(c.Smiles);
            Check(!string.IsNullOrEmpty(inchi), $"{name}: SMILES -> InChI succeeded");
            Check(!string.IsNullOrEmpty(ket), $"{name}: SMILES -> KET succeeded");

            Check(Equals(Functions.InchiFormula(inchi!), Functions.SmilesFormula(c.Smiles)), $"{name}: INCHI2FORMULA parity");
            Check(Equals(Functions.KetFormula(ket!), Functions.SmilesFormula(c.Smiles)), $"{name}: KET2FORMULA parity");

            AssertApprox(Functions.InchiMW(inchi!), c.Mw, 0.05, $"{name}: INCHI2MW parity");
            AssertApprox(Functions.KetMW(ket!), c.Mw, 0.05, $"{name}: KET2MW parity");

            AssertApprox(Functions.InchiExactMass(inchi!), c.ExactMass, 0.002, $"{name}: INCHI2EXACTMASS parity");
            AssertApprox(Functions.KetExactMass(ket!), c.ExactMass, 0.002, $"{name}: KET2EXACTMASS parity");

            // KET always preserves exact bond orders (exact LogP & TPSA even for amides)
            AssertApprox(Functions.KetLogP(ket!), c.LogP, 0.05, $"{name}: KET2LOGP parity");
            AssertApprox(Functions.KetTpsa(ket!), c.TpsaNoSp, 0.2, $"{name}: KET2TPSA(false) parity");
            AssertApprox(Functions.KetTpsa(ket!, true), c.TpsaWithSp, 0.2, $"{name}: KET2TPSA(true) parity");

            if (nonTautomericLogPTpsaNames.Contains(name))
            {
                AssertApprox(Functions.InchiLogP(inchi!), c.LogP, 0.05, $"{name}: INCHI2LOGP parity");
                AssertApprox(Functions.InchiTpsa(inchi!), c.TpsaNoSp, 0.2, $"{name}: INCHI2TPSA(false) parity");
                AssertApprox(Functions.InchiTpsa(inchi!, true), c.TpsaWithSp, 0.2, $"{name}: INCHI2TPSA(true) parity");
            }

            Check(Equals(Functions.InchiHeavy(inchi!), c.Heavy) && Equals(Functions.KetHeavy(ket!), c.Heavy), $"{name}: HEAVY parity");
            Check(Equals(Functions.InchiRings(inchi!), c.Rings) && Equals(Functions.KetRings(ket!), c.Rings), $"{name}: RINGS parity");
            Check(Equals(Functions.InchiAromatic(inchi!), c.Aromatic) && Equals(Functions.KetAromatic(ket!), c.Aromatic), $"{name}: AROMATIC parity");
            AssertApprox(Functions.InchiFsp3(inchi!), c.Fsp3, 0.02, $"{name}: INCHI2FSP3 parity");
            AssertApprox(Functions.KetFsp3(ket!), c.Fsp3, 0.02, $"{name}: KET2FSP3 parity");

            var ik = Functions.SmilesInchiKey(c.Smiles);
            Check(Equals(Functions.InchiKey(inchi!), ik) && Equals(Functions.KetInchiKey(ket!), ik), $"{name}: InChIKey parity");
            Check(Functions.InchiSmiles(inchi!) is string isml && !string.IsNullOrEmpty(isml), $"{name}: INCHI2SMILES non-empty");
            Check(Functions.KetSmiles(ket!) is string ksml && !string.IsNullOrEmpty(ksml), $"{name}: KET2SMILES non-empty");
        }

        // Standard InChI (1S) mobile-H amide/imidol nuance: Paracetamol amide LogP=1.35 vs InChI imidic acid tautomer LogP=2.00
        var paraInchi = MoleculeRenderer.TryConvertToInchi("CC(=O)Nc1ccc(O)cc1")!;
        Check(Functions.InchiLogP(paraInchi) is double paraInchiLogP && Math.Abs(paraInchiLogP - 2.00) <= 0.05,
            $"Paracetamol Standard InChI reconstructs mobile-H tautomer (LogP ~ 2.00), got {Functions.InchiLogP(paraInchi)}");

        // KET highlight dearomatization + chiral rotation regression checks
        var ibuKet = MoleculeRenderer.TryConvertToKet("CC(C)Cc1ccc(cc1)[C@@H](C)C(=O)O")!;
        var hlKetOpt = RenderOptions.Parse("white", "hl:c1ccccc1 num stereo", null);
        AssertValidPng(MoleculeRenderer.RenderKet(ibuKet, hlKetOpt), "Highlighted chiral KET PNG");
        AssertValidSvg(MoleculeRenderer.RenderSvg("KET:" + ibuKet, hlKetOpt), "Highlighted chiral KET SVG");

        foreach (var rot in new[] { "rot90", "rot180", "rot270" })
        {
            var rotOpt = RenderOptions.Parse("trans", "stereo", rot);
            AssertValidPng(MoleculeRenderer.Render("CC(C)Cc1ccc(cc1)[C@@H](C)C(=O)O", rotOpt), $"Chiral SMILES {rot}");
            AssertValidPng(MoleculeRenderer.RenderKet(ibuKet, rotOpt), $"Chiral KET {rot}");
        }

        PassSuite("Unit: Stereochemistry, Permutation Invariance & SMILES/InChI/KET Round-Trip Parity");
    }

    private static void RunSmartsAndAdversarialGuardrailTests()
    {
        // 1. SMARTS substructure truth table across SMILES.MATCH, INCHI.MATCH & KET.MATCH
        (string Name, string Target, string Smarts, bool Expected)[] smartsCases =
        [
            ("Benzene in Aspirin", "CC(=O)Oc1ccccc1C(=O)O", "c1ccccc1", true),
            ("Benzene in Caffeine (negative)", "Cn1cnc2c1c(=O)n(C)c(=O)n2C", "c1ccccc1", false),
            ("Pyridine in Nicotine", "CN1CCC[C@H]1c2cccnc2", "c1ccncc1", true),
            ("Carboxyl in Aspirin", "CC(=O)Oc1ccccc1C(=O)O", "[CX3](=O)[OX2H1]", true),
            ("Carboxyl in Caffeine (negative)", "Cn1cnc2c1c(=O)n(C)c(=O)n2C", "[CX3](=O)[OX2H1]", false),
            ("Beta-lactam in Penicillin G", "CC1([C@@H](N2[C@H](S1)[C@@H](C2=O)NC(=O)Cc3ccccc3)C(=O)O)C", "O=C1CCN1", true),
            ("Sulfonamide in Sulfamethoxazole", "Cc1cc(NS(=O)(=O)c2ccc(N)cc2)no1", "S(=O)(=O)N", true),
            ("Halogen in Atorvastatin", "CC(C)c1c(C(=O)Nc2ccccc2)c(c(n1CC[C@@H](C[C@@H](CC(=O)O)O)O)c3ccc(F)cc3)c4ccccc4", "[F,Cl,Br,I]", true),
            ("Recursive nitro in Nitroglycerin", "O=[N+]([O-])OCC(CO[N+](=O)[O-])O[N+](=O)[O-]", "[$([NX3](=O)=O),$([NX3+](=O)[O-])]", true),
            ("5-membered ring in Indole", "c1ccc2[nH]ccc2c1", "[r5]", true),
            ("5-membered ring in Naphthalene (negative)", "c1ccc2ccccc2c1", "[r5]", false),
            ("Phenol in Paracetamol", "CC(=O)Nc1ccc(O)cc1", "c1ccccc1[OH]", true),
            ("Guanidine in Metformin HCl salt", "CN(C)C(=N)NC(=N)N.Cl", "C(=N)N", true)
        ];

        foreach (var tc in smartsCases)
        {
            var inchi = MoleculeRenderer.TryConvertToInchi(tc.Target)!;
            var ket = MoleculeRenderer.TryConvertToKet(tc.Target)!;
            Check(Equals(Functions.SmilesMatch(tc.Target, tc.Smarts), tc.Expected), $"{tc.Name} (SMILES.MATCH)");
            Check(Equals(Functions.InchiMatch(inchi, tc.Smarts), tc.Expected), $"{tc.Name} (INCHI.MATCH)");
            Check(Equals(Functions.KetMatch(ket, tc.Smarts), tc.Expected), $"{tc.Name} (KET.MATCH)");
        }

        // 2. Complex SMARTS highlight token parsing + live PNG rendering
        string[][] hlRenderCases =
        [
            [@"hl:F/C=C/F", "F/C=C/F"],
            [@"bw+hl:F/C=C/F", "F/C=C/F"],
            ["hl:[CX3;R0](=O)[OX2H1] bw", "CC(=O)Oc1ccccc1C(=O)O"],
            ["hl:[$([NX3](=O)=O)] num stereo", "O=[N+]([O-])OCC(CO[N+](=O)[O-])O[N+](=O)[O-]"],
            ["hl:O=[N+]([O-]) white", "O=[N+]([O-])c1ccccc1"]
        ];
        foreach (var pair in hlRenderCases)
        {
            var opt = RenderOptions.Parse("#1E293B", pair[0], null);
            Check(!string.IsNullOrWhiteSpace(opt.HighlightQuery), $"Complex hl parsed: {pair[0]}");
            AssertValidPng(MoleculeRenderer.Render(pair[1], opt), $"Complex hl PNG ({pair[0]})");
        }

        // 3. Adversarial / malformed / oversized / format-mismatch / query / reaction guardrails
        string[] invalidInputs =
        [
            "", "   ", "c1ccccc", "CC(=O", "[Na", "C(C)(C)(C)(C)C",
            "InChI=1S/bogus", "{}", "{\"root\":{\"nodes\":[]}}", new string('C', 16385)
        ];
        foreach (var bad in invalidInputs)
        {
            var mw = Functions.SmilesMW(bad);
            var lp = Functions.SmilesLogP(bad);
            Check(mw.Equals(ExcelError.ExcelErrorValue) || mw.Equals(ExcelError.ExcelErrorNA), $"Invalid MW guarded: '{(bad.Length > 20 ? bad.Substring(0, 20) + "..." : bad)}'");
            Check(lp.Equals(ExcelError.ExcelErrorValue) || lp.Equals(ExcelError.ExcelErrorNA), "Invalid LogP guarded");
            var mt = Functions.SmilesMatch(bad, "c1ccccc1");
            Check(mt.Equals(ExcelError.ExcelErrorValue), "Invalid SMILES.MATCH target returns #VALUE!");
        }

        Check(Equals(Functions.SmilesMatch("c1ccccc1", "[invalid_smarts"), ExcelError.ExcelErrorValue), "Invalid SMARTS query returns #VALUE!");
        Check(Equals(Functions.InchiMatch("CCO", "C"), ExcelError.ExcelErrorValue), "INCHI.MATCH rejects plain SMILES");
        Check(Equals(Functions.KetMatch("CCO", "C"), ExcelError.ExcelErrorValue), "KET.MATCH rejects plain SMILES");
        Check(Equals(Functions.InchiLogP("c1ccccc1"), ExcelError.ExcelErrorValue), "INCHI2LOGP rejects plain SMILES");
        Check(Equals(Functions.KetLogP("c1ccccc1"), ExcelError.ExcelErrorValue), "KET2LOGP rejects plain SMILES");
        Check(Equals(Functions.KetImg("CCO"), ExcelError.ExcelErrorValue), "KET2IMG rejects plain SMILES");
        Check(Equals(Functions.SmilesFsp3("O"), 0.0), "SMILES2FSP3 returns 0.0 for 0-carbon inorganic molecule (Water)");

        // Reactions & query R-groups
        const string rxn = "O=C(O)c1ccccc1O.CC(=O)OC(=O)C>>CC(=O)Oc1ccccc1C(=O)O";
        Check(Equals(Functions.SmilesMW(rxn), ExcelError.ExcelErrorNA), "Reaction in SMILES2MW returns #N/A");
        Check(Equals(Functions.SmilesLogP(rxn), ExcelError.ExcelErrorNA), "Reaction in SMILES2LOGP returns #N/A");
        Check(Equals(Functions.SmilesName(rxn), ExcelError.ExcelErrorNA), "Reaction in SMILES2NAME returns #N/A");
        Check(Equals(Functions.SmilesMatch(rxn, "c1ccccc1"), ExcelError.ExcelErrorValue), "Reaction in SMILES.MATCH returns #VALUE!");

        foreach (var q in new[] { "*c1ccccc1", "[R1]c1ccc(O)cc1" })
        {
            AssertValidPng(MoleculeRenderer.Render(q, RenderOptions.Default), $"Query {q} PNG");
            Check(Equals(Functions.SmilesMW(q), ExcelError.ExcelErrorNA), $"Query {q} SMILES2MW returns #N/A");
            Check(Equals(Functions.SmilesLogP(q), ExcelError.ExcelErrorNA), $"Query {q} SMILES2LOGP returns #N/A");
            Check(Equals(Functions.SmilesTpsa(q), ExcelError.ExcelErrorNA), $"Query {q} SMILES2TPSA returns #N/A");
        }

        PassSuite("Unit: SMARTS Substructure Matrix, Complex Highlighting & Adversarial Input Guardrails");
    }

    private static void RunOfflineCatalogCasQuickInsertAndSdfTests()
    {
        // 1. Validate EVERY key in PubChemResolver.BuiltInNameOrCasKeys offline
        var catalogKeys = PubChemResolver.BuiltInNameOrCasKeys.ToList();
        Check(catalogKeys.Count >= 50, $"Built-in offline dictionary has >= 50 keys, got {catalogKeys.Count}");
        foreach (var key in catalogKeys)
        {
            if (PubChemResolver.IsNumericTriplet(key))
                Check(PubChemResolver.IsValidCasNumber(key), $"Built-in CAS key '{key}' passes mathematical check-digit validation");

            var smiles = PubChemResolver.ResolveNameToSmilesOffline(key);
            Check(!string.IsNullOrWhiteSpace(smiles), $"Built-in key '{key}' resolves offline");
            var info = MoleculeRenderer.TryGetMoleculeInfo(smiles, resolvePubChem: false);
            Check(info != null && !string.IsNullOrWhiteSpace(info.Formula) && info.MolecularWeight > 0 && info.InChIKey.Length == 27,
                $"Built-in key '{key}' -> '{smiles}' computes valid Formula/MW/InChIKey");
            AssertValidPng(MoleculeRenderer.Render(smiles!, RenderOptions.Default), $"Built-in '{key}'");
        }

        // 2. Invalid CAS check-digit numbers must fail in 0ms without hitting network
        string[] validCas = ["50-78-2", "58-08-2", "103-90-2", "15687-27-1", "33069-62-4", "141-78-6"];
        foreach (var c in validCas)
            Check(PubChemResolver.IsValidCasNumber(c) && MoleculeActionHelper.IsValidCasNumber(c), $"Valid CAS: {c}");

        string[] invalidCas = ["50-78-3", "58-08-9", "103-90-0", "00-00-0", "1-23-4", "12345678-90-1"];
        var sw = Stopwatch.StartNew();
        foreach (var badCas in invalidCas)
        {
            Check(!PubChemResolver.IsValidCasNumber(badCas), $"Invalid CAS rejected: {badCas}");
            Check(PubChemResolver.ResolveNameToSmiles(badCas, timeoutMs: 3000) == null, $"PubChemResolver returns null locally for invalid CAS {badCas}");
            Check(Equals(Functions.NameSmiles(badCas), ExcelError.ExcelErrorNA), $"NAME2SMILES returns #N/A for invalid CAS {badCas}");
        }
        sw.Stop();
        Check(sw.ElapsedMilliseconds < 250, $"All invalid CAS rejections ran offline in <250ms (actual: {sw.ElapsedMilliseconds}ms)");

        // 3. QuickInsert ambiguity classification + Ketcher preload safety
        string[] ambiguousAlpha = ["I", "C", "N", "O", "CO", "CN", "NO", "CC", "CS", "IN", "ON", "COP", "CON", "SIN", "PIN", "BIN", "ICON", "COIN", "COCO"];
        foreach (var word in ambiguousAlpha)
        {
            Check(MoleculeActionHelper.IsAmbiguousPlainAlphaSmiles(word), $"Plain word '{word}' flagged as ambiguous");
            var ok = MoleculeActionHelper.TryResolveQuickInsertInput(word, RenderOptions.Default, out var res, resolvePubChem: false);
            if (ok && res != null)
                Check(res.RequiresConfirmation && res.SourceKind == QuickInsertSourceKind.AmbiguousSmiles,
                    $"QuickInsert requires confirmation for '{word}'");
            Check(StructureMetadata.ExtractPreloadForKetcher(word) == "",
                $"ExtractPreloadForKetcher blocks ambiguous text '{word}'");
        }

        foreach (var explicitInput in new[] { "CCO", "CC(=O)Oc1ccccc1C(=O)O", "50-78-2", "InChI=1S/C2H6O/c1-2-3/h3H,2H2,1H3" })
        {
            var ok = MoleculeActionHelper.TryResolveQuickInsertInput(explicitInput, RenderOptions.Default, out var res, resolvePubChem: false);
            Check(ok && res != null && !res.RequiresConfirmation, $"Explicit/CAS input '{explicitInput}' inserts without confirmation");
        }

        foreach (var chemName in new[] { "Aspirin", "Caffeine", "THF", "DMSO", "Salicylic Acid" })
        {
            var ok = MoleculeActionHelper.TryResolveQuickInsertInput(chemName, RenderOptions.Default, out var res, resolvePubChem: false);
            Check(ok && res != null && res.RequiresConfirmation && res.SourceKind == QuickInsertSourceKind.ChemicalName,
                $"Chemical name '{chemName}' resolves offline and requires confirmation");
        }

        foreach (var badInput in new[] { "50-78-3", "Review manuscript Figure 2 before submission." })
        {
            Check(!MoleculeActionHelper.TryResolveQuickInsertInput(badInput, RenderOptions.Default, out _, resolvePubChem: false),
                $"Rejects non-structure / invalid CAS '{badInput}'");
        }

        // 4. SdfManager resilience on 10 diverse compounds + corrupt record + CRLF + missing title + maxRecords
        var recordsInput = BenchmarkMatrix.Take(10).Select(c => new SdfRecord
        {
            Name = c.Name == "Caffeine" ? "Caffeine α-form (UTF-8)" : c.Name,
            Smiles = c.Smiles,
            Properties = new Dictionary<string, string> { ["Series"] = "Benchmark-10" }
        }).ToList();

        var sdfBytes = SdfManager.WriteSdfBytes(recordsInput);
        var roundTrip = SdfManager.ReadSdfBytes(sdfBytes);
        Check(roundTrip.Count == 10, $"SDF 10-compound round-trip count=10, got {roundTrip.Count}");
        for (int i = 0; i < 10; i++)
            Check(roundTrip[i].Formula == BenchmarkMatrix[i].Formula, $"SDF round-trip [{i}] formula {roundTrip[i].Formula}");

        // Inject a malformed V2000 record in the middle and verify remaining valid records survive
        var sdfText = Encoding.UTF8.GetString(sdfBytes);
        int sepIdx = sdfText.IndexOf("$$$$", StringComparison.Ordinal);
        string corruptBlock = "\nCorruptMol\n  BadProg\n\n999999999 BROKEN V2000\nM  END\n$$$$";
        var mixedBytes = Encoding.UTF8.GetBytes(sdfText.Insert(sepIdx + 4, corruptBlock));
        var recovered = SdfManager.ReadSdfBytes(mixedBytes);
        Check(recovered.Count == 10 && recovered[9].Formula == BenchmarkMatrix[9].Formula,
            $"SDF skips corrupt record cleanly and recovers all 10 valid records, got {recovered.Count}");

        // Fallback title from <NAME> property + LF-only line endings + maxRecords limit
        string noTitleRecord = sdfText.Substring(0, sdfText.IndexOf('\n'));
        string lfSdf = ("\n" + sdfText.Substring(sdfText.IndexOf('\n') + 1))
            .Replace("\r\n", "\n")
            .Replace("M  END\n", "M  END\n>  <NAME>\nFallbackPropertyTitle\n\n");
        _ = noTitleRecord;
        var lfParsed = SdfManager.ReadSdfBytes(Encoding.UTF8.GetBytes(lfSdf), maxRecords: 4);
        Check(lfParsed.Count == 4, $"SDF LF-only + maxRecords=4 returns 4, got {lfParsed.Count}");
        Check(lfParsed[0].Name == "FallbackPropertyTitle", $"SDF falls back to <NAME> property, got '{lfParsed[0].Name}'");

        Check(SdfManager.ReadSdfBytes(Array.Empty<byte>()).Count == 0, "Empty SDF bytes returns 0 records");
        Check(SdfManager.ReadSdfBytes(Encoding.UTF8.GetBytes("not an sdf file")).Count == 0, "Garbage SDF bytes returns 0 records");

        PassSuite("Unit: Built-In Offline Catalog, CAS Checksum Guard, Quick-Insert & SDF Resilience");
    }

    private static void RunFormulaRewriterAndTpsaOptionMatrixTests()
    {
        // 1. Batch Style ExcelFormulaRewriter stress matrix
        (string Input, string Bg, string Style, FormulaRewriteStatus ExpectedStatus, string[] MustContain, string[] MustNotContain)[] matrix =
        [
            ("=SMILES2IMG('Lab''s Sheet'!A2, , \"bw num stereo\", \"rot90\")", "trans", "white",
                FormulaRewriteStatus.Updated,
                ["'Lab''s Sheet'!A2", "\"white num stereo\"", "\"rot90\""],
                ["bw"]),
            ("=SMILES2IMG(SUBSTITUTE($B$2,\"[R]\",A2),\"white\",\"hl:[CX3](=O)[OX2H1] bw\",\"rot270\")", "trans", "color",
                FormulaRewriteStatus.Updated,
                ["SUBSTITUTE($B$2,\"[R]\",A2)", "\"color hl:[CX3](=O)[OX2H1]\"", "\"rot270\""],
                ["\"white\""]),
            ("=@_xll.INCHI2IMG(B5, \"#0F172A\", \"ink=#1E40AF num h all-c\")", "white", "bw",
                FormulaRewriteStatus.Updated,
                ["=@_xll.INCHI2IMG(B5,", "\"white\"", "\"bw num h all-c\""],
                ["#1E40AF"]),
            ("=smiles2img.float(D4, \"bw\")", "trans", "white",
                FormulaRewriteStatus.Updated,
                ["=smiles2img.float(D4, \"trans\", \"white\")"],
                ["bw"]),
            ("=KET2IMG(C6, , \"hl:\"&$B$3, \"rot180\")", "#0F172A", "white",
                FormulaRewriteStatus.Updated,
                ["=KET2IMG(C6, \"#0F172A\", \"white \" & (\"hl:\"&$B$3), \"rot180\")"],
                []),
            ("=IF(A2=\"\",\"\",SMILES2IMG(A2))", "trans", "bw",
                FormulaRewriteStatus.SkippedComposite,
                ["=IF(A2=\"\",\"\",SMILES2IMG(A2))"],
                []),
            ("=LET(x,A2,SMILES2IMG(x))", "trans", "bw",
                FormulaRewriteStatus.SkippedComposite,
                ["=LET(x,A2,SMILES2IMG(x))"],
                []),
            ("=SMILES2IMG(A2)&\"\"", "trans", "bw",
                FormulaRewriteStatus.SkippedComposite,
                [],
                []),
            ("=\"SMILES2IMG(A2)\"", "trans", "bw",
                FormulaRewriteStatus.NotApplicable,
                ["=\"SMILES2IMG(A2)\""],
                []),
            ("=SUM(A1:A10)", "trans", "bw",
                FormulaRewriteStatus.NotApplicable,
                ["=SUM(A1:A10)"],
                [])
        ];

        foreach (var tc in matrix)
        {
            var status = ExcelFormulaRewriter.TryRewriteImageStyle(tc.Input, tc.Bg, tc.Style, out var rewritten);
            Check(status == tc.ExpectedStatus, $"Rewriter status for '{tc.Input}': expected {tc.ExpectedStatus}, got {status}");
            foreach (var sub in tc.MustContain)
                Check(rewritten.Contains(sub), $"Rewriter '{tc.Input}' output '{rewritten}' contains '{sub}'");
            foreach (var bad in tc.MustNotContain)
                Check(!rewritten.Contains(bad), $"Rewriter '{tc.Input}' output '{rewritten}' omits '{bad}'");
        }

        // Multi-cycle idempotency on dynamic hl: expression (no nested "bw " & ("white " & ...) accumulation)
        var dyn = "=SMILES2IMG(A2, , \"hl:\" & $B$3, \"rot90\")";
        Check(ExcelFormulaRewriter.TryRewriteImageStyle(dyn, "trans", "bw", out dyn) == FormulaRewriteStatus.Updated, "Cycle 1 ok");
        Check(ExcelFormulaRewriter.TryRewriteImageStyle(dyn, "white", "color", out dyn) == FormulaRewriteStatus.Updated, "Cycle 2 ok");
        Check(ExcelFormulaRewriter.TryRewriteImageStyle(dyn, "trans", "white", out dyn) == FormulaRewriteStatus.Updated, "Cycle 3 ok");
        Check(dyn == "=SMILES2IMG(A2, \"trans\", \"white \" & (\"hl:\" & $B$3), \"rot90\")", $"Multi-cycle idempotent output, got: {dyn}");

        // 2. RenderOptions background, ink, delimiters & silent legacy flip ignore
        var optBg = RenderOptions.Parse("bg=FFFF00", "line=#1e40af;num|stereo,h+all-c+hl:[N+](=O)[O-]", "ROT270");
        Check(optBg.Background == "#ffff00" && optBg.BaseColor == "#1e40af" && optBg.AtomNumbers && optBg.StereoLabels &&
              optBg.UnfoldHydrogens && optBg.AllCarbons && optBg.Transform == "rot270" && optBg.HighlightQuery == "[N+](=O)[O-]",
            "RenderOptions parses all background, ink, delimiter, and bracketed SMARTS variants");

        foreach (var transparentAlias in new[] { "", "trans", "transparent", "none", "clear", "nobg" })
            Check(RenderOptions.Parse(transparentAlias, null, null).Background == "transparent", $"Transparent alias '{transparentAlias}'");

        // 3. Functions.SmilesTpsa include_sp optional parameter matrix on Thiophene (0.0 vs 28.24)
        foreach (var spTrue in new object[] { true, 1.0, "true", "1", "yes", "sp", "s", "p", "pubchem", "ALL" })
            AssertApprox(Functions.SmilesTpsa("c1ccsc1", spTrue), 28.24, 0.2, $"TPSA include_sp='{spTrue}'");
        foreach (var spFalse in new object?[] { null, false, 0.0, "false", "0", "no", "" })
            AssertApprox(Functions.SmilesTpsa("c1ccsc1", spFalse), 0.0, 0.2, $"TPSA include_sp='{spFalse}'");

        PassSuite("Unit: ExcelFormulaRewriter & RenderOptions Edge-Case Stress Matrix");
    }
}
