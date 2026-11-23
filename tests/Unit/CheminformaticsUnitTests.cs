using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using Smiles2Img;
using Smiles2Img.Editor;
using Smiles2Img.Rendering;
using static Smiles2Img.Tests.TestAssert;

namespace Smiles2Img.Tests.Unit;

internal static class CheminformaticsUnitTests
{
    public static void RunAll()
    {
        RunCorePropertyAndInchiTests();
        RunPhysicochemicalAndRuleOf5Tests();
        RunStructuralComplexityTests();
        RunSubstructureAndSimilarityTests();
        RunPubChemAndNameResolutionTests();
        RunMultiFormatSummariesAndSdfTests();
    }

    private static void RunCorePropertyAndInchiTests()
    {
        var aspirinInfo = MoleculeRenderer.TryGetMoleculeInfo("CC(=O)Oc1ccccc1C(=O)O");
        Check(aspirinInfo != null, "TryGetMoleculeInfo succeeds for Aspirin SMILES");
        Check(aspirinInfo!.Formula == "C9H8O4", "Aspirin formula is C9H8O4, got: " + aspirinInfo.Formula);
        Check(Math.Abs(aspirinInfo.MolecularWeight - 180.16) < 0.1, "Aspirin MW ~ 180.16, got: " + aspirinInfo.MolecularWeight);
        Check(Math.Abs(aspirinInfo.ExactMass - 180.0423) < 0.01, "Aspirin Exact Mass ~ 180.0423, got: " + aspirinInfo.ExactMass);
        Check(!string.IsNullOrEmpty(aspirinInfo.Smiles), "Aspirin canonical SMILES non-empty");
        Check(!string.IsNullOrEmpty(aspirinInfo.InChI) && aspirinInfo.InChI.StartsWith("InChI=1S/C9H8O4"), "Aspirin info includes InChI");
        Check(!string.IsNullOrEmpty(aspirinInfo.InChIKey), "Aspirin info includes InChIKey");

        var ketInfo = MoleculeRenderer.TryGetMoleculeInfo(RenderingUnitTests.SampleKet);
        Check(ketInfo != null && ketInfo.Formula == "CH4O", "TryGetMoleculeInfo succeeds for methanol KET");

        var mwVal = Functions.SmilesMW("CCO");
        Check(mwVal is double d && Math.Abs(d - 46.07) < 0.1, "Functions.SmilesMW for CCO ~ 46.07");
        Check(Equals(Functions.SmilesFormula("CCO"), "C2H6O"), "Functions.SmilesFormula for CCO is C2H6O");

        var ethanolInchi = MoleculeRenderer.TryConvertToInchi("CCO");
        Check(ethanolInchi != null && ethanolInchi.StartsWith("InChI=1S/C2H6O", StringComparison.OrdinalIgnoreCase), "InChI for CCO starts with InChI=1S/C2H6O");
        var ethanolInchiKey = MoleculeRenderer.TryConvertToInchiKey("CCO");
        Check(string.Equals(ethanolInchiKey, "LFQSCWFLJHTTHZ-UHFFFAOYSA-N", StringComparison.OrdinalIgnoreCase), "InChIKey for CCO is LFQSCWFLJHTTHZ-UHFFFAOYSA-N");

        Check(Functions.SmilesInchi("CCO") is string s && s.StartsWith("InChI=1S/C2H6O"), "Functions.SmilesInchi works for CCO");
        Check(Equals(Functions.SmilesInchiKey("CCO"), "LFQSCWFLJHTTHZ-UHFFFAOYSA-N"), "Functions.SmilesInchiKey works for CCO");

        Check(MoleculeRenderer.TryConvertToSmiles(ethanolInchi!) == "CCO", "TryConvertToSmiles converts ethanol InChI to CCO");
        Check(Equals(Functions.InchiFormula(ethanolInchi!), "C2H6O"), "Functions.InchiFormula for ethanol InChI is C2H6O");
        Check(Equals(Functions.InchiMW(ethanolInchi!), 46.07), "Functions.InchiMW for ethanol InChI is 46.07");
        Check(Equals(Functions.InchiKey(ethanolInchi!), "LFQSCWFLJHTTHZ-UHFFFAOYSA-N"), "Functions.InchiKey from InChI string");
        Check(Equals(Functions.InchiSmiles(ethanolInchi!), "CCO"), "Functions.InchiSmiles from InChI string");

        var pureLocalInfo = MoleculeRenderer.TryGetMoleculeInfo("CCO", resolvePubChem: false);
        Check(pureLocalInfo != null && pureLocalInfo.Formula == "C2H6O" && pureLocalInfo.MolecularWeight > 46.0 && pureLocalInfo.CompoundName == "",
            "TryGetMoleculeInfo with resolvePubChem: false computes local properties without PubChem");
        Check(ReferenceEquals(
            MoleculeRenderer.TryGetMoleculeInfo("CC(=O)Oc1ccccc1C(=O)O", resolvePubChem: false),
            MoleculeRenderer.TryGetMoleculeInfo("CC(=O)Oc1ccccc1C(=O)O", resolvePubChem: false)),
            "TryGetMoleculeInfo _infoCache returns identical cached instance for resolvePubChem=false");

        PassSuite("Unit: Core Molecular Properties, InChI/InChIKey & Thread-Safe Cache");
    }

    private static void RunPhysicochemicalAndRuleOf5Tests()
    {
        var aspirinInfo = MoleculeRenderer.TryGetMoleculeInfo("CC(=O)Oc1ccccc1C(=O)O")!;
        var ethanolInchi = MoleculeRenderer.TryConvertToInchi("CCO")!;

        Check(aspirinInfo.LogP > 0.5 && aspirinInfo.LogP < 2.5, $"Aspirin LogP in expected range, got: {aspirinInfo.LogP}");
        Check(aspirinInfo.Tpsa > 50 && aspirinInfo.Tpsa < 80, $"Aspirin TPSA in expected range, got: {aspirinInfo.Tpsa}");

        Check(Functions.SmilesExactMass("CCO") is double em1 && Math.Abs(em1 - 46.0419) < 0.01, "SMILES2EXACTMASS for CCO");
        Check(Functions.InchiExactMass(ethanolInchi) is double em2 && Math.Abs(em2 - 46.0419) < 0.01, "INCHI2EXACTMASS for ethanol");
        Check(Functions.KetExactMass(RenderingUnitTests.SampleKet) is double em3 && Math.Abs(em3 - 32.0262) < 0.01, "KET2EXACTMASS for methanol KET");
        Check(Functions.SmilesLogP("CC(=O)Oc1ccccc1C(=O)O") is double lp && lp > 0.5, "SMILES2LOGP for Aspirin");
        Check(Functions.SmilesTpsa("CC(=O)Oc1ccccc1C(=O)O") is double tpsa && tpsa > 50, "SMILES2TPSA for Aspirin");

        // Regression: negative LogP molecules must not be swallowed to 0.00 by Indigo.Net checkResult(double) bug
        Check(Functions.SmilesLogP("OC[C@H]1OC(O)[C@H](O)[C@@H](O)[C@@H]1O") is double lpGlu && Math.Abs(lpGlu - (-3.22)) < 0.05,
            $"SMILES2LOGP for Glucose ~ -3.22, got: {Functions.SmilesLogP("OC[C@H]1OC(O)[C@H](O)[C@@H](O)[C@@H]1O")}");
        Check(Functions.SmilesLogP("CN(C)C(=N)NC(=N)N") is double lpMet && Math.Abs(lpMet - (-1.03)) < 0.05,
            $"SMILES2LOGP for Metformin ~ -1.03, got: {Functions.SmilesLogP("CN(C)C(=N)NC(=N)N")}");
        Check(Functions.SmilesLogP("Cn1cnc2c1c(=O)n(C)c(=O)n2C") is double lpCaf && Math.Abs(lpCaf - (-1.03)) < 0.05,
            $"SMILES2LOGP for Caffeine ~ -1.03, got: {Functions.SmilesLogP("Cn1cnc2c1c(=O)n(C)c(=O)n2C")}");
        Check(Functions.SmilesLogP("OC(=O)CC(O)(C(=O)O)CC(=O)O") is double lpCit && Math.Abs(lpCit - (-1.25)) < 0.05,
            $"SMILES2LOGP for Citric Acid ~ -1.25, got: {Functions.SmilesLogP("OC(=O)CC(O)(C(=O)O)CC(=O)O")}");
        Check(Functions.SmilesLogP("NCC(=O)O") is double lpGly && Math.Abs(lpGly - (-0.97)) < 0.05,
            $"SMILES2LOGP for Glycine ~ -0.97, got: {Functions.SmilesLogP("NCC(=O)O")}");
        Check(Functions.SmilesLogP("CO") is double lpMeoh && Math.Abs(lpMeoh - (-0.39)) < 0.05,
            $"SMILES2LOGP for Methanol ~ -0.39, got: {Functions.SmilesLogP("CO")}");
        Check(Functions.SmilesLogP("O") is double lpWater && Math.Abs(lpWater - (-0.44)) < 0.05,
            $"SMILES2LOGP for Water ~ -0.44, got: {Functions.SmilesLogP("O")}");
        Check(Functions.SmilesLogP("CC(=O)Oc1ccccc1C(=O)O") is double lpAsp && Math.Abs(lpAsp - 1.31) < 0.05,
            $"SMILES2LOGP for Aspirin ~ 1.31, got: {Functions.SmilesLogP("CC(=O)Oc1ccccc1C(=O)O")}");

        // Regression: aromatic SMILES, Kekule SMILES, InChI, and KET must return identical descriptors
        const string cafArom = "Cn1cnc2c1c(=O)n(C)c(=O)n2C";
        const string cafKek = "CN1C=NC2=C1C(=O)N(C(=O)N2C)C";
        var cafInchi = MoleculeRenderer.TryConvertToInchi(cafArom)!;
        var cafKet = MoleculeRenderer.TryConvertToKet(cafArom)!;
        foreach (var cafInput in new[] { cafArom, cafKek, cafInchi, cafKet })
        {
            Check(Functions.SmilesLogP(cafInput) is double clp && Math.Abs(clp - (-1.03)) < 0.05,
                $"Caffeine LogP consistent (~ -1.03) across representations, got: {Functions.SmilesLogP(cafInput)}");
            Check(Functions.SmilesTpsa(cafInput) is double ctpsa && Math.Abs(ctpsa - 61.82) < 0.05,
                $"Caffeine TPSA consistent (~ 61.82) across representations, got: {Functions.SmilesTpsa(cafInput)}");
        }
        Check(Functions.SmilesLogP("c1ccncc1") is double lpPyr1 && Functions.SmilesLogP("C1=CC=NC=C1") is double lpPyr2 &&
              Math.Abs(lpPyr1 - lpPyr2) < 0.01 && Math.Abs(lpPyr1 - 1.08) < 0.05,
            "Pyridine LogP consistent between aromatic and Kekule SMILES");
        Check(Functions.SmilesTpsa("c1ccncc1") is double tpsaPyr1 && Functions.SmilesTpsa("C1=CC=NC=C1") is double tpsaPyr2 &&
              Math.Abs(tpsaPyr1 - tpsaPyr2) < 0.01 && Math.Abs(tpsaPyr1 - 12.89) < 0.05,
            "Pyridine TPSA consistent between aromatic and Kekule SMILES");

        // Regression: invalid InChI / KET / SMILES conversion guardrails & strict INCHI2* / KET2* prefix enforcement
        Check(Equals(Functions.InchiSmiles("hello world"), ExcelError.ExcelErrorValue), "INCHI2SMILES rejects non-InChI input");
        Check(Equals(Functions.InchiMW("CCO"), ExcelError.ExcelErrorValue), "INCHI2MW rejects non-InChI SMILES");
        Check(Equals(Functions.InchiFormula("CCO"), ExcelError.ExcelErrorValue), "INCHI2FORMULA rejects non-InChI SMILES");
        Check(Equals(Functions.InchiExactMass("CCO"), ExcelError.ExcelErrorValue), "INCHI2EXACTMASS rejects non-InChI SMILES");
        Check(Equals(Functions.InchiKey("CCO"), ExcelError.ExcelErrorValue), "INCHI2INCHIKEY rejects non-InChI SMILES");
        Check(Equals(Functions.InchiName("CCO"), ExcelError.ExcelErrorValue), "INCHI2NAME rejects non-InChI SMILES");
        Check(Equals(Functions.InchiLogP("CCO"), ExcelError.ExcelErrorValue), "INCHI2LOGP rejects non-InChI SMILES");
        Check(Equals(Functions.InchiTpsa("CCO"), ExcelError.ExcelErrorValue), "INCHI2TPSA rejects non-InChI SMILES");
        Check(Equals(Functions.InchiHeavy("CCO"), ExcelError.ExcelErrorValue), "INCHI2HEAVY rejects non-InChI SMILES");
        Check(Equals(Functions.InchiRings("CCO"), ExcelError.ExcelErrorValue), "INCHI2RINGS rejects non-InChI SMILES");
        Check(Equals(Functions.InchiAromatic("CCO"), ExcelError.ExcelErrorValue), "INCHI2AROMATIC rejects non-InChI SMILES");
        Check(Equals(Functions.InchiFsp3("CCO"), ExcelError.ExcelErrorValue), "INCHI2FSP3 rejects non-InChI SMILES");
        Check(Equals(Functions.InchiMatch("CCO", "CO"), ExcelError.ExcelErrorValue), "INCHI.MATCH rejects non-InChI target");

        Check(Equals(Functions.KetSmiles("CCO"), ExcelError.ExcelErrorValue), "KET2SMILES rejects non-KET SMILES");
        Check(Equals(Functions.KetMW("CCO"), ExcelError.ExcelErrorValue), "KET2MW rejects non-KET SMILES");
        Check(Equals(Functions.KetFormula("CCO"), ExcelError.ExcelErrorValue), "KET2FORMULA rejects non-KET SMILES");
        Check(Equals(Functions.KetExactMass("CCO"), ExcelError.ExcelErrorValue), "KET2EXACTMASS rejects non-KET SMILES");
        Check(Equals(Functions.KetInchi("CCO"), ExcelError.ExcelErrorValue), "KET2INCHI rejects non-KET SMILES");
        Check(Equals(Functions.KetInchiKey("CCO"), ExcelError.ExcelErrorValue), "KET2INCHIKEY rejects non-KET SMILES");
        Check(Equals(Functions.KetName("CCO"), ExcelError.ExcelErrorValue), "KET2NAME rejects non-KET SMILES");
        Check(Equals(Functions.KetLogP("CCO"), ExcelError.ExcelErrorValue), "KET2LOGP rejects non-KET SMILES");
        Check(Equals(Functions.KetTpsa("CCO"), ExcelError.ExcelErrorValue), "KET2TPSA rejects non-KET SMILES");
        Check(Equals(Functions.KetHeavy("CCO"), ExcelError.ExcelErrorValue), "KET2HEAVY rejects non-KET SMILES");
        Check(Equals(Functions.KetRings("CCO"), ExcelError.ExcelErrorValue), "KET2RINGS rejects non-KET SMILES");
        Check(Equals(Functions.KetAromatic("CCO"), ExcelError.ExcelErrorValue), "KET2AROMATIC rejects non-KET SMILES");
        Check(Equals(Functions.KetFsp3("CCO"), ExcelError.ExcelErrorValue), "KET2FSP3 rejects non-KET SMILES");
        Check(Equals(Functions.KetMatch("CCO", "c1ccccc1"), ExcelError.ExcelErrorValue), "KET.MATCH rejects non-KET target");

        Check(MoleculeRenderer.TryConvertToSmiles("hello world") == null, "TryConvertToSmiles returns null for invalid SMILES");
        Check(Equals(Functions.SmilesInchi("InChI=GARBAGE_INVALID"), ExcelError.ExcelErrorNA), "SMILES2INCHI rejects malformed InChI=");
        Check(Equals(Functions.InchiKey("1S/INVALID_INCHI"), ExcelError.ExcelErrorNA), "INCHI2INCHIKEY rejects malformed 1S/ InChI");

        // Regression: half-up MW rounding (MidpointRounding.AwayFromZero)
        Check(Equals(Functions.SmilesMW("O"), 18.02), $"SMILES2MW for Water == 18.02, got: {Functions.SmilesMW("O")}");
        Check(Equals(Functions.SmilesMW("C(=O)O"), 46.03), $"SMILES2MW for Formic Acid == 46.03, got: {Functions.SmilesMW("C(=O)O")}");

        // Regression: dummy / R-group atoms rejected from quantitative descriptors (#N/A)
        Check(Equals(Functions.SmilesMW("*c1ccccc1"), ExcelError.ExcelErrorNA), "SMILES2MW rejects R-group *c1ccccc1");
        Check(Equals(Functions.SmilesFormula("*c1ccccc1"), ExcelError.ExcelErrorNA), "SMILES2FORMULA rejects R-group *c1ccccc1");
        Check(Equals(Functions.SmilesLogP("*c1ccccc1"), ExcelError.ExcelErrorNA), "SMILES2LOGP rejects R-group *c1ccccc1");
        Check(Equals(Functions.SmilesTpsa("*"), ExcelError.ExcelErrorNA), "SMILES2TPSA rejects dummy *");
        Check(Equals(Functions.SmilesHeavy("*"), ExcelError.ExcelErrorNA), "SMILES2HEAVY rejects dummy *");

        // Regression: TPSA include_sp optional parameter (Ertl N/O default vs PubChem S/P) + INCHI2*/KET2* parity
        const string amoxicillin = "CC1([C@@H](N2[C@H](S1)[C@@H](C2=O)NC(=O)[C@@H](c3ccc(O)cc3)N)C(=O)O)C";
        var amoxInchi = MoleculeRenderer.TryConvertToInchi(amoxicillin)!;
        var amoxKet = MoleculeRenderer.TryConvertToKet(amoxicillin)!;

        Check(Functions.SmilesTpsa(amoxicillin) is double tpsaNoSp && Math.Abs(tpsaNoSp - 132.96) < 0.05,
            $"SMILES2TPSA default (N/O only) for Amoxicillin ~ 132.96, got: {Functions.SmilesTpsa(amoxicillin)}");
        Check(Functions.SmilesTpsa(amoxicillin, true) is double tpsaBoolSp && Math.Abs(tpsaBoolSp - 158.26) < 0.05,
            $"SMILES2TPSA(include_sp=true) for Amoxicillin ~ 158.26, got: {Functions.SmilesTpsa(amoxicillin, true)}");
        Check(Functions.SmilesTpsa(amoxicillin, "sp") is double tpsaStrSp && Math.Abs(tpsaStrSp - 158.26) < 0.05,
            $"SMILES2TPSA(include_sp=\"sp\") for Amoxicillin ~ 158.26, got: {Functions.SmilesTpsa(amoxicillin, "sp")}");

        // Full KET2* parity against Amoxicillin SMILES (lossless 2D graph round-trip)
        Check(Equals(Functions.KetSmiles(amoxKet), MoleculeRenderer.TryConvertToSmiles(amoxKet)), "KET2SMILES converts Amoxicillin KET cleanly");
        Check(Equals(Functions.KetLogP(amoxKet), Functions.SmilesLogP(amoxicillin)), "KET2LOGP matches SMILES2LOGP on Amoxicillin");
        Check(Equals(Functions.KetTpsa(amoxKet), Functions.SmilesTpsa(amoxicillin)), "KET2TPSA matches SMILES2TPSA (N/O)");
        Check(Equals(Functions.KetTpsa(amoxKet, "sp"), Functions.SmilesTpsa(amoxicillin, "sp")), "KET2TPSA(include_sp=\"sp\") matches SMILES2TPSA");
        Check(Equals(Functions.KetHeavy(amoxKet), Functions.SmilesHeavy(amoxicillin)), "KET2HEAVY matches SMILES2HEAVY");
        Check(Equals(Functions.KetRings(amoxKet), Functions.SmilesRings(amoxicillin)), "KET2RINGS matches SMILES2RINGS");
        Check(Equals(Functions.KetAromatic(amoxKet), Functions.SmilesAromatic(amoxicillin)), "KET2AROMATIC matches SMILES2AROMATIC");
        Check(Equals(Functions.KetFsp3(amoxKet), Functions.SmilesFsp3(amoxicillin)), "KET2FSP3 matches SMILES2FSP3");
        Check(Equals(Functions.KetMatch(amoxKet, "c1ccc(O)cc1"), true), "KET.MATCH finds phenol ring in Amoxicillin KET");

        // Full INCHI2* parity against Ibuprofen (non-tautomeric chiral aromatic carboxylic acid) + Amoxicillin S/P TPSA toggle
        const string ibuprofen = "CC(C)Cc1ccc(cc1)[C@@H](C)C(=O)O";
        var ibuInchi = MoleculeRenderer.TryConvertToInchi(ibuprofen)!;
        Check(Equals(Functions.InchiLogP(ibuInchi), Functions.SmilesLogP(ibuprofen)), "INCHI2LOGP matches SMILES2LOGP on Ibuprofen");
        Check(Equals(Functions.InchiTpsa(ibuInchi), Functions.SmilesTpsa(ibuprofen)), "INCHI2TPSA matches SMILES2TPSA on Ibuprofen");
        Check(Equals(Functions.InchiHeavy(ibuInchi), Functions.SmilesHeavy(ibuprofen)), "INCHI2HEAVY matches SMILES2HEAVY on Ibuprofen");
        Check(Equals(Functions.InchiRings(ibuInchi), Functions.SmilesRings(ibuprofen)), "INCHI2RINGS matches SMILES2RINGS on Ibuprofen");
        Check(Equals(Functions.InchiAromatic(ibuInchi), Functions.SmilesAromatic(ibuprofen)), "INCHI2AROMATIC matches SMILES2AROMATIC on Ibuprofen");
        Check(Equals(Functions.InchiFsp3(ibuInchi), Functions.SmilesFsp3(ibuprofen)), "INCHI2FSP3 matches SMILES2FSP3 on Ibuprofen");
        Check(Equals(Functions.InchiMatch(ibuInchi, "c1ccccc1"), true), "INCHI.MATCH finds benzene ring in Ibuprofen InChI");
        Check(Equals(Functions.InchiLogP(amoxInchi), Functions.SmilesLogP(amoxInchi)), "INCHI2LOGP is identical to SMILES2LOGP(amoxInchi)");
        Check(Functions.InchiTpsa(amoxInchi, true) is double iSp && Functions.InchiTpsa(amoxInchi) is double iNoSp && iSp > iNoSp + 20.0,
            "INCHI2TPSA(include_sp=true) adds sulfur PSA contribution over default Ertl N/O");
        Check(Equals(Functions.InchiMatch(ethanolInchi, "c1ccccc1"), false), "INCHI.MATCH returns false for benzene in Ethanol InChI");
        Check(Equals(Functions.KetMatch(RenderingUnitTests.SampleKet, "c1ccccc1"), false), "KET.MATCH returns false for benzene in Methanol KET");

        PassSuite("Unit: Physicochemical Descriptors & Cross-Notation Parity (SMILES / InChI / KET)");
    }

    private static void RunStructuralComplexityTests()
    {
        var aspirinInfo = MoleculeRenderer.TryGetMoleculeInfo("CC(=O)Oc1ccccc1C(=O)O")!;
        Check(aspirinInfo.HeavyAtomCount == 13, $"Aspirin HeavyAtomCount == 13, got: {aspirinInfo.HeavyAtomCount}");
        Check(aspirinInfo.RingCount == 1, $"Aspirin RingCount == 1, got: {aspirinInfo.RingCount}");
        Check(aspirinInfo.AromaticRingCount == 1, $"Aspirin AromaticRingCount == 1, got: {aspirinInfo.AromaticRingCount}");
        Check(Math.Abs(aspirinInfo.Fsp3 - 0.11) < 0.01, $"Aspirin Fsp3 == 0.11, got: {aspirinInfo.Fsp3}");

        Check(Equals(Functions.SmilesHeavy("CC(=O)Oc1ccccc1C(=O)O"), 13), "SMILES2HEAVY for Aspirin == 13");
        Check(Equals(Functions.SmilesRings("CN1C=NC2=C1C(=O)N(C(=O)N2C)C"), 2), "SMILES2RINGS for Caffeine == 2");
        Check(Equals(Functions.SmilesAromatic("CN1C=NC2=C1C(=O)N(C(=O)N2C)C"), 2), "SMILES2AROMATIC for Caffeine == 2");
        Check(Equals(Functions.SmilesRings("C1CCCCC1"), 1), "SMILES2RINGS for Cyclohexane == 1");
        Check(Equals(Functions.SmilesAromatic("C1CCCCC1"), 0), "SMILES2AROMATIC for Cyclohexane == 0");
        Check(Functions.SmilesFsp3("C1CCCCC1") is double fsp3Cyc && Math.Abs(fsp3Cyc - 1.0) < 0.01, "SMILES2FSP3 for Cyclohexane == 1.0");
        Check(Functions.SmilesFsp3("CC(=O)Oc1ccccc1C(=O)O") is double fsp3Asp && Math.Abs(fsp3Asp - 0.11) < 0.01, "SMILES2FSP3 for Aspirin == 0.11");
        Check(Equals(Functions.SmilesHeavy("[H][H]"), 0), "SMILES2HEAVY for H2 == 0");

        var molstarInfo = MoleculeActionHelper.TryResolveMolstarInfo("MOLSTAR:{\"id\":\"test\",\"pdb\":\"1CRN\"}", "Mol* 3D: 1CRN", false);
        Check(molstarInfo != null && molstarInfo.IsMacromolecule && !molstarInfo.FormatMultiLine().Contains("LogP") && !molstarInfo.FormatMarkdownTable().Contains("LogP"),
            "Molstar Macromolecule MoleculeInfo omits small-molecule descriptors");

        PassSuite("Unit: Structural Complexity Descriptors (Heavy, Rings, Aromatic, Fsp3 & Macromolecule Guard)");
    }

    private static void RunSubstructureAndSimilarityTests()
    {
        Check(Equals(Functions.SmilesMatch("CC(=O)Oc1ccccc1C(=O)O", "c1ccccc1"), true), "SMILES.MATCH finds benzene in Aspirin");
        Check(Equals(Functions.SmilesMatch("CCO", "c1ccccc1"), false), "SMILES.MATCH returns false for benzene in Ethanol");
        Check(Equals(Functions.SmilesMatch("c1ccccc1", "C"), false), "SMILES.MATCH treats SMARTS 'C' as aliphatic only (false for benzene)");
        Check(Equals(Functions.SmilesMatch("c1ccncc1", "N"), false), "SMILES.MATCH treats SMARTS 'N' as aliphatic only (false for pyridine)");
        Check(Equals(Functions.SmilesMatch("Oc1ccccc1", "C-O"), false), "SMILES.MATCH treats SMARTS 'C-O' as aliphatic (false for phenol)");
        Check(Equals(Functions.SmilesMatch("Oc1ccccc1", "c-O"), true), "SMILES.MATCH matches aromatic phenol with SMARTS 'c-O'");
        Check(Equals(Functions.SmilesMatch("CC(=O)Oc1ccccc1C(=O)O", "C(=O)[OH]"), true), "SMILES.MATCH matches carboxylic acid SMARTS in Aspirin");
        Check(Equals(Functions.SmilesMatch("CCOC(=O)C", "C(=O)[OH]"), false), "SMILES.MATCH distinguishes ester from carboxylic acid SMARTS");
        Check(Equals(Functions.SmilesMatch("C1=CC=CC=C1", "c1ccccc1"), true), "SMILES.MATCH normalizes Kekule target to match aromatic SMARTS");

        var hlOpts = Functions.ParseOptions(null, "hl:c1ccccc1,num", null);
        Check(hlOpts.HighlightQuery == "c1ccccc1" && hlOpts.AtomNumbers == true, "ParseOptions extracts hl:c1ccccc1 and num");
        var hlPng = MoleculeRenderer.Render("CC(=O)Oc1ccccc1C(=O)O", hlOpts);
        Check(hlPng != null && hlPng.Length > 200, "Substructure highlighted PNG rendered");

        PassSuite("Unit: Substructure Search (SMILES.MATCH / INCHI.MATCH / KET.MATCH) & Highlighting");
    }

    private static void RunPubChemAndNameResolutionTests()
    {
        var resolvedEthanol = PubChemResolver.ResolveCompound("CCO");
        Check(string.Equals(resolvedEthanol.CompoundName, "Ethanol", StringComparison.OrdinalIgnoreCase), "PubChem resolves CCO to Ethanol");
        Check(string.Equals(PubChemResolver.ResolveCompound("c1ccccc1").CompoundName, "Benzene", StringComparison.OrdinalIgnoreCase), "Case-sensitive lookup resolves c1ccccc1 to Benzene");
        Check(string.Equals(PubChemResolver.ResolveCompound("C1CCCCC1").CompoundName, "Cyclohexane", StringComparison.OrdinalIgnoreCase), "Case-sensitive lookup resolves C1CCCCC1 to Cyclohexane");
        Check(string.Equals(Functions.SmilesName("c1ccccc1")?.ToString(), "Benzene", StringComparison.OrdinalIgnoreCase), "SMILES2NAME resolves c1ccccc1 to Benzene");
        Check(string.Equals(Functions.SmilesName("CC(=O)Oc1ccccc1C(=O)O")?.ToString(), "Aspirin", StringComparison.OrdinalIgnoreCase), "SMILES2NAME hits offline canonical Aspirin key");
        Check(!string.Equals(PubChemResolver.ResolveCompound("C1CCNCC1", timeoutMs: 100).CompoundName, "Pyridine", StringComparison.OrdinalIgnoreCase), "Piperidine C1CCNCC1 does not collide with Pyridine");
        Check(!string.Equals(PubChemResolver.ResolveCompound("c1ccoc1", timeoutMs: 100).CompoundName, "Tetrahydrofuran", StringComparison.OrdinalIgnoreCase), "Furan c1ccoc1 does not collide with Tetrahydrofuran");
        Check(Functions.NameSmiles("Aspirin") is string aspSmiles && !string.IsNullOrEmpty(aspSmiles), "NAME2SMILES resolves Aspirin");
        Check(Functions.NameSmiles("50-78-2") is string casSmiles && !string.IsNullOrEmpty(casSmiles), "NAME2SMILES resolves CAS 50-78-2");
        Check(Functions.NameSmiles("Caffeine") is string cafSmiles && !string.IsNullOrEmpty(cafSmiles), "NAME2SMILES resolves Caffeine");

        PassSuite("Unit: PubChem Compound Name & Reverse Name/CAS-to-SMILES Resolution");
    }

    private static void RunMultiFormatSummariesAndSdfTests()
    {
        var ccoInfo = MoleculeRenderer.TryGetMoleculeInfo("CCO", resolvePubChem: true)!;
        string ccoSummary = ccoInfo.FormatClipboardSummary();
        Check(!string.IsNullOrWhiteSpace(ccoSummary), "CCO summary formatted");

        string ccoTsv = ccoInfo.FormatClipboardSummary(MoleculeSummaryFormat.ExcelTsvRow);
        Check(ccoTsv.Contains("C2H6O") && ccoTsv.Contains("\t"), "CCO TSV row contains formula and tabs");

        string ccoTsvHeader = ccoInfo.FormatClipboardSummary(MoleculeSummaryFormat.ExcelTsvWithHeader);
        Check(ccoTsvHeader.Contains("MW (g/mol)") && ccoTsvHeader.Contains("\r\n") && !ccoTsvHeader.Contains("Chiral"), "CCO TSV with header has 12 columns and no Chiral Centers");

        string ccoMulti = ccoInfo.FormatClipboardSummary(MoleculeSummaryFormat.MultiLine);
        Check(ccoMulti.Contains("Molecular Weight: 46.07 g/mol") && ccoMulti.Contains("\r\n"), "CCO MultiLine formatted with line breaks");

        string ccoMd = ccoInfo.FormatClipboardSummary(MoleculeSummaryFormat.MarkdownTable);
        Check(ccoMd.StartsWith("| Property | Value |") && ccoMd.Contains("| **Formula** | C2H6O |"), "CCO Markdown table formatted correctly");

        string ccoJson = ccoInfo.FormatClipboardSummary(MoleculeSummaryFormat.Json);
        Check(ccoJson.Contains("\"formula\": \"C2H6O\"") && ccoJson.Contains("\"molecularWeight\": 46.07") && !ccoJson.Contains("chiralCenters"), "CCO JSON formatted correctly");

        string ccoHtml = ccoInfo.FormatHtmlTable();
        Check(ccoHtml.StartsWith("<table") && ccoHtml.Contains("46.07 g/mol"), "CCO HTML table formatted correctly");
        Check(!ccoHtml.Contains("&amp;#"), "No double-encoded HTML entities in ccoHtml");
        Check(ccoHtml.Contains("Å²") || ccoHtml.Contains("&#197;&#178;"), "TPSA unit present in ccoHtml");

        string ccoHorizHtml = ccoInfo.FormatHtmlHorizontalTable(includeHeader: true);
        Check(ccoHorizHtml.StartsWith("<table") && ccoHorizHtml.Contains(">Name</th>") && ccoHorizHtml.Contains("46.07"), "Horizontal HTML table with header");
        Check(!ccoHorizHtml.Contains("&amp;#"), "No double-encoded HTML entities in ccoHorizHtml");

        string ccoHorizNoHeader = ccoInfo.FormatHtmlHorizontalTable(includeHeader: false);
        Check(!ccoHorizNoHeader.Contains("<th"), "Horizontal HTML table without header has no th tags");

        Check(ccoInfo.FormatHtmlForSummary(MoleculeSummaryFormat.ExcelTsvWithHeader) != null, "FormatHtmlForSummary for ExcelTsvWithHeader");
        Check(ccoInfo.FormatHtmlForSummary(MoleculeSummaryFormat.MultiLine) != null, "FormatHtmlForSummary for MultiLine");

        var dirtyInfo = new MoleculeInfo { CompoundName = "Special\tSubstance\r\nName", Formula = "C6\tH6", Smiles = "c1ccccc1" };
        string dirtyTsv = dirtyInfo.FormatExcelTsv(includeHeader: false);
        Check(!dirtyTsv.Contains("\r") && !dirtyTsv.Contains("\n"), "Dirty TSV row contains no line breaks");
        Check(dirtyTsv.Split('\t').Length == 12, $"Dirty TSV row maintains exactly 12 columns, got: {dirtyTsv.Split('\t').Length}");

        var sdfInput = new[]
        {
            new SdfRecord { Smiles = "CCO", Name = "Ethanol", Properties = new Dictionary<string, string> { ["Source"] = "UnitTest" } },
            new SdfRecord { Smiles = "CC(=O)Oc1ccccc1C(=O)O", Name = "Aspirin" }
        };
        var sdfBytes = SdfManager.WriteSdfBytes(sdfInput);
        Check(sdfBytes != null && sdfBytes.Length > 100, "SdfManager.WriteSdfBytes produces non-empty SDF bytes");
        var sdfRecords = SdfManager.ReadSdfBytes(sdfBytes!);
        Check(sdfRecords.Count == 2, $"SdfManager.ReadSdfBytes read 2 records, got: {sdfRecords.Count}");
        Check(sdfRecords[0].Formula == "C2H6O", $"SDF record 0 formula is C2H6O, got: {sdfRecords[0].Formula}");
        Check(sdfRecords[1].Formula == "C9H8O4", $"SDF record 1 formula is C9H8O4, got: {sdfRecords[1].Formula}");

        PassSuite("Unit: Multi-Format Summaries (12-Col TSV, Markdown, JSON, CF_HTML) & SDF File Parsing");
    }
}
