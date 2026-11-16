using System;
using System.Collections.Generic;
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
        Check(aspirinInfo.HBondDonors == 1, $"Aspirin HBD == 1, got: {aspirinInfo.HBondDonors}");
        Check(aspirinInfo.HBondAcceptors >= 3, $"Aspirin HBA >= 3, got: {aspirinInfo.HBondAcceptors}");
        Check(aspirinInfo.RotatableBonds >= 2, $"Aspirin RotBonds >= 2, got: {aspirinInfo.RotatableBonds}");
        Check(aspirinInfo.PassesLipinski == true, "Aspirin passes Lipinski Rule of 5");

        Check(Functions.SmilesExactMass("CCO") is double em1 && Math.Abs(em1 - 46.0419) < 0.01, "SMILES2EXACTMASS for CCO");
        Check(Functions.InchiExactMass(ethanolInchi) is double em2 && Math.Abs(em2 - 46.0419) < 0.01, "INCHI2EXACTMASS for ethanol");
        Check(Functions.KetExactMass(RenderingUnitTests.SampleKet) is double em3 && Math.Abs(em3 - 32.0262) < 0.01, "KET2EXACTMASS for methanol KET");
        Check(Functions.SmilesLogP("CC(=O)Oc1ccccc1C(=O)O") is double lp && lp > 0.5, "SMILES2LOGP for Aspirin");
        Check(Functions.SmilesTpsa("CC(=O)Oc1ccccc1C(=O)O") is double tpsa && tpsa > 50, "SMILES2TPSA for Aspirin");
        Check(Equals(Functions.SmilesHbd("CC(=O)Oc1ccccc1C(=O)O"), 1), "SMILES2HBD for Aspirin == 1");
        Check(Functions.SmilesHba("CC(=O)Oc1ccccc1C(=O)O") is int hba && hba >= 3, "SMILES2HBA for Aspirin >= 3");
        Check(Functions.SmilesRotBonds("CC(=O)Oc1ccccc1C(=O)O") is int rb && rb >= 2, "SMILES2ROTBONDS for Aspirin >= 2");
        Check(Equals(Functions.SmilesLipinski("CC(=O)Oc1ccccc1C(=O)O"), "PASS"), "SMILES2LIPINSKI for Aspirin == PASS");

        PassSuite("Unit: Physicochemical Descriptors (ExactMass, LogP, TPSA, HBD, HBA, RotBonds, Lipinski)");
    }

    private static void RunStructuralComplexityTests()
    {
        var aspirinInfo = MoleculeRenderer.TryGetMoleculeInfo("CC(=O)Oc1ccccc1C(=O)O")!;
        Check(aspirinInfo.HeavyAtomCount == 13, $"Aspirin HeavyAtomCount == 13, got: {aspirinInfo.HeavyAtomCount}");
        Check(aspirinInfo.RingCount == 1, $"Aspirin RingCount == 1, got: {aspirinInfo.RingCount}");
        Check(aspirinInfo.AromaticRingCount == 1, $"Aspirin AromaticRingCount == 1, got: {aspirinInfo.AromaticRingCount}");
        Check(Math.Abs(aspirinInfo.Fsp3 - 0.11) < 0.01, $"Aspirin Fsp3 == 0.11, got: {aspirinInfo.Fsp3}");
        Check(aspirinInfo.ChiralCenters == 0, $"Aspirin ChiralCenters == 0, got: {aspirinInfo.ChiralCenters}");

        Check(Equals(Functions.SmilesHeavy("CC(=O)Oc1ccccc1C(=O)O"), 13), "SMILES2HEAVY for Aspirin == 13");
        Check(Equals(Functions.SmilesRings("CN1C=NC2=C1C(=O)N(C(=O)N2C)C"), 2), "SMILES2RINGS for Caffeine == 2");
        Check(Equals(Functions.SmilesAromatic("CN1C=NC2=C1C(=O)N(C(=O)N2C)C"), 2), "SMILES2AROMATIC for Caffeine == 2");
        Check(Equals(Functions.SmilesRings("C1CCCCC1"), 1), "SMILES2RINGS for Cyclohexane == 1");
        Check(Equals(Functions.SmilesAromatic("C1CCCCC1"), 0), "SMILES2AROMATIC for Cyclohexane == 0");
        Check(Functions.SmilesFsp3("C1CCCCC1") is double fsp3Cyc && Math.Abs(fsp3Cyc - 1.0) < 0.01, "SMILES2FSP3 for Cyclohexane == 1.0");
        Check(Functions.SmilesFsp3("CC(=O)Oc1ccccc1C(=O)O") is double fsp3Asp && Math.Abs(fsp3Asp - 0.11) < 0.01, "SMILES2FSP3 for Aspirin == 0.11");
        Check(Equals(Functions.SmilesChiral("C[C@H](N)C(=O)O"), 1), "SMILES2CHIRAL for L-Alanine == 1");
        Check(Equals(Functions.SmilesChiral("CCO"), 0), "SMILES2CHIRAL for Ethanol == 0");

        var molstarInfo = MoleculeActionHelper.TryResolveMolstarInfo("MOLSTAR:{\"id\":\"test\",\"pdb\":\"1CRN\"}", "Mol* 3D: 1CRN", false);
        Check(molstarInfo != null && molstarInfo.IsMacromolecule && !molstarInfo.FormatMultiLine().Contains("Lipinski") && !molstarInfo.FormatMarkdownTable().Contains("LogP"),
            "Molstar Macromolecule MoleculeInfo omits small-molecule descriptors");

        PassSuite("Unit: Structural Complexity Descriptors (Heavy, Rings, Aromatic, Fsp3, Chiral & Macromolecule Guard)");
    }

    private static void RunSubstructureAndSimilarityTests()
    {
        Check(Equals(Functions.SmilesMatch("CC(=O)Oc1ccccc1C(=O)O", "c1ccccc1"), true), "SMILES.MATCH finds benzene in Aspirin");
        Check(Equals(Functions.SmilesMatch("CCO", "c1ccccc1"), false), "SMILES.MATCH returns false for benzene in Ethanol");
        Check(Functions.SmilesSimilarity("c1ccccc1", "Cc1ccccc1", "tanimoto") is double sim && sim > 0.3 && sim <= 1.0, "SMILES.SIMILARITY Tanimoto between benzene and toluene");
        Check(Functions.SmilesSimilarity("CCO", "CCO", "tanimoto") is double simSelf && Math.Abs(simSelf - 1.0) < 1e-6, "SMILES.SIMILARITY identical molecule == 1.0");
        Check(Functions.SmilesSimilarity("c1ccccc1", "Cc1ccccc1", "tversky") is double simTv && simTv > 0.3 && simTv <= 1.0, "SMILES.SIMILARITY Tversky");

        var hlOpts = Functions.ParseOptions(null, "hl:c1ccccc1,num", null);
        Check(hlOpts.HighlightQuery == "c1ccccc1" && hlOpts.AtomNumbers == true, "ParseOptions extracts hl:c1ccccc1 and num");
        var hlPng = MoleculeRenderer.Render("CC(=O)Oc1ccccc1C(=O)O", hlOpts);
        Check(hlPng != null && hlPng.Length > 200, "Substructure highlighted PNG rendered");

        PassSuite("Unit: Substructure Search (SMILES.MATCH), Similarity (Tanimoto/Tversky) & Highlighting");
    }

    private static void RunPubChemAndNameResolutionTests()
    {
        var resolvedEthanol = PubChemResolver.ResolveCompound("CCO");
        Check(string.Equals(resolvedEthanol.CompoundName, "Ethanol", StringComparison.OrdinalIgnoreCase), "PubChem resolves CCO to Ethanol");
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
        Check(ccoTsvHeader.Contains("MW (g/mol)") && ccoTsvHeader.Contains("\r\n"), "CCO TSV with header has header row");

        string ccoMulti = ccoInfo.FormatClipboardSummary(MoleculeSummaryFormat.MultiLine);
        Check(ccoMulti.Contains("Molecular Weight: 46.07 g/mol") && ccoMulti.Contains("\r\n"), "CCO MultiLine formatted with line breaks");

        string ccoMd = ccoInfo.FormatClipboardSummary(MoleculeSummaryFormat.MarkdownTable);
        Check(ccoMd.StartsWith("| Property | Value |") && ccoMd.Contains("| **Formula** | C2H6O |"), "CCO Markdown table formatted correctly");

        string ccoJson = ccoInfo.FormatClipboardSummary(MoleculeSummaryFormat.Json);
        Check(ccoJson.Contains("\"formula\": \"C2H6O\"") && ccoJson.Contains("\"molecularWeight\": 46.07"), "CCO JSON formatted correctly");

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
        Check(dirtyTsv.Split('\t').Length == 17, $"Dirty TSV row maintains exactly 17 columns, got: {dirtyTsv.Split('\t').Length}");

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

        PassSuite("Unit: Multi-Format Summaries (17-Col TSV, Markdown, JSON, CF_HTML) & SDF File Round-Trip");
    }
}
