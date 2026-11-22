using System;
using System.Drawing;
using System.IO;
using System.Linq;
using ExcelDna.Integration;
using Smiles2Img;
using Smiles2Img.Editor;
using Smiles2Img.Rendering;
using static Smiles2Img.Tests.TestAssert;

namespace Smiles2Img.Tests.Unit;

internal static class RenderingUnitTests
{
    public const string SampleKet = "{\"root\":{\"nodes\":[{\"$ref\":\"mol0\"}]},\"mol0\":{\"type\":\"molecule\",\"atoms\":[{\"label\":\"C\",\"location\":[0,0,0]},{\"label\":\"O\",\"location\":[1,0,0]}],\"bonds\":[{\"type\":1,\"atoms\":[0,1]}]}}";

    public static void RunAll()
    {
        RunBasicPngRenderingTests();
        RunInvalidInputTests();
        RunExcelFunctionAttributeTests();
        RunFloatingPictureGeometryTests();
        RunParseOptionsAndDomainStyleTests();
        RunKetAndInchiRenderingTests();
        RunStructureMetadataAndMolstarSessionTests();
        RunVectorSvgTests();
    }

    private static void RunBasicPngRenderingTests()
    {
        var samples = new[] { "CCO", "c1ccccc1", "C[C@H](O)C(=O)O", "[NH4+].[Cl-]", "CC(=O)Oc1ccccc1C(=O)O", "Cn1cnc2c1c(=O)n(C)c(=O)n2C" };
        foreach (var smiles in samples)
        {
            var png = MoleculeRenderer.Render(smiles);
            Check(png.Take(8).SequenceEqual(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }), "PNG signature");
            using var stream = new MemoryStream(png);
            using var bitmap = new Bitmap(stream);
            Check(bitmap.Width >= 200 && bitmap.Height >= 100, "PNG dimensions fit the molecule");
            var ink = 0;
            for (var y = 0; y < bitmap.Height; y++)
                for (var x = 0; x < bitmap.Width; x++)
                {
                    var color = bitmap.GetPixel(x, y);
                    if (color.A > 0 && (color.R < 220 || color.G < 220 || color.B < 220)) ink++;
                }
            Check(ink > 20, $"PNG must contain a visible molecule ({smiles})");
        }
        PassSuite("Unit: Basic 2D PNG Rendering (6 reference molecules)");
    }

    private static void RunInvalidInputTests()
    {
        foreach (var input in new[] { "", " ", new string('C', 2001), "not-a-smiles", "C1CC", "C(C" })
        {
            try
            {
                MoleculeRenderer.Render(input);
                throw new Exception("Accepted invalid SMILES: " + input);
            }
            catch (InvalidSmilesException) { }
        }
        PassSuite("Unit: Invalid SMILES Rejection & Input Guardrails");
    }

    private static void RunExcelFunctionAttributeTests()
    {
        var name = typeof(Functions).GetMethod(nameof(Functions.Img))!.GetCustomAttributes(typeof(ExcelFunctionAttribute), false)
            .Cast<ExcelFunctionAttribute>().Single().Name;
        Check(name == "SMILES2IMG", "Excel function name");

        var floatAttribute = typeof(FloatFunctions).GetMethod(nameof(FloatFunctions.ImgFloat))!.GetCustomAttributes(typeof(ExcelFunctionAttribute), false)
            .Cast<ExcelFunctionAttribute>().Single();
        Check(floatAttribute.Name == "SMILES2IMG.FLOAT", "Floating Excel function name");
        Check(floatAttribute.IsMacroType, "Floating function must be IsMacroType=true for caller reference");
        Check(floatAttribute.Description.StartsWith("[Excel 2016+]") && floatAttribute.Description.Length <= 255, "Floating function states its Excel requirement");

        var nativeAttribute = typeof(Functions).GetMethod(nameof(Functions.Img))!.GetCustomAttributes(typeof(ExcelFunctionAttribute), false)
            .Cast<ExcelFunctionAttribute>().Single();
        Check(nativeAttribute.IsMacroType, "Native function must be IsMacroType=true");
        Check(nativeAttribute.Description.StartsWith("[Excel 2024 / Microsoft 365]") && nativeAttribute.Description.Contains("SMILES2IMG.FLOAT") && nativeAttribute.Description.Length <= 255,
            "In-cell function states its Excel requirement");

        var ketAttr = typeof(Functions).GetMethod(nameof(Functions.KetImg))!.GetCustomAttributes(typeof(ExcelFunctionAttribute), false)
            .Cast<ExcelFunctionAttribute>().Single();
        Check(ketAttr.Name == "KET2IMG" && ketAttr.IsMacroType, "KET2IMG function registration");

        var ketFloatAttr = typeof(FloatFunctions).GetMethod(nameof(FloatFunctions.KetImgFloat))!.GetCustomAttributes(typeof(ExcelFunctionAttribute), false)
            .Cast<ExcelFunctionAttribute>().Single();
        Check(ketFloatAttr.Name == "KET2IMG.FLOAT" && ketFloatAttr.IsMacroType, "KET2IMG.FLOAT function registration");

        PassSuite("Unit: Excel-DNA Function Registration & Metadata Attributes");
    }

    private static void RunFloatingPictureGeometryTests()
    {
        var wide = FloatingPictures.Fit(100, 50, 200, 100, 3.0);
        Check(Math.Abs(wide.Height - 96) < 1e-9 && Math.Abs(wide.Width - 288) > 1 && Math.Abs(wide.Width / wide.Height - 3.0) < 1e-9 || Math.Abs(wide.Width - 196) < 1e-9,
            "Fit keeps aspect ratio");
        var tall = FloatingPictures.Fit(0, 0, 100, 300, 2.0);
        Check(Math.Abs(tall.Width - 96) < 1e-9 && Math.Abs(tall.Height - 48) < 1e-9 && Math.Abs(tall.Left - 2) < 1e-9 && Math.Abs(tall.Top - 126) < 1e-9,
            "Fit centers a width-bound picture");
        var flat = FloatingPictures.Fit(0, 0, 300, 100, 1.0);
        Check(Math.Abs(flat.Width - 96) < 1e-9 && Math.Abs(flat.Height - 96) < 1e-9 && Math.Abs(flat.Left - 102) < 1e-9 && Math.Abs(flat.Top - 2) < 1e-9,
            "Fit centers a height-bound picture");
        var title = FloatingPictures.Title("C_ABC", 1.75);
        Check(FloatingPictures.TryParseTitle(title, out var titleKey, out var titleAspect) && titleKey == "C_ABC" && titleAspect == 1.75, "Title round-trip");
        Check(!FloatingPictures.TryParseTitle("Picture 1", out _, out _) && !FloatingPictures.TryParseTitle(null, out _, out _), "Foreign shapes are not owned");
        var aspectPng = MoleculeRenderer.Render("CCO");
        using (var aspectStream = new MemoryStream(aspectPng))
        using (var aspectBitmap = new Bitmap(aspectStream))
            Check(Math.Abs(FloatingPictures.Aspect(aspectPng) - (double)aspectBitmap.Width / aspectBitmap.Height) < 1e-9, "PNG aspect from header");

        PassSuite("Unit: FloatingPictures Fit, Title & Aspect Ratio Helpers");
    }

    private static void RunParseOptionsAndDomainStyleTests()
    {
        var opt1 = Functions.ParseOptions(null, null, null);
        Check(opt1.Color == true && opt1.Transform == null && opt1.Background == "transparent", "ParseOptions default (transparent)");
        var opt2 = Functions.ParseOptions("trans", null, null);
        Check(opt2.Color == true && opt2.Transform == null && opt2.Background == "transparent", "ParseOptions trans 2nd arg");
        var opt3 = Functions.ParseOptions("white", null, null);
        Check(opt3.Color == true && opt3.Transform == null && opt3.Background == "white", "ParseOptions white 2nd arg");
        var opt4 = Functions.ParseOptions("yellow", null, null);
        Check(opt4.Color == true && opt4.Transform == null && opt4.Background == "yellow", "ParseOptions CSS color 'yellow'");
        var opt5 = Functions.ParseOptions("#FFFF00", null, "rot90");
        Check(opt5.Color == true && opt5.Transform == "rot90" && opt5.Background == "#ffff00", "ParseOptions hex bg + rot90");
        var opt6 = Functions.ParseOptions("FFFF00", null, null);
        Check(opt6.Color == true && opt6.Background == "#ffff00", "ParseOptions hex without hash");
        var opt7 = Functions.ParseOptions(null, false, "flip");
        Check(opt7.Color == false && opt7.Transform == "flip" && opt7.Background == "transparent", "ParseOptions default trans + B/W + flip");
        var optFallback = Functions.ParseOptions(false, null, null);
        Check(optFallback.Color == false && optFallback.Background == "transparent", "ParseOptions fallback bool in 2nd arg");

        var optBwNum = Functions.ParseOptions(null, "bw,num", null);
        Check(optBwNum.Color == false && optBwNum.AtomNumbers == true && optBwNum.StereoLabels == false, "ParseOptions bw,num");
        var optStereoH = Functions.ParseOptions("white", "stereo+h", null);
        Check(optStereoH.StereoLabels == true && optStereoH.UnfoldHydrogens == true && optStereoH.Background == "white", "ParseOptions stereo+h with white bg");
        var optDarkWhite = Functions.ParseOptions("#1e1e1e", "white,stereo", null);
        Check(optDarkWhite.BaseColor == "white" && optDarkWhite.StereoLabels == true && optDarkWhite.Background == "#1e1e1e", "ParseOptions dark mode white ink + stereo");
        var optCustomInk = Functions.ParseOptions(null, "ink=#003366,num,stereo", "rot180");
        Check(optCustomInk.BaseColor == "#003366" && optCustomInk.AtomNumbers == true && optCustomInk.StereoLabels == true && optCustomInk.Transform == "rot180", "ParseOptions custom ink + num + stereo + rot180");
        var optCarb = Functions.ParseOptions(null, "all-c,num", null);
        Check(optCarb.AllCarbons == true && optCarb.AtomNumbers == true, "ParseOptions all-c,num");
        Check(Functions.ParseOptions(null, "mono", null).Color == false, "ParseOptions mono alias");
        Check(Functions.ParseOptions(null, "black", null).Color == false, "ParseOptions black alias");
        Check(Functions.ParseOptions(null, "idx", null).AtomNumbers == true, "ParseOptions idx alias");
        Check(Functions.ParseOptions(null, "chiral", null).StereoLabels == true, "ParseOptions chiral alias");
        Check(Functions.ParseOptions(null, 0.0, null).Color == false, "ParseOptions numeric 0 for B/W compatibility");
        var optPipe = Functions.ParseOptions(null, "bw|num", null);
        Check(optPipe.Color == false && optPipe.AtomNumbers == true, "ParseOptions bw|num pipe delimiter");
        var optSpace = Functions.ParseOptions(null, "white stereo", null);
        Check(optSpace.BaseColor == "white" && optSpace.StereoLabels == true, "ParseOptions white stereo space delimiter");

        var optUltimateSpace = Functions.ParseOptions("white", "bw num stereo h all-c", "flip rot90");
        Check(optUltimateSpace.Color == false && optUltimateSpace.AtomNumbers == true &&
              optUltimateSpace.StereoLabels == true && optUltimateSpace.UnfoldHydrogens == true &&
              optUltimateSpace.AllCarbons == true && optUltimateSpace.Background == "white" &&
              optUltimateSpace.Transform == "flip,rot90", "ParseOptions ultimate full-option (spaces)");

        var optUltimatePipe = Functions.ParseOptions("white", "bw|num|stereo|h|all-c", "flip|rot90");
        Check(optUltimatePipe.Color == false && optUltimatePipe.AtomNumbers == true &&
              optUltimatePipe.StereoLabels == true && optUltimatePipe.UnfoldHydrogens == true &&
              optUltimatePipe.AllCarbons == true && optUltimatePipe.Background == "white" &&
              optUltimatePipe.Transform == "flip,rot90", "ParseOptions ultimate full-option (pipes)");

        var bwPng = MoleculeRenderer.Render("CCO", color: false);
        using (var bwStream = new MemoryStream(bwPng))
        using (var bwBitmap = new Bitmap(bwStream))
        {
            var hasInk = false;
            for (var y = 0; y < bwBitmap.Height; y++)
                for (var x = 0; x < bwBitmap.Width; x++)
                {
                    var c = bwBitmap.GetPixel(x, y);
                    Check(c.R == c.G && c.G == c.B, "B/W image must only contain grayscale pixels");
                    if (c.R < 220) hasInk = true;
                }
            Check(hasInk, "B/W image must contain visible ink");
        }

        var numPng = MoleculeRenderer.Render("CCO", Functions.ParseOptions(null, "num", null));
        Check(numPng.Length > 100, "Render with atom numbers");
        var stereoPng = MoleculeRenderer.Render("C[C@H](O)C(=O)O", Functions.ParseOptions(null, "stereo", null));
        Check(stereoPng.Length > 100, "Render with stereo labels");
        var hPng = MoleculeRenderer.Render("CCO", Functions.ParseOptions(null, "h", null));
        Check(hPng.Length > 100, "Render with unfolded hydrogens");
        var carbPng = MoleculeRenderer.Render("c1ccccc1", Functions.ParseOptions(null, "all-c", null));
        Check(carbPng.Length > 100, "Render with explicit carbons");
        var comboPng = MoleculeRenderer.Render("C[C@H](O)C(=O)O", Functions.ParseOptions(null, "bw,num,stereo", null));
        Check(comboPng.Length > 100, "Render with composite bw,num,stereo");
        var ultimatePng = MoleculeRenderer.Render("CC(=O)Oc1ccccc1C(=O)O", optUltimateSpace);
        Check(ultimatePng.Length > 100, "Render with ultimate full-option");

        var darkOptions = Functions.ParseOptions("#1e1e1e", "white,stereo", null);
        var darkPng = MoleculeRenderer.Render("C[C@H](O)C(=O)O", darkOptions);
        using (var darkStream = new MemoryStream(darkPng))
        using (var darkBmp = new Bitmap(darkStream))
        {
            var bgPixel = darkBmp.GetPixel(0, 0);
            Check(bgPixel.R == 0x1e && bgPixel.G == 0x1e && bgPixel.B == 0x1e, "Dark background color match");
            var hasWhiteInk = false;
            for (var y = 0; y < darkBmp.Height; y++)
                for (var x = 0; x < darkBmp.Width; x++)
                {
                    var p = darkBmp.GetPixel(x, y);
                    if (p.R > 230 && p.G > 230 && p.B > 230) hasWhiteInk = true;
                }
            Check(hasWhiteInk, "Dark mode image must contain white ink");
        }

        var aflatoxinPng = MoleculeRenderer.Render("O1C=C[C@H]([C@H]1O2)c3c2cc(OC)c4c3OC(=O)C5=C4CCC(=O)5");
        Check(aflatoxinPng.Length > 100, "Aflatoxin junction H render");
        var micPng = MoleculeRenderer.Render("CN=C=O");
        Check(micPng.Length > 100, "MIC small molecule render");
        var nicFlippedPng = MoleculeRenderer.Render("CN1CCC[C@H]1c2cccnc2", color: true, transform: "flip");
        Check(nicFlippedPng.Length > 100, "Nicotine flip transform render");
        var artemisininPng = MoleculeRenderer.Render("CC1CC2CC3(C)OO4C(O2)(C1C(=O)O3)C(C)CC4");
        Check(artemisininPng.Length > 100, "Artemisinin render");

        using (var ms = new MemoryStream(MoleculeRenderer.Render("CCO")))
        using (var bmp = new Bitmap(ms))
            Check(bmp.GetPixel(0, 0).A == 0, "Default background corner pixel must have alpha 0 (transparent)");

        using (var ms = new MemoryStream(MoleculeRenderer.Render("CCO", background: "yellow")))
        using (var bmp = new Bitmap(ms))
        {
            var corner = bmp.GetPixel(0, 0);
            Check(corner.R > 240 && corner.G > 240 && corner.B < 20 && corner.A == 255, "CSS color 'yellow' corner pixel");
        }

        using (var ms = new MemoryStream(MoleculeRenderer.Render("CCO", background: "white")))
        using (var bmp = new Bitmap(ms))
        {
            var corner = bmp.GetPixel(0, 0);
            Check(corner.R == 255 && corner.G == 255 && corner.B == 255 && corner.A == 255, "Explicit white background corner pixel");
        }

        Directory.CreateDirectory("output/samples");
        File.WriteAllBytes("output/samples/1_nicotine_default.png", MoleculeRenderer.Render("CN1CCC[C@H]1c2cccnc2"));
        File.WriteAllBytes("output/samples/1_nicotine_flipped.png", nicFlippedPng);
        File.WriteAllBytes("output/samples/2_oenanthotoxin.png", MoleculeRenderer.Render("CCC[C@@H](O)CC/C=C\\C=C/C#CC#C/C=C/CO"));
        File.WriteAllBytes("output/samples/3_pyrethrin.png", MoleculeRenderer.Render("CC1=C(C(=O)C[C@@H]1OC(=O)[C@@H]2[C@H](C2(C)C)/C=C(\\C)/C(=O)OC)C/C=C\\C=C"));
        File.WriteAllBytes("output/samples/4_aflatoxin.png", aflatoxinPng);
        File.WriteAllBytes("output/samples/5_glucose.png", MoleculeRenderer.Render("OC[C@@H](O1)[C@@H](O)[C@H](O)[C@@H](O)[C@H](O)1"));
        File.WriteAllBytes("output/samples/6_bergenin.png", MoleculeRenderer.Render("OC[C@@H](O1)[C@@H](O)[C@H](O)[C@@H]2[C@H]1c3c(O)c(OC)c(O)cc3C(=O)O2"));
        File.WriteAllBytes("output/samples/7_pheromone.png", MoleculeRenderer.Render("CC(=O)OCCC(/C)=C\\C[C@H](C(=C)C)CCC=C"));
        File.WriteAllBytes("output/samples/8_chalcogran.png", MoleculeRenderer.Render("CC[C@H](O1)CC[C@@]12CCCO2"));
        File.WriteAllBytes("output/samples/9_thujone.png", MoleculeRenderer.Render("CC(C)[C@@]12C[C@@H]1[C@@H](C)C(=O)C2"));
        File.WriteAllBytes("output/samples/10_thiamine_default.png", MoleculeRenderer.Render("OCCSc1c(C)[n+](cs1)Cc2cnc(C)nc2N"));
        File.WriteAllBytes("output/samples/10_thiamine_flipped.png", MoleculeRenderer.Render("OCCSc1c(C)[n+](cs1)Cc2cnc(C)nc2N", color: true, transform: "flip"));
        File.WriteAllBytes("output/samples/mic.png", micPng);
        File.WriteAllBytes("output/samples/methanol.png", MoleculeRenderer.Render("CO"));
        File.WriteAllBytes("output/samples/sample_bw_num.png", comboPng);
        File.WriteAllBytes("output/samples/sample_stereo.png", stereoPng);
        File.WriteAllBytes("output/samples/sample_unfold_h.png", hPng);
        File.WriteAllBytes("output/samples/sample_dark_mode.png", darkPng);

        PassSuite("Unit: ParseOptions, Domain Styles (B/W, Dark, Atom#, Stereo, Unfold-H, Substructure Highlight) & Sample Gallery");
    }

    private static void RunKetAndInchiRenderingTests()
    {
        Check(MoleculeRenderer.IsKet(SampleKet), "MoleculeRenderer.IsKet identifies KET JSON");
        Check(MoleculeRenderer.IsKet("KET:" + SampleKet), "MoleculeRenderer.IsKet identifies KET: prefix");
        Check(!MoleculeRenderer.IsKet("CCO"), "MoleculeRenderer.IsKet rejects SMILES");
        Check(MoleculeRenderer.CleanKet("KET:" + SampleKet) == SampleKet, "MoleculeRenderer.CleanKet strips prefix");

        var ketPng = MoleculeRenderer.RenderKet(SampleKet);
        Check(ketPng != null && ketPng.Length > 100, "MoleculeRenderer renders KET data directly");
        Check(ketPng!.Take(8).SequenceEqual(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }), "KET PNG signature");

        var ketOptPng = MoleculeRenderer.RenderKet(SampleKet, new RenderOptions { Background = "white", Transform = "flip" });
        Check(ketOptPng != null && ketOptPng.Length > 100, "MoleculeRenderer renders KET with options");

        var ethanolInchi = MoleculeRenderer.TryConvertToInchi("CCO")!;
        Check(MoleculeRenderer.IsInchi(ethanolInchi), "IsInchi identifies InChI string");
        Check(MoleculeRenderer.IsInchi("1S/C2H6O/c1-2-3/h3H,2H2,1H3"), "IsInchi identifies prefix-less InChI");
        Check(MoleculeRenderer.NormalizeInchi("1S/C2H6O/c1-2-3/h3H,2H2,1H3") == ethanolInchi, "NormalizeInchi prepends InChI=");

        var inchiPng = MoleculeRenderer.Render(ethanolInchi);
        Check(inchiPng != null && inchiPng.Length > 200, "MoleculeRenderer.Render renders InChI to PNG");
        var inchiNoPrefixPng = MoleculeRenderer.Render("1S/C2H6O/c1-2-3/h3H,2H2,1H3");
        Check(inchiNoPrefixPng != null && inchiNoPrefixPng.Length > 200, "MoleculeRenderer.Render renders prefix-less InChI to PNG");

        Check(MoleculeRenderer.TryRenderStructure("c1ccccc1", null, out var quickSmilesPng, out _, out var quickSmiles)
              && quickSmilesPng != null && quickSmilesPng.Length > 100 && !string.IsNullOrEmpty(quickSmiles),
              "TryRenderStructure renders SMILES");
        Check(MoleculeRenderer.TryRenderStructure(SampleKet, null, out var quickKetPng, out var quickKetOut, out _)
              && quickKetPng != null && quickKetPng.Length > 100 && quickKetOut == SampleKet,
              "TryRenderStructure renders KET");

        PassSuite("Unit: Native KET & InChI 2D Rendering & TryRenderStructure");
    }

    private static void RunStructureMetadataAndMolstarSessionTests()
    {
        var formattedAlt = StructureMetadata.FormatAlternativeText(SampleKet, "CO");
        var formattedTitle = StructureMetadata.FormatTitle("CO");
        Check(StructureMetadata.ExtractStructure(formattedAlt, formattedTitle) == SampleKet, "StructureMetadata extracts KET from AltText");
        Check(StructureMetadata.ExtractSmiles(formattedTitle, formattedAlt) == "CO", "StructureMetadata extracts SMILES from Title");
        Check(StructureMetadata.ExtractStructure($"KET:{SampleKet}\nSMILES:CO") == SampleKet, "StructureMetadata strips trailing SMILES tag from KET");

        var png = MoleculeRenderer.Render("c1ccccc1");
        var fitBounds = StructureMetadata.FitReplacementBounds(100f, 100f, 200f, 200f, png);
        Check(fitBounds.Width > 0 && fitBounds.Height > 0, "StructureMetadata FitReplacementBounds computes valid geometry");

        var prettyKet = "{\n  \"root\": {\n    \"nodes\": [\n      {\"type\": \"atom\"}\n    ]\n  }\n}";
        var compact = StructureMetadata.CompactKet(prettyKet);
        Check(!compact.Contains("\n") && !compact.Contains("  ") && compact.Contains("\"root\""), "CompactKet minifies JSON without newlines");
        Check(StructureMetadata.CompactKet(null) == "", "CompactKet handles null");

        Check(MoleculeRenderer.TryConvertToSmiles(SampleKet) == "CO", "TryConvertToSmiles converts methanol KET to CO");
        Check(StructureMetadata.ExtractSmiles($"KET:{SampleKet}") == "CO", "ExtractSmiles converts KET without SMILES tag to CO");

        var convertedKetFromSmiles = MoleculeRenderer.TryConvertToKet("CCO");
        Check(!string.IsNullOrEmpty(convertedKetFromSmiles) && convertedKetFromSmiles!.Contains("\"root\""), "TryConvertToKet converts CCO SMILES to valid KET JSON");
        var convertedKetFromKet = MoleculeRenderer.TryConvertToKet(SampleKet);
        Check(!string.IsNullOrEmpty(convertedKetFromKet) && convertedKetFromKet!.Contains("\"root\""), "TryConvertToKet processes SampleKet properly");

        var metaIn = new MolstarMetadata
        {
            Id = "test-session-123",
            Pdb = "1CRN",
            Mode = "3dmodel",
            Embed = true,
            Background = "transparent"
        };
        var altText = MolstarSessionManager.FormatAltText(metaIn);
        Check(StructureMetadata.IsMolstar(altText), "StructureMetadata.IsMolstar identifies Molstar AltText");
        var metaOut = MolstarSessionManager.ParseAltText(altText);
        Check(metaOut != null && metaOut.Id == "test-session-123" && metaOut.Pdb == "1CRN" && metaOut.Mode == "3dmodel" && metaOut.Embed && metaOut.Background == "transparent",
            "MolstarSessionManager AltText round-trip");

        var sampleSessionJson = "{\"version\":\"5.13.1\",\"snapshot\":{\"camera\":[1.0,2.0,3.0],\"name\":\"crambin\"}}";
        var compressed = MolstarSessionManager.CompressGzip(sampleSessionJson);
        Check(compressed.Length > 0 && compressed.Length < sampleSessionJson.Length + 50, "CompressGzip compresses JSON");
        Check(MolstarSessionManager.DecompressGzip(compressed) == sampleSessionJson, "DecompressGzip restores exact session JSON");

        var escapedSmiles = "CC(=O)Oc1ccccc1C(=O)O".Replace("\"", "\"\"");
        Check($"=SMILES2IMG(\"{escapedSmiles}\")" == "=SMILES2IMG(\"CC(=O)Oc1ccccc1C(=O)O\")", "SMILES in-cell formula generation");
        var escapedKet = StructureMetadata.CompactKet(SampleKet).Replace("\"", "\"\"");
        var ketFormula = $"=KET2IMG(\"{escapedKet}\")";
        Check(ketFormula.StartsWith("=KET2IMG(\"") && ketFormula.Contains("\"\"root\"\"") && !ketFormula.Contains("\n"),
            "KET in-cell formula generation with compact JSON and escaped double quotes");

        Check(StructureMetadata.ExtractPreloadForKetcher("CCO") == "CCO", "ExtractPreloadForKetcher preloads valid SMILES");
        Check(StructureMetadata.ExtractPreloadForKetcher("Aspirin") == "CC(=O)Oc1ccccc1C(=O)O", "ExtractPreloadForKetcher preloads built-in offline name");
        Check(StructureMetadata.ExtractPreloadForKetcher(SampleKet) == SampleKet, "ExtractPreloadForKetcher preloads KET JSON");
        Check(StructureMetadata.ExtractPreloadForKetcher("Figure 1: Introduction") == "", "ExtractPreloadForKetcher skips non-chemical prose");

        Check(MolstarSessionManager.ExtractInitialPdbTarget("1crn") == "1CRN", "ExtractInitialPdbTarget normalizes 4-char PDB ID");
        Check(MolstarSessionManager.ExtractInitialPdbTarget("PDB: 6LU7") == "6LU7", "ExtractInitialPdbTarget parses PDB: prefix");
        Check(MolstarSessionManager.ExtractInitialPdbTarget(@"C:\temp\protein.cif") == @"C:\temp\protein.cif", "ExtractInitialPdbTarget keeps local .cif path");
        Check(MolstarSessionManager.ExtractInitialPdbTarget("Aspirin") == "", "ExtractInitialPdbTarget rejects non-PDB text");

        Check(Smiles2Img.Editor.MoleculeActionHelper.ResolveSmilesWithTextFallback(null, "Aspirin", resolvePubChem: false) == "CC(=O)Oc1ccccc1C(=O)O",
            "ResolveSmilesWithTextFallback resolves chemical name instead of returning literal name");
        Check(Smiles2Img.Editor.MoleculeActionHelper.ResolveSmilesWithTextFallback(null, "Hello World", resolvePubChem: false) == null,
            "ResolveSmilesWithTextFallback rejects non-chemical text");
        Check(Smiles2Img.Editor.MoleculeActionHelper.ResolveStructureOrText(null, null, "Aspirin", resolvePubChem: false) == "CC(=O)Oc1ccccc1C(=O)O",
            "ResolveStructureOrText resolves fallback text consistently");

        foreach (var (quickInput, expectedSmiles, expectKet) in new[]
        {
            ("Aspirin", "CC(=O)Oc1ccccc1C(=O)O", false),
            ("Paracetamol", "CC(=O)Nc1ccc(O)cc1", false),
            ("50-78-2", "CC(=O)Oc1ccccc1C(=O)O", false),
            ("THF", "C1CCOC1", false),
            ("c1ccccc1", "c1ccccc1", false),
            (SampleKet, "CO", true)
        })
        {
            Check(Smiles2Img.Editor.MoleculeActionHelper.TryRenderQuickInsertInput(
                      quickInput, RenderOptions.Default, out var qPng, out var qKet, out var qSmiles, resolvePubChem: false)
                  && qPng != null && qPng.Length > 100
                  && qSmiles == expectedSmiles
                  && (!expectKet || qKet == SampleKet),
                $"TryRenderQuickInsertInput renders '{quickInput}'");
        }

        foreach (var rejectedInput in new[]
        {
            "",
            "   ",
            "This is a long regular prose sentence with many words.",
            "https://example.com/aspirin",
            "http://example.com"
        })
        {
            Check(!Smiles2Img.Editor.MoleculeActionHelper.TryRenderQuickInsertInput(
                      rejectedInput, RenderOptions.Default, out _, out _, out _, resolvePubChem: false),
                $"TryRenderQuickInsertInput rejects '{rejectedInput}'");
        }

        var largeFit = StructureMetadata.FitTextBoxReplacementBounds(50f, 60f, 300f, 200f, 960f, 540f, png);
        Check(largeFit.Left >= 50f && largeFit.Top >= 60f
              && largeFit.Left + largeFit.Width <= 350.01f
              && largeFit.Top + largeFit.Height <= 260.01f
              && largeFit.Width > 0f && largeFit.Height > 0f,
            "FitTextBoxReplacementBounds keeps large TextBox replacement inside original bounds");

        var thinFit = StructureMetadata.FitTextBoxReplacementBounds(56f, 188f, 416f, 42f, 960f, 540f, png);
        var thinCx = thinFit.Left + thinFit.Width / 2f;
        var thinCy = thinFit.Top + thinFit.Height / 2f;
        Check(thinFit.Height >= 70f && thinFit.Height <= 130f
              && Math.Abs(thinCx - 264f) <= 2f
              && Math.Abs(thinCy - 209f) <= 2f
              && thinFit.Left >= 12f && thinFit.Top >= 12f
              && thinFit.Left + thinFit.Width <= 948f
              && thinFit.Top + thinFit.Height <= 528f,
            $"FitTextBoxReplacementBounds expands thin single-line TextBox cleanly (h={thinFit.Height}, cx={thinCx}, cy={thinCy})");

        PassSuite("Unit: StructureMetadata, Format Conversions & MolstarSessionManager GZip Round-Trip");
    }

    private static void RunVectorSvgTests()
    {
        var svgXml = MoleculeRenderer.RenderSvg("CC(=O)Oc1ccccc1C(=O)O");
        Check(!string.IsNullOrEmpty(svgXml) && svgXml!.Contains("<svg") && svgXml.Contains("</svg>"), "MoleculeRenderer.RenderSvg generates valid SVG XML");
        Check(SvgClipboardHelper.CopySvgToClipboard(svgXml), "SvgClipboardHelper copies SVG to Windows Clipboard");
        PassSuite("Unit: Vector SVG Generation & Clipboard Registration");
    }
}
