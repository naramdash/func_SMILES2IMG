using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using ExcelDna.Integration;
using Smiles2Img;
using Smiles2Img.Rendering;
using Smiles2Img.Word;
using Smiles2Img.PowerPoint;

namespace Smiles2Img.Tests;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            NativeIndigo.Initialize(AppDomain.CurrentDomain.BaseDirectory);
            Console.WriteLine("Process: " + (Environment.Is64BitProcess ? "x64" : "x86"));
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
                Check(ink > 20, "PNG must contain a visible molecule");
                Console.WriteLine("PASS PNG: " + smiles);
            }
            foreach (var input in new[] { "", " ", new string('C', 2001), "not-a-smiles", "C1CC", "C(C" })
            {
                try { MoleculeRenderer.Render(input); throw new Exception("Accepted invalid SMILES: " + input); }
                catch (InvalidSmilesException) { }
            }
            Console.WriteLine("PASS invalid input");
            var name = typeof(Functions).GetMethod(nameof(Functions.Img))!.GetCustomAttributes(typeof(ExcelFunctionAttribute), false)
                .Cast<ExcelFunctionAttribute>().Single().Name;
            Check(name == "SMILES2IMG", "Excel function name");
            Console.WriteLine("PASS Excel function registration name");

            var floatAttribute = typeof(FloatFunctions).GetMethod(nameof(FloatFunctions.ImgFloat))!.GetCustomAttributes(typeof(ExcelFunctionAttribute), false)
                .Cast<ExcelFunctionAttribute>().Single();
            Check(floatAttribute.Name == "SMILES2IMG.FLOAT", "Floating Excel function name");
            Check(floatAttribute.IsMacroType, "Floating function must be IsMacroType=true for caller reference");
            Check(floatAttribute.Description.StartsWith("[Excel 2016+]") && floatAttribute.Description.Length <= 255, "Floating function states its Excel requirement");
            var nativeAttribute = typeof(Functions).GetMethod(nameof(Functions.Img))!.GetCustomAttributes(typeof(ExcelFunctionAttribute), false)
                .Cast<ExcelFunctionAttribute>().Single();
            Check(nativeAttribute.IsMacroType, "Native function must be IsMacroType=true");
            var nativeDescription = nativeAttribute.Description;
            Check(nativeDescription.StartsWith("[Excel 2024 / Microsoft 365]") && nativeDescription.Contains("SMILES2IMG.FLOAT") && nativeDescription.Length <= 255,
                "In-cell function states its Excel requirement");
            Console.WriteLine("PASS SMILES2IMG.FLOAT registration and version notes");

            var wide = FloatingPictures.Fit(100, 50, 200, 100, 3.0); // cell wider than needed: height-bound, centered horizontally
            Check(Math.Abs(wide.Height - 96) < 1e-9 && Math.Abs(wide.Width - 288) > 1 && Math.Abs(wide.Width / wide.Height - 3.0) < 1e-9 || Math.Abs(wide.Width - 196) < 1e-9,
                "Fit keeps aspect ratio");
            var tall = FloatingPictures.Fit(0, 0, 100, 300, 2.0); // width-bound, centered vertically
            Check(Math.Abs(tall.Width - 96) < 1e-9 && Math.Abs(tall.Height - 48) < 1e-9 && Math.Abs(tall.Left - 2) < 1e-9 && Math.Abs(tall.Top - 126) < 1e-9,
                "Fit centers a width-bound picture");
            var flat = FloatingPictures.Fit(0, 0, 300, 100, 1.0); // height-bound, centered horizontally
            Check(Math.Abs(flat.Width - 96) < 1e-9 && Math.Abs(flat.Height - 96) < 1e-9 && Math.Abs(flat.Left - 102) < 1e-9 && Math.Abs(flat.Top - 2) < 1e-9,
                "Fit centers a height-bound picture");
            var title = FloatingPictures.Title("C_ABC", 1.75);
            Check(FloatingPictures.TryParseTitle(title, out var titleKey, out var titleAspect) && titleKey == "C_ABC" && titleAspect == 1.75, "Title round-trip");
            Check(!FloatingPictures.TryParseTitle("Picture 1", out _, out _) && !FloatingPictures.TryParseTitle(null, out _, out _), "Foreign shapes are not owned");
            var aspectPng = MoleculeRenderer.Render("CCO");
            using (var aspectStream = new MemoryStream(aspectPng))
            using (var aspectBitmap = new Bitmap(aspectStream))
                Check(Math.Abs(FloatingPictures.Aspect(aspectPng) - (double)aspectBitmap.Width / aspectBitmap.Height) < 1e-9, "PNG aspect from header");
            Console.WriteLine("PASS SMILES2IMG.FLOAT fit, title and aspect helpers");

            // Verify ParseOptions argument separation & domain styles
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

            // Domain style tokens & combinations
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
            var optMonoAlias = Functions.ParseOptions(null, "mono", null);
            Check(optMonoAlias.Color == false, "ParseOptions mono alias");
            var optBlackAlias = Functions.ParseOptions(null, "black", null);
            Check(optBlackAlias.Color == false, "ParseOptions black alias");
            var optIdxAlias = Functions.ParseOptions(null, "idx", null);
            Check(optIdxAlias.AtomNumbers == true, "ParseOptions idx alias");
            var optChiralAlias = Functions.ParseOptions(null, "chiral", null);
            Check(optChiralAlias.StereoLabels == true, "ParseOptions chiral alias");
            var optBoolCompat = Functions.ParseOptions(null, 0.0, null);
            Check(optBoolCompat.Color == false, "ParseOptions numeric 0 for B/W compatibility");
            var optPipe = Functions.ParseOptions(null, "bw|num", null);
            Check(optPipe.Color == false && optPipe.AtomNumbers == true, "ParseOptions bw|num pipe delimiter");
            var optSpace = Functions.ParseOptions(null, "white stereo", null);
            Check(optSpace.BaseColor == "white" && optSpace.StereoLabels == true, "ParseOptions white stereo space delimiter");

            // Ultimate full-option tests (spaces & pipes)
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

            Console.WriteLine("PASS ParseOptions domain styles and composite tokens");

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
            Console.WriteLine("PASS Black and White rendering");

            // Render tests for new domain features
            var numOptions = Functions.ParseOptions(null, "num", null);
            var numPng = MoleculeRenderer.Render("CCO", numOptions);
            Check(numPng.Length > 100, "Render with atom numbers");
            Console.WriteLine("PASS Atom numbering rendering");

            var stereoOptions = Functions.ParseOptions(null, "stereo", null);
            var stereoPng = MoleculeRenderer.Render("C[C@H](O)C(=O)O", stereoOptions);
            Check(stereoPng.Length > 100, "Render with stereo labels");
            Console.WriteLine("PASS Stereochemistry label rendering");

            var hOptions = Functions.ParseOptions(null, "h", null);
            var hPng = MoleculeRenderer.Render("CCO", hOptions);
            Check(hPng.Length > 100, "Render with unfolded hydrogens");
            Console.WriteLine("PASS Hydrogens unfolding rendering");

            var carbOptions = Functions.ParseOptions(null, "all-c", null);
            var carbPng = MoleculeRenderer.Render("c1ccccc1", carbOptions);
            Check(carbPng.Length > 100, "Render with explicit carbons");
            Console.WriteLine("PASS Explicit carbons rendering");

            var comboOptions = Functions.ParseOptions(null, "bw,num,stereo", null);
            var comboPng = MoleculeRenderer.Render("C[C@H](O)C(=O)O", comboOptions);
            Check(comboPng.Length > 100, "Render with composite bw,num,stereo");
            Console.WriteLine("PASS Composite options (bw,num,stereo) rendering");

            var ultimatePng = MoleculeRenderer.Render("CC(=O)Oc1ccccc1C(=O)O", optUltimateSpace);
            Check(ultimatePng.Length > 100, "Render with ultimate full-option");
            Console.WriteLine("PASS Ultimate full-option rendering");

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
            Console.WriteLine("PASS Dark mode (white ink + dark bg) rendering");

            var aflatoxinPng = MoleculeRenderer.Render("O1C=C[C@H]([C@H]1O2)c3c2cc(OC)c4c3OC(=O)C5=C4CCC(=O)5");
            Check(aflatoxinPng.Length > 100, "Aflatoxin junction H render");
            Console.WriteLine("PASS Ring-junction chiral H rendering");

            var micPng = MoleculeRenderer.Render("CN=C=O");
            Check(micPng.Length > 100, "MIC small molecule render");
            Console.WriteLine("PASS Small molecule render (MIC)");

            var nicFlippedPng = MoleculeRenderer.Render("CN1CCC[C@H]1c2cccnc2", color: true, transform: "flip");
            Check(nicFlippedPng.Length > 100, "Nicotine flip transform render");
            Console.WriteLine("PASS Transform rendering (flip)");

            var artemisinin = "CC1CC2CC3(C)OO4C(O2)(C1C(=O)O3)C(C)CC4";
            var artemisininPng = MoleculeRenderer.Render(artemisinin);
            Check(artemisininPng.Length > 100, "Artemisinin render");
            Console.WriteLine("PASS Artemisinin render");

            var defaultTransPng = MoleculeRenderer.Render("CCO");
            using (var ms = new MemoryStream(defaultTransPng))
            using (var bmp = new Bitmap(ms))
            {
                Check(bmp.GetPixel(0, 0).A == 0, "Default background corner pixel must have alpha 0 (transparent)");
            }
            Console.WriteLine("PASS Transparent default background rendering");

            var cssYellowPng = MoleculeRenderer.Render("CCO", background: "yellow");
            using (var ms = new MemoryStream(cssYellowPng))
            using (var bmp = new Bitmap(ms))
            {
                var corner = bmp.GetPixel(0, 0);
                Check(corner.R > 240 && corner.G > 240 && corner.B < 20 && corner.A == 255, "CSS color 'yellow' corner pixel");
            }
            Console.WriteLine("PASS CSS color (yellow) rendering");

            var whitePng = MoleculeRenderer.Render("CCO", background: "white");
            using (var ms = new MemoryStream(whitePng))
            using (var bmp = new Bitmap(ms))
            {
                var corner = bmp.GetPixel(0, 0);
                Check(corner.R == 255 && corner.G == 255 && corner.B == 255 && corner.A == 255, "Explicit white background corner pixel");
            }
            Console.WriteLine("PASS Explicit white background rendering");

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

            Console.WriteLine("PASS Rendered all 10 reference sample molecules + domain feature samples");

            // Ketcher integration tests
            Check(KetcherEditor.IsLikelySmiles("c1ccccc1"), "Ketcher detects benzene SMILES");
            Check(KetcherEditor.IsLikelySmiles("CC(=O)Oc1ccccc1C(=O)O"), "Ketcher detects aspirin SMILES");
            Check(!KetcherEditor.IsLikelySmiles("12345"), "Ketcher rejects plain numbers");
            Check(!KetcherEditor.IsLikelySmiles("Hello world"), "Ketcher rejects plain text");
            Check(!KetcherEditor.IsLikelySmiles(""), "Ketcher rejects empty string");

            var fingerprint = KetcherAssetManager.GetBundleFingerprint();
            Check(!string.IsNullOrWhiteSpace(fingerprint) && fingerprint.Contains("-"), "Ketcher bundle fingerprint is generated");
            var ketcherFolder = KetcherAssetManager.EnsureKetcherFolder();
            Check(Directory.Exists(ketcherFolder), "Ketcher folder exists");
            Check(File.Exists(Path.Combine(ketcherFolder, "index.html")), "Ketcher index.html exists");
            var indexContent = File.ReadAllText(Path.Combine(ketcherFolder, "index.html"));
            Check(indexContent.Contains("getCanvasSvg"), "Ketcher index.html includes getCanvasSvg");
            Check(indexContent.Contains("calculateCanvasBBox"), "Ketcher index.html includes calculateCanvasBBox");
            Check(indexContent.Contains("toggleBackground"), "Ketcher index.html includes toggleBackground");
            Console.WriteLine("PASS Ketcher editor SMILES validation, fingerprint, and offline asset extraction");

            // Test native KET rendering & KET2IMG functions
            var sampleKet = "{\"root\":{\"nodes\":[{\"$ref\":\"mol0\"}]},\"mol0\":{\"type\":\"molecule\",\"atoms\":[{\"label\":\"C\",\"location\":[0,0,0]},{\"label\":\"O\",\"location\":[1,0,0]}],\"bonds\":[{\"type\":1,\"atoms\":[0,1]}]}}";
            Check(MoleculeRenderer.IsKet(sampleKet), "MoleculeRenderer.IsKet identifies KET JSON");
            Check(MoleculeRenderer.IsKet("KET:" + sampleKet), "MoleculeRenderer.IsKet identifies KET: prefix");
            Check(!MoleculeRenderer.IsKet("CCO"), "MoleculeRenderer.IsKet rejects SMILES");
            Check(MoleculeRenderer.CleanKet("KET:" + sampleKet) == sampleKet, "MoleculeRenderer.CleanKet strips prefix");

            var ketPng = MoleculeRenderer.RenderKet(sampleKet);
            Check(ketPng != null && ketPng.Length > 100, "MoleculeRenderer renders KET data directly");
            Check(ketPng!.Take(8).SequenceEqual(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }), "KET PNG signature");

            var ketOptPng = MoleculeRenderer.RenderKet(sampleKet, new RenderOptions { Background = "white", Transform = "flip" });
            Check(ketOptPng != null && ketOptPng.Length > 100, "MoleculeRenderer renders KET with options");

            var ketAttr = typeof(Functions).GetMethod(nameof(Functions.KetImg))!.GetCustomAttributes(typeof(ExcelFunctionAttribute), false)
                .Cast<ExcelFunctionAttribute>().Single();
            Check(ketAttr.Name == "KET2IMG" && ketAttr.IsMacroType, "KET2IMG function registration");

            var ketFloatAttr = typeof(FloatFunctions).GetMethod(nameof(FloatFunctions.KetImgFloat))!.GetCustomAttributes(typeof(ExcelFunctionAttribute), false)
                .Cast<ExcelFunctionAttribute>().Single();
            Check(ketFloatAttr.Name == "KET2IMG.FLOAT" && ketFloatAttr.IsMacroType, "KET2IMG.FLOAT function registration");
            Console.WriteLine("PASS Native KET rendering, IsKet, and KET2IMG function registration");

            // Test MoleculeInfo calculation (Formula, Molecular Weight, Exact Mass, Canonical SMILES, PubChem Name)
            var aspirinInfo = MoleculeRenderer.TryGetMoleculeInfo("CC(=O)Oc1ccccc1C(=O)O");
            Check(aspirinInfo != null, "TryGetMoleculeInfo succeeds for Aspirin SMILES");
            Check(aspirinInfo!.Formula == "C9H8O4", "Aspirin formula is C9H8O4, got: " + aspirinInfo.Formula);
            Check(Math.Abs(aspirinInfo.MolecularWeight - 180.16) < 0.1, "Aspirin MW ~ 180.16, got: " + aspirinInfo.MolecularWeight);
            Check(Math.Abs(aspirinInfo.ExactMass - 180.0423) < 0.01, "Aspirin Exact Mass ~ 180.0423, got: " + aspirinInfo.ExactMass);
            Check(!string.IsNullOrEmpty(aspirinInfo.Smiles), "Aspirin canonical SMILES non-empty");

            var resolvedEthanol = PubChemResolver.ResolveCompound("CCO");
            Check(string.Equals(resolvedEthanol.CompoundName, "Ethanol", StringComparison.OrdinalIgnoreCase), "PubChem resolves CCO to Ethanol, got: " + resolvedEthanol.CompoundName);

            var ketInfo = MoleculeRenderer.TryGetMoleculeInfo(sampleKet);
            Check(ketInfo != null && ketInfo.Formula == "CH4O", "TryGetMoleculeInfo succeeds for methanol KET, got: " + (ketInfo?.Formula ?? "null"));

            // Test Excel Functions SMILES2NAME / SMILES2MW / SMILES2FORMULA
            var mwVal = Functions.SmilesMW("CCO");
            Check(mwVal is double d && Math.Abs(d - 46.07) < 0.1, "Functions.SmilesMW for CCO ~ 46.07, got: " + mwVal);
            var formVal = Functions.SmilesFormula("CCO");
            Check(Equals(formVal, "C2H6O"), "Functions.SmilesFormula for CCO is C2H6O, got: " + formVal);

            // Test InChI & InChIKey conversion
            var ethanolInchi = MoleculeRenderer.TryConvertToInchi("CCO");
            Check(ethanolInchi != null && ethanolInchi.StartsWith("InChI=1S/C2H6O", StringComparison.OrdinalIgnoreCase), "InChI for CCO starts with InChI=1S/C2H6O, got: " + ethanolInchi);
            var ethanolInchiKey = MoleculeRenderer.TryConvertToInchiKey("CCO");
            Check(string.Equals(ethanolInchiKey, "LFQSCWFLJHTTHZ-UHFFFAOYSA-N", StringComparison.OrdinalIgnoreCase), "InChIKey for CCO is LFQSCWFLJHTTHZ-UHFFFAOYSA-N, got: " + ethanolInchiKey);
            Check(!string.IsNullOrEmpty(aspirinInfo.InChI) && aspirinInfo.InChI.StartsWith("InChI=1S/C9H8O4"), "Aspirin info includes InChI");
            Check(!string.IsNullOrEmpty(aspirinInfo.InChIKey), "Aspirin info includes InChIKey");

            var inchiFuncVal = Functions.SmilesInchi("CCO");
            Check(inchiFuncVal is string s && s.StartsWith("InChI=1S/C2H6O"), "Functions.SmilesInchi works for CCO");
            var inchiKeyFuncVal = Functions.SmilesInchiKey("CCO");
            Check(Equals(inchiKeyFuncVal, "LFQSCWFLJHTTHZ-UHFFFAOYSA-N"), "Functions.SmilesInchiKey works for CCO");

            // Test MoleculeInfoDialog instantiation
            using (var dlg = new Smiles2Img.Editor.MoleculeInfoDialog(aspirinInfo))
            {
                Check(dlg != null && dlg.Text.Contains("Molecule"), "MoleculeInfoDialog instantiates cleanly");
            }

            Console.WriteLine("PASS MoleculeInfo calculation (Name, Formula, MW, InChI, ExactMass) & PubChem resolution");

            // Test InChI Input Support & Functions
            Check(MoleculeRenderer.IsInchi(ethanolInchi), "IsInchi identifies InChI string");
            Check(MoleculeRenderer.IsInchi("1S/C2H6O/c1-2-3/h3H,2H2,1H3"), "IsInchi identifies prefix-less InChI");
            Check(MoleculeRenderer.NormalizeInchi("1S/C2H6O/c1-2-3/h3H,2H2,1H3") == ethanolInchi, "NormalizeInchi prepends InChI=");

            // Test InChI rendering to PNG
            var inchiPng = MoleculeRenderer.Render(ethanolInchi!);
            Check(inchiPng != null && inchiPng.Length > 200, "MoleculeRenderer.Render renders InChI to PNG");
            var inchiNoPrefixPng = MoleculeRenderer.Render("1S/C2H6O/c1-2-3/h3H,2H2,1H3");
            Check(inchiNoPrefixPng != null && inchiNoPrefixPng.Length > 200, "MoleculeRenderer.Render renders prefix-less InChI to PNG");

            // Test InChI conversion to SMILES, Formula, MW, InChIKey
            var convertedInchiSmiles = MoleculeRenderer.TryConvertToSmiles(ethanolInchi!);
            Check(convertedInchiSmiles == "CCO", "TryConvertToSmiles converts ethanol InChI to CCO, got: " + convertedInchiSmiles);
            var inchiFormulaVal = Functions.InchiFormula(ethanolInchi!);
            Check(Equals(inchiFormulaVal, "C2H6O"), "Functions.InchiFormula for ethanol InChI is C2H6O, got: " + inchiFormulaVal);
            var inchiMWVal = Functions.InchiMW(ethanolInchi!);
            Check(Equals(inchiMWVal, 46.07), "Functions.InchiMW for ethanol InChI is 46.07, got: " + inchiMWVal);
            var inchiKeyFromInchi = Functions.InchiKey(ethanolInchi!);
            Check(Equals(inchiKeyFromInchi, "LFQSCWFLJHTTHZ-UHFFFAOYSA-N"), "Functions.InchiKey from InChI string");
            var inchiSmilesFuncVal = Functions.InchiSmiles(ethanolInchi!);
            Check(Equals(inchiSmilesFuncVal, "CCO"), "Functions.InchiSmiles from InChI string");

            // Test pure local calculation with resolvePubChem: false (zero network calls)
            var pureLocalInfo = MoleculeRenderer.TryGetMoleculeInfo("CCO", resolvePubChem: false);
            Check(pureLocalInfo != null && pureLocalInfo.Formula == "C2H6O" && pureLocalInfo.MolecularWeight > 46.0 && pureLocalInfo.CompoundName == "",
                "TryGetMoleculeInfo with resolvePubChem: false computes local properties without PubChem");
            Console.WriteLine("PASS InChI input rendering, INCHI2* worksheet functions, and zero-network local calculations");

            // Test TryRenderStructure (Quick Paste helper)
            Check(MoleculeRenderer.TryRenderStructure("c1ccccc1", null, out var quickSmilesPng, out var quickKet, out var quickSmiles)
                  && quickSmilesPng != null && quickSmilesPng.Length > 100 && !string.IsNullOrEmpty(quickSmiles),
                  "TryRenderStructure renders SMILES");
            Check(MoleculeRenderer.TryRenderStructure(sampleKet, null, out var quickKetPng, out var quickKetOut, out _)
                  && quickKetPng != null && quickKetPng.Length > 100 && quickKetOut == sampleKet,
                  "TryRenderStructure renders KET");
            Console.WriteLine("PASS TryRenderStructure for Quick Paste");

            // Test StructureMetadata shared helper
            var formattedAlt = StructureMetadata.FormatAlternativeText(sampleKet, "CO");
            var formattedTitle = StructureMetadata.FormatTitle("CO");
            Check(StructureMetadata.ExtractStructure(formattedAlt, formattedTitle) == sampleKet, "StructureMetadata extracts KET from AltText");
            Check(StructureMetadata.ExtractSmiles(formattedTitle, formattedAlt) == "CO", "StructureMetadata extracts SMILES from Title");
            Check(StructureMetadata.ExtractStructure($"KET:{sampleKet}\nSMILES:CO") == sampleKet, "StructureMetadata strips trailing SMILES tag from KET");
            var fitBounds = StructureMetadata.FitReplacementBounds(100f, 100f, 200f, 200f, quickSmilesPng!);
            Check(fitBounds.Width > 0 && fitBounds.Height > 0, "StructureMetadata FitReplacementBounds computes valid geometry");
            var prettyKet = "{\n  \"root\": {\n    \"nodes\": [\n      {\"type\": \"atom\"}\n    ]\n  }\n}";
            var compact = StructureMetadata.CompactKet(prettyKet);
            Check(!compact.Contains("\n") && !compact.Contains("  ") && compact.Contains("\"root\""), "CompactKet minifies JSON without newlines");
            Check(StructureMetadata.CompactKet(null) == "", "CompactKet handles null");

            // Test KET to SMILES conversion
            var convertedSmiles = MoleculeRenderer.TryConvertToSmiles(sampleKet);
            Check(convertedSmiles == "CO", "TryConvertToSmiles converts methanol KET to CO, got: " + convertedSmiles);
            var smilesFromKetAlt = StructureMetadata.ExtractSmiles($"KET:{sampleKet}");
            Check(smilesFromKetAlt == "CO", "ExtractSmiles converts KET without SMILES tag to CO, got: " + smilesFromKetAlt);

            // Test SMILES to KET and KET conversion
            var convertedKetFromSmiles = MoleculeRenderer.TryConvertToKet("CCO");
            Check(!string.IsNullOrEmpty(convertedKetFromSmiles) && convertedKetFromSmiles!.Contains("\"root\""), "TryConvertToKet converts CCO SMILES to valid KET JSON");
            var convertedKetFromKet = MoleculeRenderer.TryConvertToKet(sampleKet);
            Check(!string.IsNullOrEmpty(convertedKetFromKet) && convertedKetFromKet!.Contains("\"root\""), "TryConvertToKet processes sampleKet properly");
            Console.WriteLine("PASS StructureMetadata shared helper, CompactKet, KET-to-SMILES & SMILES-to-KET conversion");

            // Test MolstarSessionManager & MolstarMetadata
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
            var decompressed = MolstarSessionManager.DecompressGzip(compressed);
            Check(decompressed == sampleSessionJson, "DecompressGzip restores exact session JSON");
            Console.WriteLine("PASS MolstarSessionManager metadata & gzip compression round-trip");

            // Live WebView2 + Ketcher E2E Test (1: SMILES -> KET + Transparent Canvas PNG)
            string? e2eKet = null;
            using (var form1 = new KetcherForm("PowerPoint", "CC(=O)Oc1ccccc1C(=O)O")
            {
                AutoApplyAfterLoadForTest = true,
                ToggleWhiteBackgroundForTest = false
            })
            {
                var dr1 = form1.ShowDialog();
                Check(dr1 == System.Windows.Forms.DialogResult.OK, "E2E KetcherForm #1 returned DialogResult.OK");
                Check(!string.IsNullOrWhiteSpace(form1.ResultKet) && form1.ResultKet!.Contains("\"root\""), "E2E KetcherForm #1 returned valid KET JSON");
                Check(!string.IsNullOrWhiteSpace(form1.ResultSmiles), "E2E KetcherForm #1 returned non-empty SMILES: " + form1.ResultSmiles);
                Check(form1.ResultBackground == "transparent", "E2E KetcherForm #1 background is transparent");
                Check(form1.ResultPngBytes != null && form1.ResultPngBytes.Length > 500, "E2E KetcherForm #1 captured canvas PNG bytes (" + (form1.ResultPngBytes?.Length ?? 0) + " bytes)");
                using (var ms1 = new MemoryStream(form1.ResultPngBytes!))
                using (var bmp1 = new Bitmap(ms1))
                {
                    Check(bmp1.Width >= 100 && bmp1.Height >= 100, $"E2E KetcherForm #1 canvas PNG dimensions ({bmp1.Width}x{bmp1.Height}) are valid molecule canvas");
                    Check(bmp1.GetPixel(0, 0).A == 0, "E2E KetcherForm #1 corner pixel is transparent");
                }
                File.WriteAllBytes("output/samples/e2e_ketcher_trans.png", form1.ResultPngBytes!);
                e2eKet = form1.ResultKet;
                Console.WriteLine($"PASS E2E WebView2 Ketcher #1 (Transparent PNG: {form1.ResultPngBytes!.Length} bytes, SMILES: {form1.ResultSmiles})");
            }

            // Live WebView2 + Ketcher E2E Test (2: KET Round-Trip -> White Background Canvas PNG)
            using (var form2 = new KetcherForm("Word", e2eKet!)
            {
                AutoApplyAfterLoadForTest = true,
                ToggleWhiteBackgroundForTest = true
            })
            {
                var dr2 = form2.ShowDialog();
                Check(dr2 == System.Windows.Forms.DialogResult.OK, "E2E KetcherForm #2 returned DialogResult.OK");
                Check(!string.IsNullOrWhiteSpace(form2.ResultKet) && form2.ResultKet!.Contains("\"root\""), "E2E KetcherForm #2 round-tripped KET JSON");
                Check(form2.ResultBackground == "white", "E2E KetcherForm #2 background toggled to white");
                Check(form2.ResultPngBytes != null && form2.ResultPngBytes.Length > 500, "E2E KetcherForm #2 captured white-bg canvas PNG bytes");
                using (var ms2 = new MemoryStream(form2.ResultPngBytes!))
                using (var bmp2 = new Bitmap(ms2))
                {
                    var corner = bmp2.GetPixel(0, 0);
                    Check(corner.A == 255 && corner.R == 255 && corner.G == 255 && corner.B == 255, "E2E KetcherForm #2 corner pixel is opaque white");
                }
                File.WriteAllBytes("output/samples/e2e_ketcher_white.png", form2.ResultPngBytes!);
                Console.WriteLine($"PASS E2E WebView2 Ketcher #2 (KET Round-Trip + White BG PNG: {form2.ResultPngBytes!.Length} bytes)");
            }

            // Live WebView2 + Ketcher E2E Test (3: Complex Molecule Caffeine -> KET & Molecular Properties)
            using (var form3 = new KetcherForm("Excel", "Cn1cnc2c1c(=O)n(C)c(=O)n2C")
            {
                AutoApplyAfterLoadForTest = true,
                ToggleWhiteBackgroundForTest = false
            })
            {
                var dr3 = form3.ShowDialog();
                Check(dr3 == System.Windows.Forms.DialogResult.OK, "E2E KetcherForm #3 returned DialogResult.OK");
                Check(!string.IsNullOrWhiteSpace(form3.ResultKet) && form3.ResultKet!.Contains("\"root\""), "E2E KetcherForm #3 returned valid KET JSON");
                var caffeineInfo = MoleculeRenderer.TryGetMoleculeInfo(form3.ResultKet!);
                Check(caffeineInfo != null && caffeineInfo.Formula == "C8H10N4O2", $"E2E KetcherForm #3 caffeine formula is C8H10N4O2, got: {caffeineInfo?.Formula}");
                Check(Math.Abs((caffeineInfo?.MolecularWeight ?? 0) - 194.19) < 0.2, $"E2E KetcherForm #3 caffeine MW ~ 194.19, got: {caffeineInfo?.MolecularWeight}");
                Check(form3.ResultPngBytes != null && form3.ResultPngBytes.Length > 500, "E2E KetcherForm #3 captured canvas PNG bytes");
                File.WriteAllBytes("output/samples/e2e_ketcher_caffeine.png", form3.ResultPngBytes!);
                Console.WriteLine($"PASS E2E WebView2 Ketcher #3 (Caffeine KET + Formula {caffeineInfo?.Formula} + MW {caffeineInfo?.MolecularWeight:F2})");
            }

            // In-Cell formula generation & escaping tests
            var testSmiles = "CC(=O)Oc1ccccc1C(=O)O";
            var escapedSmiles = testSmiles.Replace("\"", "\"\"");
            var smilesFormula = $"=SMILES2IMG(\"{escapedSmiles}\")";
            Check(smilesFormula == "=SMILES2IMG(\"CC(=O)Oc1ccccc1C(=O)O\")", "SMILES in-cell formula generation");

            var compactSampleKet = StructureMetadata.CompactKet(sampleKet);
            var escapedKet = compactSampleKet.Replace("\"", "\"\"");
            var ketFormula = $"=KET2IMG(\"{escapedKet}\")";
            Check(ketFormula.StartsWith("=KET2IMG(\"") && ketFormula.Contains("\"\"root\"\"") && !ketFormula.Contains("\n"),
                "KET in-cell formula generation with compact JSON and escaped double quotes");
            Console.WriteLine("PASS In-Cell formula generation and JSON escaping");

            // Office Ribbon & Context Menu XML and reflection callback validation
            var excelRibbon = new SmilesRibbon();
            var excelXml = excelRibbon.GetCustomUI("");
            ValidateRibbonXml(excelXml, typeof(SmilesRibbon), "Excel");
            var excelDoc = XDocument.Parse(excelXml);
            Check(excelDoc.Descendants().Any(e => (string?)e.Attribute("id") == "menuContextExcelCell" &&
                                                   (string?)e.Attribute("insertAfterMso") == "PasteGalleryMini"),
                "Excel ContextMenuCell contains menuContextExcelCell anchored after PasteGalleryMini");
            Check(excelDoc.Descendants().Any(e => (string?)e.Attribute("id") == "sepContextCellPre" &&
                                                   (string?)e.Attribute("insertAfterMso") == "PasteGalleryMini"),
                "Excel ContextMenuCell contains sepContextCellPre anchored after PasteGalleryMini");
            Check(excelDoc.Descendants().Any(e => (string?)e.Attribute("id") == "sepContextMolecule" &&
                                                   (string?)e.Attribute("insertAfterMso") == "PasteGalleryMini"),
                "Excel ContextMenuCell contains sepContextMolecule anchored after PasteGalleryMini");
            Check(excelDoc.Descendants().Any(e => (string?)e.Attribute("id") == "menuContextExcelList" &&
                                                   (string?)e.Attribute("insertAfterMso") == "PasteGalleryMini"),
                "Excel ContextMenuListRange contains menuContextExcelList anchored after PasteGalleryMini");
            Check(excelDoc.Descendants().Any(e => (string?)e.Attribute("idMso") == "ContextMenuPicture" &&
                                                   e.Descendants().Any(b => (string?)b.Attribute("insertBeforeMso") == "Cut")),
                "Excel ContextMenuPicture has items anchored before Cut");
            Console.WriteLine("PASS Excel Ribbon & Context Menu XML and callbacks reflection");

            var wordAddIn = new WordAddIn();
            var wordXml = wordAddIn.GetCustomUI("");
            ValidateRibbonXml(wordXml, typeof(WordAddIn), "Word");
            var wordDoc = XDocument.Parse(wordXml);
            Check(wordDoc.Descendants().Any(e => (string?)e.Attribute("id") == "menuContextWordText" &&
                                                 (string?)e.Attribute("insertAfterMso") == "PasteGalleryMini"),
                "Word ContextMenuText contains menuContextWordText anchored after PasteGalleryMini");
            Check(wordDoc.Descendants().Any(e => (string?)e.Attribute("id") == "sepContextText" &&
                                                 (string?)e.Attribute("insertAfterMso") == "PasteGalleryMini"),
                "Word ContextMenuText contains sepContextText anchored after PasteGalleryMini");
            Check(wordDoc.Descendants().Any(e => (string?)e.Attribute("idMso") == "ContextMenuPicture" &&
                                                 e.Descendants().Any(b => (string?)b.Attribute("insertBeforeMso") == "Cut")),
                "Word ContextMenuPicture has items anchored before Cut");
            Console.WriteLine("PASS Word Ribbon & Context Menu XML and callbacks reflection");

            var pptAddIn = new PowerPointAddIn();
            var pptXml = pptAddIn.GetCustomUI("");
            ValidateRibbonXml(pptXml, typeof(PowerPointAddIn), "PowerPoint");
            var pptDoc = XDocument.Parse(pptXml);
            Check(pptDoc.Descendants().Any(e => (string?)e.Attribute("id") == "menuContextPptFrame" &&
                                                (string?)e.Attribute("insertAfterMso") == "PasteGalleryMini"),
                "PowerPoint ContextMenuFrame contains menuContextPptFrame anchored after PasteGalleryMini");
            Check(pptDoc.Descendants().Any(e => (string?)e.Attribute("id") == "menuContextPptText" &&
                                                (string?)e.Attribute("insertAfterMso") == "PasteGalleryMini"),
                "PowerPoint ContextMenuText contains menuContextPptText anchored after PasteGalleryMini");
            Check(pptDoc.Descendants().Any(e => (string?)e.Attribute("id") == "menuContextPptTable" &&
                                                (string?)e.Attribute("insertAfterMso") == "PasteGalleryMini"),
                "PowerPoint ContextMenuTable contains menuContextPptTable anchored after PasteGalleryMini");
            Check(pptDoc.Descendants().Any(e => (string?)e.Attribute("idMso") == "ContextMenuPicture" &&
                                                e.Descendants().Any(b => (string?)b.Attribute("insertBeforeMso") == "Cut")),
                "PowerPoint ContextMenuPicture has items anchored before Cut");
            Console.WriteLine("PASS PowerPoint Ribbon & Context Menu XML and callbacks reflection");

            // Validate AboutDialog instantiation and third-party notices
            using (var aboutDlg = new Smiles2Img.Editor.AboutDialog())
            {
                Check(aboutDlg is not null, "AboutDialog successfully instantiated");
                Check(!string.IsNullOrWhiteSpace(aboutDlg!.Text), "AboutDialog has a window title");
            }
            Console.WriteLine("PASS AboutDialog UI initialization and third-party notices");

            // Validate UiIcons handles
            Check(Smiles2Img.Editor.UiIcons.MoleculeInfoIcon != null && Smiles2Img.Editor.UiIcons.MoleculeInfoIcon.Handle != IntPtr.Zero, "UiIcons.MoleculeInfoIcon has valid handle");
            Check(Smiles2Img.Editor.UiIcons.AboutIcon != null && Smiles2Img.Editor.UiIcons.AboutIcon.Handle != IntPtr.Zero, "UiIcons.AboutIcon has valid handle");
            Check(Smiles2Img.Editor.UiIcons.ChemicalIcon != null && Smiles2Img.Editor.UiIcons.ChemicalIcon.Handle != IntPtr.Zero, "UiIcons.ChemicalIcon has valid handle");
            Check(Smiles2Img.Editor.UiIcons.Molstar3DIcon != null && Smiles2Img.Editor.UiIcons.Molstar3DIcon.Handle != IntPtr.Zero, "UiIcons.Molstar3DIcon has valid handle");

            // Validate CCO MoleculeInfo and MoleculeInfoDialog instantiation
            var ccoInfo = MoleculeRenderer.TryGetMoleculeInfo("CCO", resolvePubChem: true);
            Check(ccoInfo != null, "CCO MoleculeInfo resolved");
            Check(ccoInfo!.Formula == "C2H6O", $"CCO Formula == C2H6O, got: {ccoInfo.Formula}");
            using (var ccoDlg = new Smiles2Img.Editor.MoleculeInfoDialog(ccoInfo))
            {
                Check(ccoDlg is not null, "MoleculeInfoDialog for CCO successfully instantiated");
                Check(ccoDlg.Icon != null && ccoDlg.Icon!.Handle != IntPtr.Zero, "MoleculeInfoDialog icon has valid Win32 handle");
            }
            string ccoSummary = ccoInfo.FormatClipboardSummary();
            Check(!string.IsNullOrWhiteSpace(ccoSummary), "CCO summary formatted");
            var (ethName, ethIupac) = Smiles2Img.Rendering.PubChemResolver.ResolveCompound("CCO");
            Check(ethName == "Ethanol" && ethIupac == "ethanol", $"ResolveCompound for CCO returned {ethName}, {ethIupac}");
            Check(Smiles2Img.Editor.SafeClipboard.SetText(ccoSummary), "SafeClipboard.SetText for CCO summary");
            Check(Smiles2Img.Editor.SafeClipboard.ContainsText(), "SafeClipboard.ContainsText succeeds");
            Check(Smiles2Img.Editor.SafeClipboard.GetText() == ccoSummary, "SafeClipboard.GetText matches ccoSummary");
            Console.WriteLine("PASS UiIcons, SafeClipboard, and CCO MoleculeInfoDialog initialization");

            // v1.7.0 Tests: Exact Mass, Physicochemical & Drug-Likeness Properties
            Check(aspirinInfo.LogP > 0.5 && aspirinInfo.LogP < 2.5, $"Aspirin LogP in expected range, got: {aspirinInfo.LogP}");
            Check(aspirinInfo.Tpsa > 50 && aspirinInfo.Tpsa < 80, $"Aspirin TPSA in expected range, got: {aspirinInfo.Tpsa}");
            Check(aspirinInfo.HBondDonors == 1, $"Aspirin HBD == 1, got: {aspirinInfo.HBondDonors}");
            Check(aspirinInfo.HBondAcceptors >= 3, $"Aspirin HBA >= 3, got: {aspirinInfo.HBondAcceptors}");
            Check(aspirinInfo.RotatableBonds >= 2, $"Aspirin RotBonds >= 2, got: {aspirinInfo.RotatableBonds}");
            Check(aspirinInfo.PassesLipinski == true, "Aspirin passes Lipinski Rule of 5");

            Check(Functions.SmilesExactMass("CCO") is double em1 && Math.Abs(em1 - 46.0419) < 0.01, "SMILES2EXACTMASS for CCO");
            Check(Functions.InchiExactMass(ethanolInchi!) is double em2 && Math.Abs(em2 - 46.0419) < 0.01, "INCHI2EXACTMASS for ethanol");
            Check(Functions.KetExactMass(sampleKet) is double em3 && Math.Abs(em3 - 32.0262) < 0.01, "KET2EXACTMASS for methanol KET");
            Check(Functions.SmilesLogP("CC(=O)Oc1ccccc1C(=O)O") is double lp && lp > 0.5, "SMILES2LOGP for Aspirin");
            Check(Functions.SmilesTpsa("CC(=O)Oc1ccccc1C(=O)O") is double tpsa && tpsa > 50, "SMILES2TPSA for Aspirin");
            Check(Equals(Functions.SmilesHbd("CC(=O)Oc1ccccc1C(=O)O"), 1), "SMILES2HBD for Aspirin == 1");
            Check(Functions.SmilesHba("CC(=O)Oc1ccccc1C(=O)O") is int hba && hba >= 3, "SMILES2HBA for Aspirin >= 3");
            Check(Functions.SmilesRotBonds("CC(=O)Oc1ccccc1C(=O)O") is int rb && rb >= 2, "SMILES2ROTBONDS for Aspirin >= 2");
            Check(Equals(Functions.SmilesLipinski("CC(=O)Oc1ccccc1C(=O)O"), "PASS"), "SMILES2LIPINSKI for Aspirin == PASS");
            Console.WriteLine("PASS v1.7.0 Exact Mass, LogP, TPSA, HBD, HBA, RotBonds, and Lipinski Rule of 5");

            // v1.7.0 Tests: Substructure Matching, Similarity, and Substructure Highlighting
            Check(Equals(Functions.SmilesMatch("CC(=O)Oc1ccccc1C(=O)O", "c1ccccc1"), true), "SMILES.MATCH finds benzene in Aspirin");
            Check(Equals(Functions.SmilesMatch("CCO", "c1ccccc1"), false), "SMILES.MATCH returns false for benzene in Ethanol");
            Check(Functions.SmilesSimilarity("c1ccccc1", "Cc1ccccc1", "tanimoto") is double sim && sim > 0.3 && sim <= 1.0, "SMILES.SIMILARITY Tanimoto between benzene and toluene");
            Check(Functions.SmilesSimilarity("CCO", "CCO", "tanimoto") is double simSelf && Math.Abs(simSelf - 1.0) < 1e-6, "SMILES.SIMILARITY identical molecule == 1.0");
            Check(Functions.SmilesSimilarity("c1ccccc1", "Cc1ccccc1", "tversky") is double simTv && simTv > 0.3 && simTv <= 1.0, "SMILES.SIMILARITY Tversky");

            var hlOpts = Functions.ParseOptions(null, "hl:c1ccccc1,num", null);
            Check(hlOpts.HighlightQuery == "c1ccccc1" && hlOpts.AtomNumbers == true, "ParseOptions extracts hl:c1ccccc1 and num");
            var hlPng = MoleculeRenderer.Render("CC(=O)Oc1ccccc1C(=O)O", hlOpts);
            Check(hlPng != null && hlPng.Length > 200, "Substructure highlighted PNG rendered");
            Console.WriteLine("PASS v1.7.0 SMILES.MATCH, SMILES.SIMILARITY, and Substructure Highlighting");

            // v1.7.0 Tests: Reverse Name / CAS-to-Structure Resolution
            Check(Functions.NameSmiles("Aspirin") is string aspSmiles && !string.IsNullOrEmpty(aspSmiles), "NAME2SMILES resolves Aspirin");
            Check(Functions.NameSmiles("50-78-2") is string casSmiles && !string.IsNullOrEmpty(casSmiles), "NAME2SMILES resolves CAS 50-78-2");
            Check(Functions.NameSmiles("Caffeine") is string cafSmiles && !string.IsNullOrEmpty(cafSmiles), "NAME2SMILES resolves Caffeine");
            Console.WriteLine("PASS v1.7.0 NAME2SMILES & CAS-to-SMILES resolution");

            // v1.7.0 Tests: Vector SVG Rendering & Clipboard Helper
            var svgXml = MoleculeRenderer.RenderSvg("CC(=O)Oc1ccccc1C(=O)O");
            Check(!string.IsNullOrEmpty(svgXml) && svgXml!.Contains("<svg") && svgXml.Contains("</svg>"), "MoleculeRenderer.RenderSvg generates valid SVG XML");
            Check(SvgClipboardHelper.CopySvgToClipboard(svgXml), "SvgClipboardHelper copies SVG to Windows Clipboard");
            Console.WriteLine("PASS v1.7.0 Vector SVG rendering & clipboard helper");

            // v1.7.0 Tests: SDF File Import / Export Round-Trip
            var sdfInput = new[]
            {
                new SdfRecord { Smiles = "CCO", Name = "Ethanol", Properties = new System.Collections.Generic.Dictionary<string, string> { ["Source"] = "UnitTest" } },
                new SdfRecord { Smiles = "CC(=O)Oc1ccccc1C(=O)O", Name = "Aspirin" }
            };
            var sdfBytes = SdfManager.WriteSdfBytes(sdfInput);
            Check(sdfBytes != null && sdfBytes.Length > 100, "SdfManager.WriteSdfBytes produces non-empty SDF bytes");
            var sdfRecords = SdfManager.ReadSdfBytes(sdfBytes!);
            Check(sdfRecords.Count == 2, $"SdfManager.ReadSdfBytes read 2 records, got: {sdfRecords.Count}");
            Check(sdfRecords[0].Formula == "C2H6O", $"SDF record 0 formula is C2H6O, got: {sdfRecords[0].Formula}");
            Check(sdfRecords[1].Formula == "C9H8O4", $"SDF record 1 formula is C9H8O4, got: {sdfRecords[1].Formula}");
            Console.WriteLine("PASS v1.7.0 SDF Import / Export round-trip");

            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }

    private static void ValidateRibbonXml(string xml, Type addInType, string appName)
    {
        Check(!string.IsNullOrWhiteSpace(xml), $"{appName} Ribbon XML is non-empty");
        var doc = XDocument.Parse(xml);
        Check(doc.Root != null, $"{appName} XML has root element");

        var onActions = doc.Descendants()
            .Select(e => (string?)e.Attribute("onAction"))
            .Where(a => !string.IsNullOrEmpty(a))
            .Distinct()
            .ToList();

        foreach (var action in onActions)
        {
            var method = addInType.GetMethod(action!, BindingFlags.Public | BindingFlags.Instance);
            Check(method != null, $"{appName} onAction callback '{action}' must exist on {addInType.Name}");
        }

        var allIds = doc.Descendants()
            .Select(e => (string?)e.Attribute("id"))
            .Where(id => !string.IsNullOrEmpty(id))
            .ToList();
        var duplicates = allIds.GroupBy(x => x).Where(g => g.Count() > 1).Select(g => g.Key).ToList();
        Check(duplicates.Count == 0, $"{appName} duplicate control IDs: {string.Join(", ", duplicates)}");
    }

    private static void Check(bool condition, string message)
    {
        if (!condition) throw new Exception("Failed: " + message);
    }
}
