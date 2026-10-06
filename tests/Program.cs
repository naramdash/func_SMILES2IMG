using System;
using System.Drawing;
using System.IO;
using System.Linq;
using ExcelDna.Integration;
using Smiles2Img.Rendering;

namespace Smiles2Img.Tests;

internal static class Program
{
    [STAThread]
    private static int Main()
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
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }

    private static void Check(bool condition, string message)
    {
        if (!condition) throw new Exception("Failed: " + message);
    }
}
