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

            // Verify ParseOptions argument separation: (background, color, transform)
            var opt1 = Functions.ParseOptions(null, null, null);
            Check(opt1.Color == true && opt1.Transform == null && opt1.Background == null, "ParseOptions default");
            var opt2 = Functions.ParseOptions("trans", null, null);
            Check(opt2.Color == true && opt2.Transform == null && opt2.Background == "transparent", "ParseOptions trans 2nd arg");
            var opt3 = Functions.ParseOptions(null, false, null);
            Check(opt3.Color == false && opt3.Transform == null && opt3.Background == null, "ParseOptions B/W 3rd arg");
            var opt4 = Functions.ParseOptions("trans", false, null);
            Check(opt4.Color == false && opt4.Transform == null && opt4.Background == "transparent", "ParseOptions trans + B/W");
            var opt5 = Functions.ParseOptions(null, null, "flip");
            Check(opt5.Color == true && opt5.Transform == "flip" && opt5.Background == null, "ParseOptions flip 4th arg");
            var opt6 = Functions.ParseOptions("trans", false, "flip");
            Check(opt6.Color == false && opt6.Transform == "flip" && opt6.Background == "transparent", "ParseOptions trans + B/W + flip");
            var opt7 = Functions.ParseOptions("#FFFF00", null, "rot90");
            Check(opt7.Color == true && opt7.Transform == "rot90" && opt7.Background == "#FFFF00", "ParseOptions hex bg + rot90");
            var optFallback = Functions.ParseOptions(false, null, null);
            Check(optFallback.Color == false, "ParseOptions fallback bool in 2nd arg");
            Console.WriteLine("PASS ParseOptions (background, color, transform) argument separation");
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

            var aflatoxinPng = MoleculeRenderer.Render("O1C=C[C@H]([C@H]1O2)c3c2cc(OC)c4c3OC(=O)C5=C4CCC(=O)5");
            Check(aflatoxinPng.Length > 100, "Aflatoxin junction H render");
            Console.WriteLine("PASS Ring-junction chiral H rendering");

            var micPng = MoleculeRenderer.Render("CN=C=O");
            Check(micPng.Length > 100, "MIC small molecule render");
            Console.WriteLine("PASS Small molecule render (MIC)");

            var nicFlippedPng = MoleculeRenderer.Render("CN1CCC[C@H]1c2cccnc2", color: true, transform: "flip");
            Check(nicFlippedPng.Length > 100, "Nicotine flip transform render");
            Console.WriteLine("PASS Transform rendering (flip)");

            var transPng = MoleculeRenderer.Render("CCO", color: true, background: "transparent");
            using (var ms = new MemoryStream(transPng))
            using (var bmp = new Bitmap(ms))
            {
                Check(bmp.GetPixel(0, 0).A == 0, "Transparent background corner pixel must have alpha 0");
            }
            Console.WriteLine("PASS Transparent background rendering");

            var yellowPng = MoleculeRenderer.Render("CCO", color: true, background: "#FFFF00");
            using (var ms = new MemoryStream(yellowPng))
            using (var bmp = new Bitmap(ms))
            {
                var corner = bmp.GetPixel(0, 0);
                Check(corner.R > 240 && corner.G > 240 && corner.B < 20, "Yellow background corner pixel");
            }
            Console.WriteLine("PASS Custom background color rendering");

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

            Console.WriteLine("PASS Rendered all 10 reference sample molecules + small molecules");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }

    private static void Check(bool condition, string message)
    {
        if (!condition) throw new Exception("Failed: " + message);
    }
}
