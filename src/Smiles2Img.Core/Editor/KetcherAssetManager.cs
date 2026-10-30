using System.Reflection;
using Smiles2Img.Editor;

namespace Smiles2Img;

public static class KetcherAssetManager
{
    private const string ResourceName = "Smiles2Img.Ketcher.zip";
    private static string? _cachedFolder;

    public static string EnsureKetcherFolder() =>
        EmbeddedAssetExtractor.EnsureExtracted(ref _cachedFolder, "ketcher_bundle", "ketcher", ResourceName);

    public static string GetBundleFingerprint(Assembly? assembly = null) =>
        EmbeddedAssetExtractor.GetBundleFingerprint(ResourceName, assembly);
}
