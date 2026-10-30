using System.Reflection;
using Smiles2Img.Editor;

namespace Smiles2Img;

public static class MolstarAssetManager
{
    private const string ResourceName = "Smiles2Img.Molstar.zip";
    private static string? _cachedFolder;

    public static string EnsureMolstarFolder() =>
        EmbeddedAssetExtractor.EnsureExtracted(ref _cachedFolder, "molstar_bundle", "molstar", ResourceName);

    public static string GetBundleFingerprint(Assembly? assembly = null) =>
        EmbeddedAssetExtractor.GetBundleFingerprint(ResourceName, assembly);
}
