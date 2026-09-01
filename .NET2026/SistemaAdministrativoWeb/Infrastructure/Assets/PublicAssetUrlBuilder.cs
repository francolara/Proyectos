using Microsoft.Extensions.Configuration;

namespace SistemaAdministrativoWeb.Infrastructure.Assets;

public static class PublicAssetUrlBuilder
{
    private const string PublicBaseUrlKey = "FRALSECONT_SedeImagenStorage:PublicBaseUrl";
    private const string AssetVersionKey = "SedeImagenStorage:AssetVersion";

    public static string Build(IConfiguration configuration, string relativePath)
    {
        ArgumentNullException.ThrowIfNull(configuration);

        var baseUrl = configuration[PublicBaseUrlKey]?.Trim().TrimEnd('/');
        var normalizedPath = (relativePath ?? string.Empty).Trim().TrimStart('/');
        if (string.IsNullOrWhiteSpace(baseUrl) || string.IsNullOrWhiteSpace(normalizedPath))
        {
            return string.Empty;
        }

        var assetUrl = $"{baseUrl}/{normalizedPath}";
        var assetVersion = configuration[AssetVersionKey]?.Trim();
        return string.IsNullOrWhiteSpace(assetVersion)
            ? assetUrl
            : $"{assetUrl}?v={Uri.EscapeDataString(assetVersion)}";
    }
}
