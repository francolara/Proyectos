namespace SistemaControlEspaciosDeportivosWeb.Services;

public static class SedeImagenUrlHelper
{
    public static string? NormalizarUrlPublica(string? url, string? publicBaseUrl, string? bucketName)
    {
        if (string.IsNullOrWhiteSpace(url))
            return null;

        var raw = url.Trim();
        if (!Uri.TryCreate(raw, UriKind.Absolute, out var uriOriginal))
            return raw;

        if (string.IsNullOrWhiteSpace(publicBaseUrl) || !Uri.TryCreate(publicBaseUrl.Trim(), UriKind.Absolute, out var uriPublica))
            return raw;

        var key = ExtraerObjectKey(uriOriginal, bucketName);
        if (string.IsNullOrWhiteSpace(key))
            return raw;

        return $"{uriPublica.Scheme}://{uriPublica.Host}/{key}";
    }

    private static string? ExtraerObjectKey(Uri uriOriginal, string? bucketName)
    {
        var path = uriOriginal.AbsolutePath.Trim('/');
        if (string.IsNullOrWhiteSpace(path))
            return null;

        var bucket = string.IsNullOrWhiteSpace(bucketName) ? string.Empty : bucketName.Trim().Trim('/');
        if (!string.IsNullOrWhiteSpace(bucket) && path.StartsWith(bucket + "/", StringComparison.OrdinalIgnoreCase))
            return path[(bucket.Length + 1)..];

        return path;
    }
}

