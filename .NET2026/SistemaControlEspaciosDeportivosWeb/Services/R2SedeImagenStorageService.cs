using Amazon;
using Amazon.Runtime;
using Amazon.S3;
using Amazon.S3.Model;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Options;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Webp;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Processing;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public class R2SedeImagenStorageService(IOptions<SedeImagenStorageSettings> options) : ISedeImagenStorageService
{
    private static readonly HashSet<string> ExtensionesPermitidas = new(StringComparer.OrdinalIgnoreCase)
    {
        ".jpg", ".jpeg", ".png"
    };

    private readonly SedeImagenStorageSettings _settings = options.Value;

    public async Task<(bool Ok, string Mensaje)> ProbarConexionAsync(CancellationToken cancellationToken = default)
    {
        if (!_settings.Enabled)
            return (false, "SedeImagenStorage esta deshabilitado. Activa SedeImagenStorage:Enabled=true.");

        if (string.IsNullOrWhiteSpace(_settings.Endpoint) ||
            string.IsNullOrWhiteSpace(_settings.BucketName) ||
            string.IsNullOrWhiteSpace(_settings.AccessKey) ||
            string.IsNullOrWhiteSpace(_settings.SecretKey))
            return (false, "Configuracion incompleta de SedeImagenStorage (Endpoint/BucketName/AccessKey/SecretKey).");

        var accessKey = _settings.AccessKey.Trim();
        var bucketName = _settings.BucketName.Trim();
        var keyDebug = accessKey.Length >= 4 ? $"***{accessKey[^4..]}" : "***";

        var (okActual, mensajeActual, invalidaActual) = await ProbarEndpointAsync(_settings.Endpoint, bucketName, cancellationToken);
        if (okActual)
        {
            var (okWriteActual, writeMsgActual) = await ProbarEscrituraAsync(_settings.Endpoint, bucketName, cancellationToken);
            if (okWriteActual)
                return (true, $"Conexion OK a R2 (lectura/escritura). Bucket: {bucketName}. Endpoint: {_settings.Endpoint}. Key: {keyDebug}. {mensajeActual}");

            return (false, $"Conexion de lectura OK pero falla escritura en R2. Endpoint: {_settings.Endpoint}. Bucket: {bucketName}. Detalle: {writeMsgActual}");
        }

        var endpointAlterno = ObtenerEndpointAlterno(_settings.Endpoint);
        if (string.IsNullOrWhiteSpace(endpointAlterno) || string.Equals(endpointAlterno, _settings.Endpoint, StringComparison.OrdinalIgnoreCase))
        {
            return invalidaActual
                ? (false, $"La Access Key de R2 no existe o fue revocada. Endpoint: {_settings.Endpoint}. Key actual termina en: {keyDebug}. Detalle: {mensajeActual}")
                : (false, $"No se pudo validar conexion con R2. Endpoint: {_settings.Endpoint}. Bucket: {bucketName}. Error: {mensajeActual}");
        }

        var (okAlt, mensajeAlt, invalidaAlt) = await ProbarEndpointAsync(endpointAlterno, bucketName, cancellationToken);
        if (okAlt)
        {
            return (false, $"Las credenciales son validas, pero tu endpoint configurado no. Fallo en: {_settings.Endpoint}. Funciona en: {endpointAlterno}. Actualiza SedeImagenStorage:Endpoint.");
        }

        if (invalidaActual && invalidaAlt)
        {
            return (false, $"La Access Key de R2 no existe o fue revocada en ambos endpoints (default/EU). Key actual termina en: {keyDebug}. Error: {mensajeActual}");
        }

        return (false, $"No se pudo validar conexion R2. Endpoint actual: {_settings.Endpoint} => {mensajeActual}. Endpoint alterno: {endpointAlterno} => {mensajeAlt}.");
    }

    public async Task<IReadOnlyList<string>> UploadSedeImagenesAsync(int negocioId, int? sedeId, IEnumerable<IFormFile> archivos, CancellationToken cancellationToken = default)
    {
        if (!_settings.Enabled)
            throw new InvalidOperationException("La carga de imagenes esta deshabilitada. Configura SedeImagenStorage:Enabled=true.");

        if (string.IsNullOrWhiteSpace(_settings.Endpoint) ||
            string.IsNullOrWhiteSpace(_settings.BucketName) ||
            string.IsNullOrWhiteSpace(_settings.AccessKey) ||
            string.IsNullOrWhiteSpace(_settings.SecretKey))
            throw new InvalidOperationException("Configuracion incompleta de SedeImagenStorage (Endpoint/BucketName/AccessKey/SecretKey).");

        var archivosValidos = archivos.Where(x => x is not null && x.Length > 0).ToList();
        if (archivosValidos.Count == 0)
            return Array.Empty<string>();

        var accessKey = _settings.AccessKey.Trim();
        var bucketName = _settings.BucketName.Trim();

        using var cliente = CreateClient();
        var urls = new List<string>();

        foreach (var archivo in archivosValidos)
        {
            if (archivo.Length > _settings.MaxImageBytes)
                throw new InvalidOperationException($"La imagen {archivo.FileName} supera el tamaño permitido de {_settings.MaxImageBytes / 1024 / 1024} MB.");

            var extension = Path.GetExtension(archivo.FileName ?? string.Empty);
            if (string.IsNullOrWhiteSpace(extension) || !ExtensionesPermitidas.Contains(extension))
                throw new InvalidOperationException($"Formato no permitido en {archivo.FileName}. Solo se acepta JPG/JPEG o PNG.");

            if (string.IsNullOrWhiteSpace(archivo.ContentType) ||
                !(archivo.ContentType.Equals("image/jpeg", StringComparison.OrdinalIgnoreCase) ||
                  archivo.ContentType.Equals("image/png", StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException($"El archivo {archivo.FileName} no es una imagen JPG/PNG valida.");

            await using var streamOrigen = archivo.OpenReadStream();
            var nombreArchivo = string.IsNullOrWhiteSpace(archivo.FileName) ? "imagen.jpg" : archivo.FileName;
            var imagenProcesada = await ProcesarImagenAsync(streamOrigen, nombreArchivo, cancellationToken);

            var key = BuildObjectKey(negocioId, sedeId, ".webp");
            imagenProcesada.Position = 0;

            var putRequest = new PutObjectRequest
            {
                BucketName = bucketName,
                Key = key,
                InputStream = imagenProcesada,
                ContentType = "image/webp",
                AutoCloseStream = false,
                DisablePayloadSigning = true,
                UseChunkEncoding = false
            };

            try
            {
                await cliente.PutObjectAsync(putRequest, cancellationToken);
            }
            catch (AmazonS3Exception ex) when (ex.Message.Contains("authorization header is malformed", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("No se pudo autenticar con R2. Verifica Endpoint, AccessKey/SecretKey y la region de firma (us-east-1).");
            }
            catch (AmazonS3Exception ex) when (
                string.Equals(ex.ErrorCode, "InvalidAccessKeyId", StringComparison.OrdinalIgnoreCase) ||
                ex.Message.Contains("Access Key Id you provided does not exist", StringComparison.OrdinalIgnoreCase))
            {
                var endpointDebug = _settings.Endpoint;
                var keyDebug = string.IsNullOrWhiteSpace(accessKey)
                    ? "(vacia)"
                    : $"***{accessKey[^4..]}";
                throw new InvalidOperationException(
                    $"La Access Key de R2 no existe o fue revocada. Endpoint actual: {endpointDebug}. Key actual termina en: {keyDebug}. " +
                    "Genera una nueva clave S3 en Cloudflare R2 (no token cfat_) y actualiza User Secrets.");
            }
            catch (AmazonS3Exception ex)
            {
                throw new InvalidOperationException($"No se pudo subir la imagen a R2: {ex.Message}");
            }

            await imagenProcesada.DisposeAsync();
            urls.Add(BuildPublicUrl(key));
        }

        return urls;
    }

    public async Task<IReadOnlyList<string>> DeleteSedeImagenesAsync(IEnumerable<string> urls, CancellationToken cancellationToken = default)
    {
        if (!_settings.Enabled)
            return Array.Empty<string>();

        if (string.IsNullOrWhiteSpace(_settings.Endpoint) ||
            string.IsNullOrWhiteSpace(_settings.BucketName) ||
            string.IsNullOrWhiteSpace(_settings.AccessKey) ||
            string.IsNullOrWhiteSpace(_settings.SecretKey))
            return Array.Empty<string>();

        var keys = (urls ?? Array.Empty<string>())
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Select(x => x.Trim())
            .Select(ExtraerObjectKeyDesdeUrl)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Cast<string>()
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (keys.Count == 0)
            return Array.Empty<string>();

        var eliminadas = new List<string>(keys.Count);
        using var cliente = CreateClient();
        var bucket = _settings.BucketName.Trim();
        foreach (var key in keys)
        {
            try
            {
                await cliente.DeleteObjectAsync(new DeleteObjectRequest
                {
                    BucketName = bucket,
                    Key = key
                }, cancellationToken);
                eliminadas.Add(key);
            }
            catch
            {
                // Se ignora para no romper el flujo de guardado del formulario.
            }
        }

        return eliminadas;
    }

    public async Task<(byte[] Contenido, string ContentType)?> ObtenerImagenVisualizacionAsync(string? url, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(url))
            return null;

        var urlTrim = url.Trim();
        if (!Uri.TryCreate(urlTrim, UriKind.Absolute, out var uriOriginal))
            return null;

        if (!EsUrlDelStorage(uriOriginal))
            return null;

        var key = ExtraerObjectKeyDesdeUrl(urlTrim);
        if (string.IsNullOrWhiteSpace(key))
            return null;

        if (!_settings.Enabled ||
            string.IsNullOrWhiteSpace(_settings.Endpoint) ||
            string.IsNullOrWhiteSpace(_settings.BucketName) ||
            string.IsNullOrWhiteSpace(_settings.AccessKey) ||
            string.IsNullOrWhiteSpace(_settings.SecretKey))
            return null;

        try
        {
            using var cliente = CreateClient();
            using var response = await cliente.GetObjectAsync(new GetObjectRequest
            {
                BucketName = _settings.BucketName.Trim(),
                Key = key
            }, cancellationToken);

            await using var source = response.ResponseStream;
            using var buffer = new MemoryStream();
            await source.CopyToAsync(buffer, cancellationToken);

            var contentType = string.IsNullOrWhiteSpace(response.Headers.ContentType)
                ? "image/webp"
                : response.Headers.ContentType;

            return (buffer.ToArray(), contentType);
        }
        catch
        {
            return null;
        }
    }

    private async Task<(bool Ok, string Mensaje, bool AccessKeyInvalida)> ProbarEndpointAsync(string endpoint, string bucketName, CancellationToken cancellationToken)
    {
        try
        {
            using var cliente = CreateClient(endpoint);
            var response = await cliente.ListObjectsV2Async(new ListObjectsV2Request
            {
                BucketName = bucketName,
                MaxKeys = 1
            }, cancellationToken);
            return (true, $"Objetos listados: {response.KeyCount}.", false);
        }
        catch (AmazonS3Exception ex) when (
            string.Equals(ex.ErrorCode, "InvalidAccessKeyId", StringComparison.OrdinalIgnoreCase) ||
            ex.Message.Contains("Access Key Id you provided does not exist", StringComparison.OrdinalIgnoreCase))
        {
            return (false, ex.Message, true);
        }
        catch (AmazonS3Exception ex)
        {
            return (false, ex.Message, false);
        }
        catch (Exception ex)
        {
            return (false, ex.Message, false);
        }
    }

    private async Task<(bool Ok, string Mensaje)> ProbarEscrituraAsync(string endpoint, string bucketName, CancellationToken cancellationToken)
    {
        var keyPrueba = $"diagnostico/test-{DateTime.UtcNow:yyyyMMddHHmmssfff}-{Guid.NewGuid():N}.txt";
        var bytes = System.Text.Encoding.UTF8.GetBytes("ok");
        try
        {
            using var cliente = CreateClient(endpoint);
            await using var ms = new MemoryStream(bytes);
            var put = new PutObjectRequest
            {
                BucketName = bucketName,
                Key = keyPrueba,
                InputStream = ms,
                ContentType = "text/plain",
                AutoCloseStream = false,
                DisablePayloadSigning = true,
                UseChunkEncoding = false
            };
            await cliente.PutObjectAsync(put, cancellationToken);

            await cliente.DeleteObjectAsync(new DeleteObjectRequest
            {
                BucketName = bucketName,
                Key = keyPrueba
            }, cancellationToken);

            return (true, "Escritura y borrado de prueba OK.");
        }
        catch (AmazonS3Exception ex)
        {
            return (false, ex.Message);
        }
        catch (Exception ex)
        {
            return (false, ex.Message);
        }
    }

    private static string ObtenerEndpointAlterno(string endpoint)
    {
        if (string.IsNullOrWhiteSpace(endpoint))
            return endpoint;

        if (endpoint.Contains(".eu.r2.cloudflarestorage.com", StringComparison.OrdinalIgnoreCase))
            return endpoint.Replace(".eu.r2.cloudflarestorage.com", ".r2.cloudflarestorage.com", StringComparison.OrdinalIgnoreCase);

        if (endpoint.Contains(".r2.cloudflarestorage.com", StringComparison.OrdinalIgnoreCase) &&
            !endpoint.Contains(".eu.r2.cloudflarestorage.com", StringComparison.OrdinalIgnoreCase))
            return endpoint.Replace(".r2.cloudflarestorage.com", ".eu.r2.cloudflarestorage.com", StringComparison.OrdinalIgnoreCase);

        return endpoint;
    }

    private bool EsUrlDelStorage(Uri uri)
    {
        if (!string.IsNullOrWhiteSpace(_settings.PublicBaseUrl) &&
            Uri.TryCreate(_settings.PublicBaseUrl.Trim(), UriKind.Absolute, out var publicUri) &&
            string.Equals(publicUri.Host, uri.Host, StringComparison.OrdinalIgnoreCase))
            return true;

        if (!string.IsNullOrWhiteSpace(_settings.Endpoint))
        {
            var endpoint = _settings.Endpoint.Trim();
            if (!endpoint.StartsWith("https://", StringComparison.OrdinalIgnoreCase) &&
                !endpoint.StartsWith("http://", StringComparison.OrdinalIgnoreCase))
                endpoint = $"https://{endpoint}";

            if (Uri.TryCreate(endpoint, UriKind.Absolute, out var endpointUri) &&
                string.Equals(endpointUri.Host, uri.Host, StringComparison.OrdinalIgnoreCase))
                return true;
        }

        var bucket = _settings.BucketName.Trim().Trim('/');
        return !string.IsNullOrWhiteSpace(bucket) &&
               uri.AbsolutePath.TrimStart('/').StartsWith(bucket + "/", StringComparison.OrdinalIgnoreCase);
    }

    private AmazonS3Client CreateClient(string? endpointOverride = null)
    {
        var servicioUrl = (endpointOverride ?? _settings.Endpoint).Trim();
        if (!servicioUrl.StartsWith("https://", StringComparison.OrdinalIgnoreCase) &&
            !servicioUrl.StartsWith("http://", StringComparison.OrdinalIgnoreCase))
            servicioUrl = $"https://{servicioUrl}";

        var accessKey = _settings.AccessKey.Trim();
        var secretKey = _settings.SecretKey.Trim();
        var region = string.IsNullOrWhiteSpace(_settings.Region) ? "us-east-1" : _settings.Region.Trim();

        var config = new AmazonS3Config
        {
            ServiceURL = servicioUrl,
            ForcePathStyle = true,
            AuthenticationRegion = region,
            SignatureVersion = "4"
        };

        var credenciales = new BasicAWSCredentials(accessKey, secretKey);
        return new AmazonS3Client(credenciales, config);
    }

    private string BuildObjectKey(int negocioId, int? sedeId, string extension)
    {
        var safeExt = extension.StartsWith('.') ? extension.ToLowerInvariant() : $".{extension.ToLowerInvariant()}";
        var sedeSegmento = sedeId.HasValue ? $"sede-{sedeId.Value}" : "sede-nueva";
        return $"sedes/negocio-{negocioId}/{sedeSegmento}/{DateTime.UtcNow:yyyyMMddHHmmssfff}-{Guid.NewGuid():N}{safeExt}";
    }

    private string BuildPublicUrl(string key)
    {
        if (!string.IsNullOrWhiteSpace(_settings.PublicBaseUrl))
            return $"{_settings.PublicBaseUrl.TrimEnd('/')}/{key}";

        var endpoint = _settings.Endpoint.Trim().TrimEnd('/');
        if (!endpoint.StartsWith("https://", StringComparison.OrdinalIgnoreCase) &&
            !endpoint.StartsWith("http://", StringComparison.OrdinalIgnoreCase))
            endpoint = $"https://{endpoint}";

        return $"{endpoint}/{_settings.BucketName}/{key}";
    }

    private string? ExtraerObjectKeyDesdeUrl(string url)
    {
        if (string.IsNullOrWhiteSpace(url))
            return null;

        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri))
            return null;

        var path = uri.AbsolutePath.Trim('/');
        if (string.IsNullOrWhiteSpace(path))
            return null;

        var bucket = _settings.BucketName.Trim().Trim('/');
        if (!string.IsNullOrWhiteSpace(bucket) &&
            path.StartsWith(bucket + "/", StringComparison.OrdinalIgnoreCase))
            return path[(bucket.Length + 1)..];

        if (!string.IsNullOrWhiteSpace(_settings.PublicBaseUrl))
        {
            var publicBase = _settings.PublicBaseUrl.Trim().TrimEnd('/');
            if (url.StartsWith(publicBase + "/", StringComparison.OrdinalIgnoreCase))
                return url[(publicBase.Length + 1)..];
        }

        return path;
    }

    private async Task<MemoryStream> ProcesarImagenAsync(Stream sourceStream, string nombreArchivo, CancellationToken cancellationToken)
    {
        using var image = await Image.LoadAsync<Rgba32>(sourceStream, cancellationToken);
        image.Mutate(ctx => ctx.AutoOrient());

        if (image.Width < image.Height)
            throw new InvalidOperationException($"La imagen {nombreArchivo} debe ser horizontal.");

        var targetWidth = Math.Max(1, _settings.TargetWidth);
        var targetHeight = Math.Max(1, _settings.TargetHeight);
        const decimal ratioObjetivo = 4m / 3m;
        var ratioActual = image.Width / (decimal)image.Height;

        Rectangle cropRect;
        if (ratioActual > ratioObjetivo)
        {
            var cropWidth = (int)Math.Round(image.Height * ratioObjetivo, MidpointRounding.AwayFromZero);
            var x = Math.Max(0, (image.Width - cropWidth) / 2);
            cropRect = new Rectangle(x, 0, cropWidth, image.Height);
        }
        else
        {
            var cropHeight = (int)Math.Round(image.Width / ratioObjetivo, MidpointRounding.AwayFromZero);
            var y = Math.Max(0, (image.Height - cropHeight) / 2);
            cropRect = new Rectangle(0, y, image.Width, cropHeight);
        }

        image.Mutate(ctx =>
        {
            ctx.Crop(cropRect);
            ctx.Resize(new ResizeOptions
            {
                Size = new Size(targetWidth, targetHeight),
                Mode = ResizeMode.Stretch,
                Sampler = KnownResamplers.Lanczos3
            });
        });

        var output = new MemoryStream();
        var maxBytes = Math.Max(1, _settings.MaxOutputBytes);

        for (var quality = 86; quality >= 52; quality -= 6)
        {
            output.SetLength(0);
            output.Position = 0;
            var encoder = new WebpEncoder
            {
                FileFormat = WebpFileFormatType.Lossy,
                Quality = quality
            };
            await image.SaveAsync(output, encoder, cancellationToken);
            if (output.Length <= maxBytes)
            {
                output.Position = 0;
                return output;
            }
        }

        throw new InvalidOperationException($"La imagen {nombreArchivo} no pudo comprimirse a {_settings.MaxOutputBytes / 1024} KB manteniendo la resolucion estandar.");
    }
}
