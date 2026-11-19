using SkiaSharp;

namespace SistemaControlEspaciosDeportivosWeb.Services;

internal static class ProcesamientoImagen
{
    public static SKCodec CrearCodec(Stream stream, string nombreArchivo, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return SKCodec.Create(new SKManagedStream(stream, disposeManagedStream: false))
            ?? throw new InvalidOperationException($"No se pudo leer la imagen {nombreArchivo}.");
    }

    public static SKBitmap Decodificar(SKCodec codec, string nombreArchivo)
    {
        var info = new SKImageInfo(codec.Info.Width, codec.Info.Height, SKColorType.Rgba8888, SKAlphaType.Premul);
        var imagen = new SKBitmap(info);
        try
        {
            if (codec.GetPixels(info, imagen.GetPixels()) != SKCodecResult.Success)
                throw new InvalidOperationException($"No se pudo leer la imagen {nombreArchivo}.");

            return imagen;
        }
        catch
        {
            imagen.Dispose();
            throw;
        }
    }

    public static SKSizeI ObtenerDimensiones(SKCodec codec) =>
        codec.EncodedOrigin is SKEncodedOrigin.LeftTop or SKEncodedOrigin.RightTop
            or SKEncodedOrigin.RightBottom or SKEncodedOrigin.LeftBottom
            ? new SKSizeI(codec.Info.Height, codec.Info.Width)
            : new SKSizeI(codec.Info.Width, codec.Info.Height);

    public static SKBitmap? Ajustar(SKBitmap imagen, SKEncodedOrigin orientacion, SKSizeI dimensiones, int ancho, int alto)
    {
        if (orientacion == SKEncodedOrigin.TopLeft && dimensiones.Width == ancho && dimensiones.Height == alto)
            return null;

        var proporcion = ancho / (decimal)alto;
        var anchoRecorte = Math.Clamp((int)Math.Round(dimensiones.Height * proporcion, MidpointRounding.AwayFromZero), 1, dimensiones.Width);
        var altoRecorte = Math.Clamp((int)Math.Round(dimensiones.Width / proporcion, MidpointRounding.AwayFromZero), 1, dimensiones.Height);
        var izquierda = (dimensiones.Width - anchoRecorte) / 2;
        var arriba = (dimensiones.Height - altoRecorte) / 2;

        var resultado = new SKBitmap(ancho, alto, SKColorType.Rgba8888, SKAlphaType.Premul);
        try
        {
            using var canvas = new SKCanvas(resultado);
            using var origen = SKImage.FromBitmap(imagen);
            canvas.Clear(SKColors.Transparent);
            canvas.Scale(ancho / (float)anchoRecorte, alto / (float)altoRecorte);
            canvas.Translate(-izquierda, -arriba);
            canvas.Concat(MatrizOrientacion(orientacion, imagen.Width, imagen.Height));
            canvas.DrawImage(origen, 0, 0, new SKSamplingOptions(SKCubicResampler.Mitchell));
            return resultado;
        }
        catch
        {
            resultado.Dispose();
            throw;
        }
    }

    private static SKMatrix MatrizOrientacion(SKEncodedOrigin orientacion, int ancho, int alto) =>
        orientacion switch
        {
            SKEncodedOrigin.TopRight => new SKMatrix(-1, 0, ancho, 0, 1, 0, 0, 0, 1),
            SKEncodedOrigin.BottomRight => new SKMatrix(-1, 0, ancho, 0, -1, alto, 0, 0, 1),
            SKEncodedOrigin.BottomLeft => new SKMatrix(1, 0, 0, 0, -1, alto, 0, 0, 1),
            SKEncodedOrigin.LeftTop => new SKMatrix(0, 1, 0, 1, 0, 0, 0, 0, 1),
            SKEncodedOrigin.RightTop => new SKMatrix(0, -1, alto, 1, 0, 0, 0, 0, 1),
            SKEncodedOrigin.RightBottom => new SKMatrix(0, -1, alto, -1, 0, ancho, 0, 0, 1),
            SKEncodedOrigin.LeftBottom => new SKMatrix(0, 1, 0, -1, 0, ancho, 0, 0, 1),
            _ => SKMatrix.Identity
        };
}
