using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using SistemaVisual.Models;

namespace SistemaVisual.Services
{
    public sealed class DownloadService : IDisposable
    {
        private readonly HttpClient cliente;
        private readonly int maximoReintentos;
        private readonly LogService log;

        public DownloadService(int timeoutSegundos, int maximoReintentos, LogService log)
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            cliente = new HttpClient { Timeout = TimeSpan.FromSeconds(timeoutSegundos) };
            cliente.DefaultRequestHeaders.UserAgent.ParseAdd("SistemaVisual-Actualizador/2.0");
            this.maximoReintentos = maximoReintentos;
            this.log = log;
        }

        public async Task<RemoteFileList> ObtenerListadoAsync(string url, CancellationToken cancellationToken)
        {
            return await EjecutarConReintentosAsync(async () =>
            {
                using (var respuesta = await cliente.GetAsync(url, HttpCompletionOption.ResponseContentRead, cancellationToken))
                {
                    respuesta.EnsureSuccessStatusCode();
                    var json = await respuesta.Content.ReadAsStringAsync();
                    cancellationToken.ThrowIfCancellationRequested();
                    try
                    {
                        return JsonConvert.DeserializeObject<RemoteFileList>(json);
                    }
                    catch (JsonException ex)
                    {
                        throw new InvalidOperationException("El Worker devolvió una respuesta JSON inválida.", ex);
                    }
                }
            }, "consulta al Worker", cancellationToken);
        }

        public async Task DescargarArchivoAsync(
            string url,
            string destino,
            long tamanoEsperado,
            Action<long, long> informarProgreso,
            CancellationToken cancellationToken)
        {
            await EjecutarConReintentosAsync(async () =>
            {
                if (File.Exists(destino))
                    File.Delete(destino);
                Directory.CreateDirectory(Path.GetDirectoryName(destino));

                using (var respuesta = await cliente.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, cancellationToken))
                {
                    respuesta.EnsureSuccessStatusCode();
                    var total = respuesta.Content.Headers.ContentLength ?? tamanoEsperado;
                    long descargado = 0;
                    using (var origen = await respuesta.Content.ReadAsStreamAsync())
                    using (var destinoStream = new FileStream(destino, FileMode.CreateNew, FileAccess.Write, FileShare.None, 81920, true))
                    {
                        var buffer = new byte[81920];
                        int leidos;
                        while ((leidos = await origen.ReadAsync(buffer, 0, buffer.Length, cancellationToken)) > 0)
                        {
                            await destinoStream.WriteAsync(buffer, 0, leidos, cancellationToken);
                            descargado += leidos;
                            if (informarProgreso != null)
                                informarProgreso(descargado, total);
                        }

                        await destinoStream.FlushAsync(cancellationToken);
                    }

                    if (descargado != tamanoEsperado)
                        throw new InvalidOperationException(string.Format(
                            "Descarga incompleta. Se esperaban {0} bytes y se recibieron {1}.", tamanoEsperado, descargado));
                }

                return true;
            }, "descarga de " + Path.GetFileName(url), cancellationToken);
        }

        private async Task<T> EjecutarConReintentosAsync<T>(
            Func<Task<T>> operacion,
            string descripcion,
            CancellationToken cancellationToken)
        {
            Exception ultimoError = null;
            for (var intento = 1; intento <= maximoReintentos; intento++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                try
                {
                    return await operacion();
                }
                catch (Exception ex) when (EsErrorTemporal(ex, cancellationToken) && intento < maximoReintentos)
                {
                    ultimoError = ex;
                    log.Error(string.Format("Falló {0}. Reintento {1} de {2}.", descripcion, intento, maximoReintentos), ex);
                    await Task.Delay(TimeSpan.FromSeconds(Math.Min(8, intento * 2)), cancellationToken);
                }
            }

            throw ultimoError ?? new InvalidOperationException("No se pudo completar " + descripcion + ".");
        }

        private static bool EsErrorTemporal(Exception ex, CancellationToken cancellationToken)
        {
            if (cancellationToken.IsCancellationRequested)
                return false;
            return ex is HttpRequestException || ex is TaskCanceledException || ex is IOException;
        }

        public void Dispose()
        {
            cliente.Dispose();
        }
    }
}
