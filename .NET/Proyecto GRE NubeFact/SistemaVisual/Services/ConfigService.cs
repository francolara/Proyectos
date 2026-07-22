using System;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using SistemaVisual.Models;

namespace SistemaVisual.Services
{
    public sealed class ConfigService
    {
        public string RutaConfiguracion
        {
            get { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "actualizador.config.json"); }
        }

        public UpdateConfig CargarOCrear(out bool archivoCreado)
        {
            archivoCreado = !File.Exists(RutaConfiguracion);

            if (archivoCreado)
            {
                var configuracionPredeterminada = UpdateConfig.CrearPredeterminada();
                File.WriteAllText(
                    RutaConfiguracion,
                    JsonConvert.SerializeObject(configuracionPredeterminada, Formatting.Indented),
                    new System.Text.UTF8Encoding(false));
                return configuracionPredeterminada;
            }

            UpdateConfig configuracion;
            try
            {
                configuracion = JsonConvert.DeserializeObject<UpdateConfig>(File.ReadAllText(RutaConfiguracion));
            }
            catch (JsonException ex)
            {
                throw new InvalidOperationException("El archivo actualizador.config.json contiene JSON inválido.", ex);
            }

            Validar(configuracion);
            return configuracion;
        }

        private static void Validar(UpdateConfig configuracion)
        {
            if (configuracion == null)
                throw new InvalidOperationException("No se pudo leer la configuración del actualizador.");

            ValidarUrl(configuracion.UrlWorker, "UrlWorker");
            ValidarUrl(configuracion.DominioDescargas, "DominioDescargas");

            if (configuracion.UrlWorker.IndexOf("CONFIGURE-SU-WORKER", StringComparison.OrdinalIgnoreCase) >= 0)
                throw new InvalidOperationException("Configure UrlWorker en actualizador.config.json antes de ejecutar el actualizador.");

            if (string.IsNullOrWhiteSpace(configuracion.NombreBucket)
                || string.IsNullOrWhiteSpace(configuracion.PrefijoRemoto)
                || !configuracion.PrefijoRemoto.EndsWith("/", StringComparison.Ordinal))
                throw new InvalidOperationException("NombreBucket y PrefijoRemoto deben estar configurados; el prefijo debe terminar en '/'.");

            if (string.IsNullOrWhiteSpace(configuracion.CarpetaLocal) || !Path.IsPathRooted(configuracion.CarpetaLocal))
                throw new InvalidOperationException("CarpetaLocal debe ser una ruta absoluta.");

            if (string.IsNullOrWhiteSpace(configuracion.EjecutablePrincipal)
                || Path.IsPathRooted(configuracion.EjecutablePrincipal)
                || configuracion.EjecutablePrincipal.IndexOf("..", StringComparison.Ordinal) >= 0)
                throw new InvalidOperationException("EjecutablePrincipal debe ser una ruta relativa segura.");

            ValidarRutaRelativa(configuracion.CarpetaTemporal, "CarpetaTemporal");
            ValidarRutaRelativa(configuracion.CarpetaRespaldo, "CarpetaRespaldo");

            if (configuracion.CarpetasExcluidas == null
                || !configuracion.CarpetasExcluidas.Exists(c => string.Equals(c, "Respaldo", StringComparison.OrdinalIgnoreCase))
                || !configuracion.CarpetasExcluidas.Exists(c => string.Equals(c, ".actualizador", StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("CarpetasExcluidas debe incluir Respaldo y .actualizador.");

            if (configuracion.TimeoutSegundos < 5 || configuracion.TimeoutSegundos > 600)
                throw new InvalidOperationException("TimeoutSegundos debe estar entre 5 y 600.");
            if (configuracion.MaximoReintentos < 1 || configuracion.MaximoReintentos > 10)
                throw new InvalidOperationException("MaximoReintentos debe estar entre 1 y 10.");
            if (configuracion.EsperaCierreVentasSegundos < 1 || configuracion.EsperaCierreVentasSegundos > 120)
                throw new InvalidOperationException("EsperaCierreVentasSegundos debe estar entre 1 y 120.");
        }

        private static void ValidarUrl(string valor, string nombre)
        {
            Uri uri;
            if (string.IsNullOrWhiteSpace(valor) || !Uri.TryCreate(valor, UriKind.Absolute, out uri)
                || (uri.Scheme != Uri.UriSchemeHttps && !(uri.Scheme == Uri.UriSchemeHttp && uri.IsLoopback))
                || !string.IsNullOrEmpty(uri.UserInfo))
                throw new InvalidOperationException(nombre + " debe ser una URL HTTPS válida y no debe contener credenciales.");
        }

        private static void ValidarRutaRelativa(string valor, string nombre)
        {
            if (string.IsNullOrWhiteSpace(valor) || Path.IsPathRooted(valor)
                || valor.Split(new[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries)
                    .Any(segmento => segmento == ".."))
                throw new InvalidOperationException(nombre + " debe ser una ruta relativa segura.");
        }
    }
}
