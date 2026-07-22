using System.Collections.Generic;
using Newtonsoft.Json;

namespace SistemaVisual.Models
{
    public sealed class UpdateConfig
    {
        [JsonProperty("UrlWorker")]
        public string UrlWorker { get; set; }

        [JsonProperty("DominioDescargas")]
        public string DominioDescargas { get; set; }

        [JsonProperty("NombreBucket")]
        public string NombreBucket { get; set; }

        [JsonProperty("PrefijoRemoto")]
        public string PrefijoRemoto { get; set; }

        [JsonProperty("CarpetaLocal")]
        public string CarpetaLocal { get; set; }

        [JsonProperty("EjecutablePrincipal")]
        public string EjecutablePrincipal { get; set; }

        [JsonProperty("CarpetaTemporal")]
        public string CarpetaTemporal { get; set; }

        [JsonProperty("CarpetaRespaldo")]
        public string CarpetaRespaldo { get; set; }

        [JsonProperty("CarpetasExcluidas")]
        public List<string> CarpetasExcluidas { get; set; }

        [JsonProperty("TimeoutSegundos")]
        public int TimeoutSegundos { get; set; }

        [JsonProperty("MaximoReintentos")]
        public int MaximoReintentos { get; set; }

        [JsonProperty("EsperaCierreVentasSegundos")]
        public int EsperaCierreVentasSegundos { get; set; }

        public static UpdateConfig CrearPredeterminada()
        {
            return new UpdateConfig
            {
                UrlWorker = "https://sistema-visual-api.larasoft-dev.workers.dev/api/archivos",
                DominioDescargas = "https://actualizaciones.fralsetech.com",
                NombreBucket = "sistemavisual-actualizaciones",
                PrefijoRemoto = "PROVEEPERU/Sistema Visual/",
                CarpetaLocal = @"C:\Sistema Visual",
                EjecutablePrincipal = "Ventas.exe",
                CarpetaTemporal = @".actualizador\temporal",
                CarpetaRespaldo = "Respaldo",
                CarpetasExcluidas = new List<string> { "Respaldo", ".actualizador" },
                TimeoutSegundos = 60,
                MaximoReintentos = 3,
                EsperaCierreVentasSegundos = 15
            };
        }
    }
}
