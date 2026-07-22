using Newtonsoft.Json;

namespace SistemaVisual.Models
{
    public sealed class UpdateFile
    {
        [JsonProperty("ruta")]
        public string Ruta { get; set; }

        [JsonProperty("clave")]
        public string Clave { get; set; }

        [JsonProperty("tamano")]
        public long Tamano { get; set; }

        [JsonProperty("etag")]
        public string ETag { get; set; }

        [JsonProperty("fechaModificacion")]
        public System.DateTimeOffset? FechaModificacion { get; set; }

        [JsonProperty("url")]
        public string Url { get; set; }
    }
}
