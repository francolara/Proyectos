using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace SistemaVisual.Models
{
    public sealed class LocalState
    {
        [JsonProperty("archivos")]
        public Dictionary<string, LocalFileState> Archivos { get; set; }
            = new Dictionary<string, LocalFileState>(StringComparer.OrdinalIgnoreCase);
    }

    public sealed class LocalFileState
    {
        [JsonProperty("etag")]
        public string ETag { get; set; }

        [JsonProperty("tamano")]
        public long Tamano { get; set; }

        [JsonProperty("fechaModificacionRemota")]
        public DateTimeOffset? FechaModificacionRemota { get; set; }

        [JsonProperty("sha256")]
        public string Sha256 { get; set; }

        [JsonProperty("fechaInstalacion")]
        public DateTimeOffset FechaInstalacion { get; set; }
    }

    public enum FileDecision
    {
        Nuevo,
        Actualizar,
        SinCambios,
        Omitir
    }

    public sealed class PlannedFile
    {
        public UpdateFile Archivo { get; set; }
        public string RutaLocal { get; set; }
        public FileDecision Decision { get; set; }
    }
}
