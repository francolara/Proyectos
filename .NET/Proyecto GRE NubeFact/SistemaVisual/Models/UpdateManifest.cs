using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace SistemaVisual.Models
{
    public sealed class RemoteFileList
    {
        [JsonProperty("prefijo")]
        public string Prefijo { get; set; }

        [JsonProperty("fechaConsulta")]
        public DateTimeOffset FechaConsulta { get; set; }

        [JsonProperty("archivos")]
        public List<UpdateFile> Archivos { get; set; }
    }
}
