using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Newtonsoft.Json;
using SistemaVisual.Models;

namespace SistemaVisual.Services
{
    public sealed class StateService
    {
        private readonly LogService log;

        public StateService(LogService log)
        {
            this.log = log;
        }

        public LocalState Cargar(string archivoEstado)
        {
            if (!File.Exists(archivoEstado))
                return new LocalState();

            try
            {
                var estado = JsonConvert.DeserializeObject<LocalState>(File.ReadAllText(archivoEstado));
                if (estado == null || estado.Archivos == null)
                    throw new JsonException("El estado no contiene la colección de archivos.");

                estado.Archivos = new Dictionary<string, LocalFileState>(estado.Archivos, StringComparer.OrdinalIgnoreCase);
                return estado;
            }
            catch (Exception ex)
            {
                var copia = archivoEstado + ".danado." + DateTime.Now.ToString("yyyyMMddHHmmss") + ".json";
                try { File.Copy(archivoEstado, copia, false); } catch { }
                log.Error("estado.json está dañado. Se conservará una copia y se resincronizarán los archivos.", ex);
                return new LocalState();
            }
        }

        public void GuardarAtomico(string archivoEstado, LocalState estado, string directorioTemporal)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(archivoEstado));
            Directory.CreateDirectory(directorioTemporal);
            var temporal = Path.Combine(directorioTemporal, "estado-" + Guid.NewGuid().ToString("N") + ".tmp");
            var respaldoEstado = archivoEstado + ".anterior";
            try
            {
                var json = JsonConvert.SerializeObject(estado, Formatting.Indented);
                using (var stream = new FileStream(temporal, FileMode.CreateNew, FileAccess.Write, FileShare.None, 4096, FileOptions.WriteThrough))
                using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
                {
                    writer.Write(json);
                    writer.Flush();
                    stream.Flush(true);
                }

                if (File.Exists(archivoEstado))
                {
                    File.Replace(temporal, archivoEstado, respaldoEstado, true);
                    if (File.Exists(respaldoEstado))
                        File.Delete(respaldoEstado);
                }
                else
                {
                    File.Move(temporal, archivoEstado);
                }
            }
            finally
            {
                if (File.Exists(temporal))
                    File.Delete(temporal);
            }
        }
    }
}
