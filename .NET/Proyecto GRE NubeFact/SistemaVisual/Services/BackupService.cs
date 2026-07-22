using System;
using System.IO;

namespace SistemaVisual.Services
{
    public sealed class BackupService
    {
        public string CrearDirectorioRespaldo(AppPaths rutas)
        {
            Directory.CreateDirectory(rutas.DirectorioRespaldos);
            var nombreBase = DateTime.Now.ToString("yyyy-MM-dd_HHmmss");
            var candidato = Path.Combine(rutas.DirectorioRespaldos, nombreBase);
            var consecutivo = 1;
            while (Directory.Exists(candidato))
            {
                candidato = Path.Combine(rutas.DirectorioRespaldos, nombreBase + "_" + consecutivo);
                consecutivo++;
            }

            Directory.CreateDirectory(candidato);
            return candidato;
        }

        public void RespaldarArchivo(string origen, string rutaRelativa, string directorioRespaldo)
        {
            var destino = Path.Combine(directorioRespaldo, rutaRelativa);
            Directory.CreateDirectory(Path.GetDirectoryName(destino));
            File.Copy(origen, destino, false);
        }
    }
}
