using System.IO;
using SistemaVisual.Models;

namespace SistemaVisual.Services
{
    public sealed class AppPaths
    {
        public AppPaths(UpdateConfig configuracion)
        {
            DirectorioBase = Path.GetFullPath(configuracion.CarpetaLocal).TrimEnd(Path.DirectorySeparatorChar);
            DirectorioRespaldos = Path.GetFullPath(Path.Combine(DirectorioBase, configuracion.CarpetaRespaldo));
            DirectorioTecnico = Path.Combine(DirectorioBase, ".actualizador");
            DirectorioTemporal = Path.GetFullPath(Path.Combine(DirectorioBase, configuracion.CarpetaTemporal));
        }

        public string DirectorioBase { get; private set; }
        public string DirectorioTecnico { get; private set; }
        public string DirectorioTemporal { get; private set; }
        public string DirectorioRespaldos { get; private set; }
        public string ArchivoEstado { get { return Path.Combine(DirectorioTecnico, "estado.json"); } }
        public string MarcadorInstalacion { get { return Path.Combine(DirectorioTecnico, "instalacion.pendiente"); } }
        public string DirectorioLogs { get { return Path.Combine(DirectorioRespaldos, "Logs"); } }
        public string ArchivoLog { get { return Path.Combine(DirectorioLogs, "actualizador.log"); } }
    }
}
