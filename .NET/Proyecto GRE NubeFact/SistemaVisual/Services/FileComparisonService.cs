using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using SistemaVisual.Models;

namespace SistemaVisual.Services
{
    public sealed class FileComparisonService
    {
        private readonly HashService hashService;

        public FileComparisonService(HashService hashService)
        {
            this.hashService = hashService;
        }

        public async Task<FileDecision> ClasificarAsync(
            UpdateFile remoto,
            string rutaLocal,
            LocalFileState estado,
            CancellationToken cancellationToken)
        {
            if (!File.Exists(rutaLocal))
                return FileDecision.Nuevo;
            if (estado == null)
                return FileDecision.Actualizar;
            if (!string.Equals(remoto.ETag, estado.ETag, StringComparison.Ordinal)
                || remoto.Tamano != estado.Tamano)
                return FileDecision.Actualizar;

            var hashActual = await hashService.CalcularSha256Async(rutaLocal, cancellationToken);
            return string.Equals(hashActual, estado.Sha256, StringComparison.OrdinalIgnoreCase)
                ? FileDecision.SinCambios
                : FileDecision.Actualizar;
        }
    }
}
