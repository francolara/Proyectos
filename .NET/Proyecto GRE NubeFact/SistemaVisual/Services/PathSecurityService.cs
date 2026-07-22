using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SistemaVisual.Services
{
    public sealed class PathSecurityService
    {
        private readonly string directorioBase;
        private readonly string prefijoBase;
        private readonly HashSet<string> directoriosExcluidos;

        public PathSecurityService(string directorioBase, IEnumerable<string> directoriosExcluidos)
        {
            this.directorioBase = Path.GetFullPath(directorioBase).TrimEnd(Path.DirectorySeparatorChar);
            prefijoBase = this.directorioBase + Path.DirectorySeparatorChar;
            this.directoriosExcluidos = new HashSet<string>(directoriosExcluidos ?? Enumerable.Empty<string>(), StringComparer.OrdinalIgnoreCase);
        }

        public bool EsCarpetaExcluida(string rutaRelativa)
        {
            var segmentos = Separar(rutaRelativa);
            return segmentos.Length > 0 && directoriosExcluidos.Contains(segmentos[0]);
        }

        public string ObtenerRutaSegura(string rutaRelativa)
        {
            if (string.IsNullOrWhiteSpace(rutaRelativa))
                throw new InvalidOperationException("El Worker devolvió una ruta vacía.");

            var normalizada = rutaRelativa.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);
            if (Path.IsPathRooted(normalizada)
                || normalizada.StartsWith("\\", StringComparison.Ordinal)
                || normalizada.StartsWith("/", StringComparison.Ordinal)
                || normalizada.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
                throw new InvalidOperationException("No se permiten rutas absolutas: " + rutaRelativa);

            var segmentos = Separar(normalizada);
            if (segmentos.Length == 0
                || segmentos.Any(s => s == "." || s == ".."
                    || s.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0
                    || s.EndsWith(".", StringComparison.Ordinal)
                    || s.EndsWith(" ", StringComparison.Ordinal)))
                throw new InvalidOperationException("La ruta contiene segmentos no permitidos: " + rutaRelativa);

            var rutaCompleta = Path.GetFullPath(Path.Combine(directorioBase, normalizada));
            if (!rutaCompleta.StartsWith(prefijoBase, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("La ruta intenta escribir fuera de C:\\Sistema Visual: " + rutaRelativa);

            var directorioActual = directorioBase;
            for (var indice = 0; indice < segmentos.Length - 1; indice++)
            {
                directorioActual = Path.Combine(directorioActual, segmentos[indice]);
                if (Directory.Exists(directorioActual)
                    && (File.GetAttributes(directorioActual) & FileAttributes.ReparsePoint) != 0)
                    throw new InvalidOperationException("La ruta atraviesa un enlace o punto de unión no permitido: " + rutaRelativa);
            }

            return rutaCompleta;
        }

        private static string[] Separar(string ruta)
        {
            return ruta.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar)
                .Split(new[] { Path.DirectorySeparatorChar }, StringSplitOptions.RemoveEmptyEntries);
        }
    }
}
