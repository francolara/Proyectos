using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace SistemaVisual.Services
{
    public sealed class ProcessService
    {
        public bool EstaEjecutandose(string ejecutable)
        {
            var procesos = Process.GetProcessesByName(Path.GetFileNameWithoutExtension(ejecutable));
            try { return procesos.Length > 0; }
            finally { foreach (var proceso in procesos) proceso.Dispose(); }
        }

        public async Task<bool> IntentarCierreNormalAsync(
            string ejecutable,
            int esperaSegundos,
            CancellationToken cancellationToken)
        {
            var procesos = Process.GetProcessesByName(Path.GetFileNameWithoutExtension(ejecutable));
            try
            {
                foreach (var proceso in procesos)
                {
                    try { proceso.CloseMainWindow(); } catch { }
                }

                var limite = DateTime.UtcNow.AddSeconds(esperaSegundos);
                while (DateTime.UtcNow < limite)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (procesos.All(p => ProcesoTerminado(p)))
                        return true;
                    await Task.Delay(500, cancellationToken);
                }

                return procesos.All(p => ProcesoTerminado(p));
            }
            finally
            {
                foreach (var proceso in procesos) proceso.Dispose();
            }
        }

        private static bool ProcesoTerminado(Process proceso)
        {
            try { return proceso.HasExited; }
            catch { return true; }
        }

        public void Iniciar(string rutaEjecutable, string directorioTrabajo)
        {
            if (!File.Exists(rutaEjecutable))
                throw new FileNotFoundException("No existe Ventas.exe después del proceso.", rutaEjecutable);

            var proceso = Process.Start(new ProcessStartInfo
            {
                FileName = rutaEjecutable,
                WorkingDirectory = directorioTrabajo,
                UseShellExecute = true
            });
            if (proceso == null)
                throw new InvalidOperationException("Windows no pudo iniciar Ventas.exe.");
            proceso.Dispose();
        }

        public void AbrirCarpeta(string ruta)
        {
            Directory.CreateDirectory(ruta);
            Process.Start(new ProcessStartInfo { FileName = ruta, UseShellExecute = true });
        }
    }
}
