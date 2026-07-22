using System;
using System.IO;
using System.Text;

namespace SistemaVisual.Services
{
    public sealed class LogService
    {
        private readonly string archivoLog;
        private readonly object bloqueo = new object();

        public LogService(string archivoLog)
        {
            this.archivoLog = archivoLog;
        }

        public void Informacion(string mensaje)
        {
            Escribir("INFO", mensaje, null);
        }

        public void Error(string mensaje, Exception excepcion)
        {
            Escribir("ERROR", mensaje, excepcion);
        }

        private void Escribir(string nivel, string mensaje, Exception excepcion)
        {
            try
            {
                lock (bloqueo)
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(archivoLog));
                    var detalle = excepcion == null ? string.Empty : Environment.NewLine + excepcion;
                    File.AppendAllText(
                        archivoLog,
                        string.Format("{0:yyyy-MM-dd HH:mm:ss} [{1}] {2}{3}{4}", DateTime.Now, nivel, mensaje, detalle, Environment.NewLine),
                        new UTF8Encoding(false));
                }
            }
            catch
            {
                // El log nunca debe ocultar el error original del actualizador.
            }
        }
    }
}
