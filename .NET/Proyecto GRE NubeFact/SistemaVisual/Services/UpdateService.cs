using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using SistemaVisual.Models;

namespace SistemaVisual.Services
{
    public sealed class UpdateService
    {
        private readonly UpdateConfig configuracion;
        private readonly AppPaths rutas;
        private readonly LogService log;
        private readonly HashService hashService;
        private readonly BackupService backupService;
        private readonly ProcessService processService;
        private readonly StateService stateService;
        private readonly FileComparisonService comparisonService;
        private readonly PathSecurityService pathSecurity;

        public UpdateService(UpdateConfig configuracion, AppPaths rutas, LogService log)
        {
            this.configuracion = configuracion;
            this.rutas = rutas;
            this.log = log;
            hashService = new HashService();
            backupService = new BackupService();
            processService = new ProcessService();
            stateService = new StateService(log);
            comparisonService = new FileComparisonService(hashService);
            pathSecurity = new PathSecurityService(rutas.DirectorioBase, configuracion.CarpetasExcluidas);
        }

        public async Task<UpdateResult> EjecutarAsync(
            IProgress<UpdateProgress> progreso,
            Func<Task<bool>> solicitarCierreVentas,
            CancellationToken cancellationToken)
        {
            var cronometro = Stopwatch.StartNew();
            var directorioExistia = Directory.Exists(rutas.DirectorioBase);
            var modoInstalacion = !directorioExistia || File.Exists(rutas.MarcadorInstalacion);
            PrepararDirectorios(modoInstalacion);
            log.Informacion("Inicio del proceso. Modo: " + (modoInstalacion ? "instalación completa" : "actualización") + ".");

            try
            {
                RemoteFileList listado;
                using (var descargas = new DownloadService(configuracion.TimeoutSegundos, configuracion.MaximoReintentos, log))
                {
                    Informar(progreso, modoInstalacion ? "Preparando instalación..." : "Consultando actualizaciones...",
                        string.Empty, 0, true, false, modoInstalacion, 0, 0, 0, 0, 0, 0);
                    log.Informacion("Consultando el endpoint del Worker.");
                    try
                    {
                        listado = await descargas.ObtenerListadoAsync(configuracion.UrlWorker, cancellationToken);
                    }
                    catch (Exception ex) when (ex is HttpRequestException || ex is TaskCanceledException)
                    {
                        log.Error("No fue posible consultar el Worker.", ex);

                        var ejecutableLocal = pathSecurity.ObtenerRutaSegura(configuracion.EjecutablePrincipal);
                        if (!modoInstalacion && File.Exists(ejecutableLocal))
                        {
                            const string mensajeSinConexion =
                                "Sin conexión: se iniciará la versión instalada";
                            log.Informacion(
                                "No se comprobaron actualizaciones. Se iniciará Ventas.exe existente sin actualizar.");
                            Informar(progreso, mensajeSinConexion, string.Empty, 100, false, false, false,
                                0, 0, 0, 0, 0, 0);

                            return new UpdateResult
                            {
                                Actualizado = false,
                                SinCambios = true,
                                SinConexion = true,
                                ModoInstalacion = false,
                                Mensaje = mensajeSinConexion
                            };
                        }

                        throw new InvalidOperationException("No se pudo consultar el servidor de actualizaciones. Revise su conexión e inténtelo nuevamente.", ex);
                    }

                    var estado = stateService.Cargar(rutas.ArchivoEstado);
                    var plan = await CrearPlanAsync(listado, estado, modoInstalacion, progreso, cancellationToken);
                    var nuevos = plan.Count(a => a.Decision == FileDecision.Nuevo);
                    var actualizados = plan.Count(a => a.Decision == FileDecision.Actualizar);
                    var sinCambios = plan.Count(a => a.Decision == FileDecision.SinCambios);
                    var omitidos = plan.Count(a => a.Decision == FileDecision.Omitir);
                    var pendientes = plan.Where(a => a.Decision == FileDecision.Nuevo || a.Decision == FileDecision.Actualizar).ToList();

                    if (pendientes.Count > 0)
                    {
                        await AsegurarVentasCerradoAsync(solicitarCierreVentas, cancellationToken);
                        ValidarEspacioDisponible(pendientes, modoInstalacion);
                        await ProcesarPendientesAsync(descargas, pendientes, estado, modoInstalacion,
                            nuevos, actualizados, sinCambios, omitidos, listado.Archivos.Count, progreso, cancellationToken);
                    }

                    var ejecutable = pathSecurity.ObtenerRutaSegura(configuracion.EjecutablePrincipal);
                    if (!File.Exists(ejecutable))
                        throw new InvalidOperationException("La operación no puede finalizar porque Ventas.exe no está disponible.");

                    if (modoInstalacion && File.Exists(rutas.MarcadorInstalacion))
                        File.Delete(rutas.MarcadorInstalacion);

                    var sinCambiosGenerales = pendientes.Count == 0;
                    var mensaje = modoInstalacion
                        ? "Sistema instalado correctamente"
                        : sinCambiosGenerales ? "El sistema ya se encuentra actualizado" : "Sistema actualizado correctamente";

                    Informar(progreso, mensaje, string.Empty, 100, false, false, modoInstalacion,
                        listado.Archivos.Count, listado.Archivos.Count, nuevos, actualizados, sinCambios, omitidos);
                    log.Informacion(string.Format(
                        "Resultado correcto. Nuevos: {0}; actualizados: {1}; sin cambios: {2}; omitidos: {3}; duración: {4}.",
                        nuevos, actualizados, sinCambios, omitidos, cronometro.Elapsed));

                    return new UpdateResult
                    {
                        Actualizado = pendientes.Count > 0,
                        SinCambios = sinCambiosGenerales,
                        ModoInstalacion = modoInstalacion,
                        Mensaje = mensaje,
                        Nuevos = nuevos,
                        Actualizados = actualizados,
                        SinCambio = sinCambios,
                        Omitidos = omitidos
                    };
                }
            }
            catch (Exception ex)
            {
                log.Error("El proceso terminó con error después de " + cronometro.Elapsed + ".", ex);
                throw;
            }
        }

        public void IniciarSistema()
        {
            var ejecutable = pathSecurity.ObtenerRutaSegura(configuracion.EjecutablePrincipal);
            log.Informacion("Iniciando Ventas.exe.");
            processService.Iniciar(ejecutable, rutas.DirectorioBase);
        }

        public void AbrirCarpetaLogs()
        {
            processService.AbrirCarpeta(rutas.DirectorioLogs);
        }

        private void PrepararDirectorios(bool modoInstalacion)
        {
            Directory.CreateDirectory(rutas.DirectorioBase);
            Directory.CreateDirectory(rutas.DirectorioTecnico);
            Directory.CreateDirectory(rutas.DirectorioTemporal);
            Directory.CreateDirectory(rutas.DirectorioRespaldos);
            Directory.CreateDirectory(rutas.DirectorioLogs);
            if (modoInstalacion)
                Directory.CreateDirectory(Path.Combine(rutas.DirectorioBase, "Temporales"));
            LimpiarTemporalesHuerfanos();

            if (modoInstalacion && !File.Exists(rutas.MarcadorInstalacion))
                File.WriteAllText(rutas.MarcadorInstalacion, DateTimeOffset.Now.ToString("O"));
        }

        private void LimpiarTemporalesHuerfanos()
        {
            try
            {
                foreach (var archivo in Directory.GetFiles(rutas.DirectorioTemporal, "*.descarga.tmp", SearchOption.TopDirectoryOnly))
                    File.Delete(archivo);
            }
            catch (Exception ex)
            {
                log.Error("No se pudieron limpiar todos los temporales de una ejecución anterior.", ex);
            }
        }

        private async Task<List<PlannedFile>> CrearPlanAsync(
            RemoteFileList listado,
            LocalState estado,
            bool modoInstalacion,
            IProgress<UpdateProgress> progreso,
            CancellationToken cancellationToken)
        {
            ValidarListado(listado);
            var plan = new List<PlannedFile>();
            var rutasVistas = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var procesados = 0;
            var nuevos = 0;
            var actualizados = 0;
            var sinCambios = 0;
            var omitidos = 0;

            foreach (var remoto in listado.Archivos)
            {
                cancellationToken.ThrowIfCancellationRequested();
                remoto.Ruta = NormalizarRuta(remoto.Ruta);
                ValidarArchivoRemoto(remoto);
                var rutaLocal = pathSecurity.ObtenerRutaSegura(remoto.Ruta);
                if (!rutasVistas.Add(rutaLocal))
                    throw new InvalidOperationException("El servidor devolvió una ruta duplicada y la actualización fue cancelada.");

                FileDecision decision;
                if (pathSecurity.EsCarpetaExcluida(remoto.Ruta))
                {
                    decision = FileDecision.Omitir;
                    omitidos++;
                    log.Informacion(remoto.Ruta + " omitido por pertenecer a una carpeta técnica local.");
                }
                else
                {
                    LocalFileState estadoArchivo;
                    estado.Archivos.TryGetValue(remoto.Ruta, out estadoArchivo);
                    decision = await comparisonService.ClasificarAsync(remoto, rutaLocal, estadoArchivo, cancellationToken);
                    if (decision == FileDecision.Nuevo) nuevos++;
                    else if (decision == FileDecision.Actualizar) actualizados++;
                    else
                    {
                        sinCambios++;
                        log.Informacion(remoto.Ruta + " sin cambios.");
                    }
                }

                plan.Add(new PlannedFile { Archivo = remoto, RutaLocal = rutaLocal, Decision = decision });
                procesados++;
                Informar(progreso, "Comparando archivos...", remoto.Ruta, Porcentaje(procesados, listado.Archivos.Count),
                    false, false, modoInstalacion, procesados, listado.Archivos.Count, nuevos, actualizados, sinCambios, omitidos);
            }

            return plan;
        }

        private void ValidarListado(RemoteFileList listado)
        {
            if (listado == null || listado.Archivos == null || listado.Archivos.Count == 0)
                throw new InvalidOperationException("El servidor no devolvió archivos para instalar o actualizar.");
            if (!string.Equals(listado.Prefijo, configuracion.PrefijoRemoto, StringComparison.Ordinal))
                throw new InvalidOperationException("El servidor devolvió un prefijo inesperado. La operación fue cancelada por seguridad.");
        }

        private void ValidarArchivoRemoto(UpdateFile archivo)
        {
            if (archivo == null || string.IsNullOrWhiteSpace(archivo.Ruta)
                || string.IsNullOrWhiteSpace(archivo.Clave) || archivo.Tamano < 0
                || string.IsNullOrWhiteSpace(archivo.ETag))
                throw new InvalidOperationException("El servidor devolvió información incompleta para uno de los archivos.");

            var claveEsperada = configuracion.PrefijoRemoto + archivo.Ruta;
            if (!string.Equals(archivo.Clave, claveEsperada, StringComparison.Ordinal))
                throw new InvalidOperationException("El servidor devolvió una clave remota que no coincide con su ruta relativa.");

            var rutaLocal = pathSecurity.ObtenerRutaSegura(archivo.Ruta);
            string ejecutableActual;
            try { ejecutableActual = Process.GetCurrentProcess().MainModule.FileName; }
            catch { ejecutableActual = string.Empty; }
            if (!string.IsNullOrEmpty(ejecutableActual)
                && string.Equals(Path.GetFullPath(ejecutableActual), rutaLocal, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("SistemaVisual.exe no puede reemplazarse mientras está en ejecución.");
        }

        private async Task AsegurarVentasCerradoAsync(
            Func<Task<bool>> solicitarCierreVentas,
            CancellationToken cancellationToken)
        {
            if (!processService.EstaEjecutandose(configuracion.EjecutablePrincipal))
                return;

            for (var intento = 1; intento <= configuracion.MaximoReintentos; intento++)
            {
                if (solicitarCierreVentas == null || !await solicitarCierreVentas())
                    throw new InvalidOperationException("La operación fue cancelada porque Ventas.exe continúa abierto.");

                log.Informacion("Solicitando cierre normal de Ventas.exe. Intento " + intento + ".");
                if (await processService.IntentarCierreNormalAsync(
                    configuracion.EjecutablePrincipal, configuracion.EsperaCierreVentasSegundos, cancellationToken))
                    return;
            }

            throw new InvalidOperationException("No se pudo cerrar Ventas.exe de forma segura. Cierre el sistema manualmente y vuelva a intentar.");
        }

        private async Task ProcesarPendientesAsync(
            DownloadService descargas,
            IList<PlannedFile> pendientes,
            LocalState estado,
            bool modoInstalacion,
            int nuevos,
            int actualizados,
            int sinCambios,
            int omitidos,
            int total,
            IProgress<UpdateProgress> progreso,
            CancellationToken cancellationToken)
        {
            string directorioRespaldo = null;
            var completados = 0;
            var procesadosBase = sinCambios + omitidos;

            foreach (var pendiente in pendientes)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var remoto = pendiente.Archivo;
                var temporal = Path.Combine(rutas.DirectorioTemporal, Guid.NewGuid().ToString("N") + ".descarga.tmp");
                try
                {
                    var url = ConstruirUrlDescarga(remoto.Ruta);
                    log.Informacion("Descargando " + remoto.Ruta + ".");
                    await descargas.DescargarArchivoAsync(
                        url,
                        temporal,
                        remoto.Tamano,
                        (recibidos, tamano) =>
                        {
                            var fraccion = tamano > 0 ? Math.Min(1d, (double)recibidos / tamano) : 0d;
                            var porcentaje = PorcentajeConFraccion(completados, pendientes.Count, fraccion);
                            Informar(progreso, modoInstalacion ? "Instalando Sistema Visual..." : "Descargando archivos...",
                                remoto.Ruta, porcentaje, tamano <= 0, false, modoInstalacion,
                                procesadosBase + completados, total, nuevos, actualizados, sinCambios, omitidos);
                        },
                        cancellationToken);

                    var hash = await hashService.CalcularSha256Async(temporal, cancellationToken);
                    var existe = File.Exists(pendiente.RutaLocal);
                    if (!modoInstalacion && existe)
                    {
                        if (directorioRespaldo == null)
                            directorioRespaldo = backupService.CrearDirectorioRespaldo(rutas);
                        var respaldoActual = directorioRespaldo;
                        await Task.Run(
                            () => backupService.RespaldarArchivo(pendiente.RutaLocal, remoto.Ruta, respaldoActual),
                            cancellationToken);
                        log.Informacion(remoto.Ruta + " respaldado antes de reemplazarse.");
                    }

                    Directory.CreateDirectory(Path.GetDirectoryName(pendiente.RutaLocal));
                    Informar(progreso, modoInstalacion ? "Instalando Sistema Visual..." : "Aplicando actualización...",
                        remoto.Ruta, Porcentaje(completados, pendientes.Count), false, true, modoInstalacion,
                        procesadosBase + completados, total, nuevos, actualizados, sinCambios, omitidos);

                    await Task.Run(
                        () => ReemplazarSeguro(temporal, pendiente.RutaLocal, existe),
                        cancellationToken);

                    estado.Archivos[remoto.Ruta] = new LocalFileState
                    {
                        ETag = remoto.ETag,
                        Tamano = remoto.Tamano,
                        FechaModificacionRemota = remoto.FechaModificacion,
                        Sha256 = hash,
                        FechaInstalacion = DateTimeOffset.Now
                    };
                    await Task.Run(
                        () => stateService.GuardarAtomico(rutas.ArchivoEstado, estado, rutas.DirectorioTemporal),
                        cancellationToken);
                    completados++;
                    log.Informacion(remoto.Ruta + (existe ? " reemplazado correctamente." : " instalado como archivo nuevo."));
                }
                catch (Exception ex)
                {
                    log.Error("Falló el procesamiento de " + remoto.Ruta + ".", ex);
                    throw new InvalidOperationException("No se pudo procesar " + remoto.Ruta + ". Los archivos completados se conservaron; puede volver a intentar.", ex);
                }
                finally
                {
                    try { if (File.Exists(temporal)) File.Delete(temporal); }
                    catch (Exception ex) { log.Error("No se pudo eliminar un archivo temporal.", ex); }
                }
            }
        }

        private string ConstruirUrlDescarga(string rutaRelativa)
        {
            var partes = (configuracion.PrefijoRemoto + rutaRelativa).Split('/');
            var codificada = string.Join("/", partes.Select(Uri.EscapeDataString));
            return configuracion.DominioDescargas.TrimEnd('/') + "/" + codificada;
        }

        private static void ReemplazarSeguro(string temporal, string destino, bool existe)
        {
            if (existe)
                File.Replace(temporal, destino, null, true);
            else
                File.Move(temporal, destino);
        }

        private void ValidarEspacioDisponible(IEnumerable<PlannedFile> pendientes, bool modoInstalacion)
        {
            var requerido = pendientes.Sum(a => Math.Max(0, a.Archivo.Tamano));
            if (!modoInstalacion)
                requerido *= 2;
            var unidad = new DriveInfo(Path.GetPathRoot(rutas.DirectorioBase));
            if (unidad.AvailableFreeSpace < requerido)
                throw new IOException("No hay espacio suficiente para completar la operación.");
        }

        private static string NormalizarRuta(string ruta)
        {
            return (ruta ?? string.Empty).Replace('\\', '/');
        }

        private static int Porcentaje(int procesados, int total)
        {
            return total <= 0 ? 0 : Math.Max(0, Math.Min(100, (int)Math.Round((double)procesados / total * 100)));
        }

        private static int PorcentajeConFraccion(int completados, int total, double fraccion)
        {
            return total <= 0 ? 0 : Math.Max(0, Math.Min(99, (int)Math.Round((completados + fraccion) / total * 100)));
        }

        private static void Informar(
            IProgress<UpdateProgress> progreso,
            string estado,
            string archivo,
            int porcentaje,
            bool indeterminado,
            bool instalando,
            bool modoInstalacion,
            int procesados,
            int total,
            int nuevos,
            int actualizados,
            int sinCambios,
            int omitidos)
        {
            if (progreso == null) return;
            progreso.Report(new UpdateProgress
            {
                Estado = estado,
                Archivo = archivo,
                Porcentaje = Math.Max(0, Math.Min(100, porcentaje)),
                Indeterminado = indeterminado,
                Instalando = instalando,
                ModoInstalacion = modoInstalacion,
                Procesados = procesados,
                Total = total,
                Nuevos = nuevos,
                Actualizados = actualizados,
                SinCambios = sinCambios,
                Omitidos = omitidos
            });
        }
    }
}
