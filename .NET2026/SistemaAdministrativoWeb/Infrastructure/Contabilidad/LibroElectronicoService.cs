using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

// Firma: FRANCO LARA - 04/08/2026 | Genera libros principales y planes complementarios con el indicador de contenido correspondiente, incluso en periodos sin movimientos.
public sealed class LibroElectronicoService(
    ILibroDiario51Service libroDiario51Service,
    ILibroDiario52Service libroDiario52Service,
    ILibroMayor61Service libroMayor61Service,
    ILibroElectronicoRepository libroElectronicoRepository,
    IAsientoRepository asientoRepository,
    IPlanCuentaRepository planCuentaRepository,
    IMonedaRepository monedaRepository,
    ITipoComprobanteRepository tipoComprobanteRepository,
    IPersonaRepository personaRepository,
    IPleValidationService pleValidationService,
    IPleFileNameService pleFileNameService,
    IPleTxtGenerator pleTxtGenerator,
    IPleDownloadStore pleDownloadStore) : ILibroElectronicoService
{
    public Task<PleConsultaResultadoDto> ConsultarAsync(LibroElectronicoConsultaRequest request, string empresa, string ruc, int paginaPreview, int tamanoPaginaPreview, int paginaHistorial, int tamanoPaginaHistorial, CancellationToken cancellationToken = default)
    {
        return ConstruirResultadoAsync(request, empresa, ruc, paginaPreview, tamanoPaginaPreview, paginaHistorial, tamanoPaginaHistorial, ejecutarValidacion: false, cancellationToken);
    }

    public Task<PleConsultaResultadoDto> ValidarAsync(LibroElectronicoConsultaRequest request, string empresa, string ruc, int paginaPreview, int tamanoPaginaPreview, int paginaHistorial, int tamanoPaginaHistorial, CancellationToken cancellationToken = default)
    {
        return ConstruirResultadoAsync(request, empresa, ruc, paginaPreview, tamanoPaginaPreview, paginaHistorial, tamanoPaginaHistorial, ejecutarValidacion: true, cancellationToken);
    }

    public async Task<PleGenerationResultDto> GenerarAsync(LibroElectronicoConsultaRequest request, string empresa, string ruc, string usuarioGeneracion, int paginaPreview, int tamanoPaginaPreview, int paginaHistorial, int tamanoPaginaHistorial, CancellationToken cancellationToken = default)
    {
        var consulta = await ConstruirResultadoAsync(request, empresa, ruc, paginaPreview, tamanoPaginaPreview, paginaHistorial, tamanoPaginaHistorial, ejecutarValidacion: true, cancellationToken);
        if (consulta.Validacion.TieneErroresCriticos)
        {
            return new PleGenerationResultDto
            {
                Generado = false,
                Mensaje = $"Se encontraron {consulta.Validacion.CantidadErrores} errores que impiden generar el archivo.",
                NombreArchivo = consulta.Resumen.NombreArchivo,
                Consulta = consulta
            };
        }

        if (consulta.GeneracionBloqueada)
        {
            return new PleGenerationResultDto
            {
                Generado = false,
                Mensaje = consulta.MensajeBloqueoGeneracion,
                NombreArchivo = consulta.Resumen.NombreArchivo,
                Consulta = consulta
            };
        }

        byte[] contenidoPrincipal;
        switch (request.LibroElectronico)
        {
            case PleLibroElectronicoCatalogo.LibroDiario52:
                contenidoPrincipal = await pleTxtGenerator.GenerarLibroDiario52Async(
                    await libroDiario52Service.ListarAsync(request, cancellationToken), cancellationToken);
                break;
            case PleLibroElectronicoCatalogo.LibroMayor61:
                contenidoPrincipal = await pleTxtGenerator.GenerarLibroMayor61Async(
                    await libroMayor61Service.ListarAsync(request, cancellationToken), cancellationToken);
                break;
            default:
                contenidoPrincipal = await pleTxtGenerator.GenerarLibroDiario51Async(
                    await libroDiario51Service.ListarAsync(request, cancellationToken), cancellationToken);
                break;
        }

        var nombreDescarga = consulta.Resumen.NombreArchivo;
        string? nombreComplementario = null;
        byte[]? contenidoComplementario = null;
        var observacionPlan = string.Empty;
        var formatoPlan = PleLibroElectronicoCatalogo.ObtenerPlanComplementario(request.LibroElectronico);
        var huellaPlan = string.Empty;
        var snapshotPlan = string.Empty;
        var cantidadPlan = 0;

        if (formatoPlan is not null)
        {
            var cuentas = (await planCuentaRepository.ListarPorEmpresaAsync(request.IdEmpresa, false, cancellationToken))
                .Where(EsCuentaPlanExportable)
                .ToList();
            huellaPlan = CalcularHuellaPlan(cuentas);
            var preparacionPlan = PrepararPlanContable(cuentas, consulta.Presentacion.SnapshotUltimaPresentacion, request.Anio, request.Mes);
            contenidoComplementario = await pleTxtGenerator.GenerarPlanContableAsync(preparacionPlan.Exportacion, cancellationToken);
            cantidadPlan = preparacionPlan.Exportacion.Count;
            snapshotPlan = JsonSerializer.Serialize(preparacionPlan.Snapshot);
            nombreComplementario = pleFileNameService.ConstruirNombreArchivo(ruc, request.Anio, request.Mes, formatoPlan, "PEN", tieneContenido: cantidadPlan > 0);
            observacionPlan = preparacionPlan.EsCompleto
                ? $" Incluye plan contable {formatoPlan} completo con {cantidadPlan} cuentas por ser la primera presentacion del ejercicio."
                : $" Incluye plan contable {formatoPlan} incremental con {cantidadPlan} cuentas nuevas o modificadas.";
        }

        await libroElectronicoRepository.RegistrarHistorialAsync(new PleHistorialRegistroRequest
        {
            IdEmpresa = request.IdEmpresa,
            Periodo = PlePeriodoHelper.FormarPeriodoContable(request.Anio, request.Mes),
            CodigoLibro = request.LibroElectronico,
            CodigoFormato = PleLibroElectronicoCatalogo.ObtenerCodigoSunat(request.LibroElectronico),
            NombreArchivo = nombreDescarga,
            CantidadRegistros = consulta.Resumen.CantidadMovimientos,
            TotalDebe = consulta.Resumen.TotalDebe,
            TotalHaber = consulta.Resumen.TotalHaber,
            Estado = "GENERADO",
            Observaciones = $"{(consulta.Validacion.CantidadAdvertencias > 0 ? $"{consulta.Validacion.CantidadAdvertencias} advertencias." : string.Empty)}{observacionPlan}".Trim(),
            UsuarioGeneracion = usuarioGeneracion,
            CodigoFormatoComplementario = formatoPlan ?? string.Empty,
            NombreArchivoComplementario = nombreComplementario ?? string.Empty,
            CantidadRegistrosComplementario = cantidadPlan,
            HuellaPlanContable = huellaPlan,
            PlanContableSnapshot = snapshotPlan
        }, cancellationToken);

        var token = pleDownloadStore.Guardar(nombreDescarga, contenidoPrincipal);
        var tokenComplementario = nombreComplementario is not null && contenidoComplementario is not null
            ? pleDownloadStore.Guardar(nombreComplementario, contenidoComplementario)
            : string.Empty;

        var consultaActualizada = await ConstruirResultadoAsync(request, empresa, ruc, paginaPreview, tamanoPaginaPreview, paginaHistorial, tamanoPaginaHistorial, ejecutarValidacion: true, cancellationToken);

        return new PleGenerationResultDto
        {
            Generado = true,
            Mensaje = formatoPlan is null
                ? "El archivo TXT fue generado correctamente."
                : $"Los archivos TXT {request.LibroElectronico} y {formatoPlan} fueron generados correctamente. Descargue ambos y selecciónelos simultáneamente en el PLE.",
            TokenDescarga = token,
            TokenDescargaComplementario = tokenComplementario,
            NombreArchivo = nombreDescarga,
            NombreArchivoComplementario = nombreComplementario ?? string.Empty,
            Consulta = consultaActualizada
        };
    }

    public PleDownloadPayload? ObtenerDescarga(string token, bool remover = false)
    {
        var payload = pleDownloadStore.Obtener(token);
        if (payload is not null && remover)
        {
            pleDownloadStore.Remover(token);
        }

        return payload;
    }

    public Task ActualizarPresentacionAsync(PlePresentacionUpdateRequest request, CancellationToken cancellationToken = default)
    {
        return libroElectronicoRepository.ActualizarPresentacionAsync(request, cancellationToken);
    }

    private async Task<PleConsultaResultadoDto> ConstruirResultadoAsync(LibroElectronicoConsultaRequest request, string empresa, string ruc, int paginaPreview, int tamanoPaginaPreview, int paginaHistorial, int tamanoPaginaHistorial, bool ejecutarValidacion, CancellationToken cancellationToken)
    {
        var libro = PleLibroElectronicoCatalogo.Normalizar(request.LibroElectronico);
        var periodoContable = PlePeriodoHelper.FormarPeriodoContable(request.Anio, request.Mes);

        IReadOnlyCollection<LibroDiario51Dto> items51 = [];
        IReadOnlyCollection<LibroDiario52Dto> items52 = [];
        IReadOnlyCollection<LibroMayor61Dto> items61 = [];

        switch (libro)
        {
            case PleLibroElectronicoCatalogo.LibroDiario52:
                items52 = await libroDiario52Service.ListarAsync(request, cancellationToken);
                break;
            case PleLibroElectronicoCatalogo.LibroMayor61:
                items61 = await libroMayor61Service.ListarAsync(request, cancellationToken);
                break;
            default:
                items51 = await libroDiario51Service.ListarAsync(request, cancellationToken);
                break;
        }

        var totalMovimientos = items51.Count + items52.Count + items61.Count;
        var nombreArchivo = pleFileNameService.ConstruirNombreArchivo(
            ruc,
            request.Anio,
            request.Mes,
            libro,
            "PEN",
            tieneContenido: totalMovimientos > 0);

        var historial = await libroElectronicoRepository.ListarHistorialAsync(request.IdEmpresa, request.Anio, request.Mes, libro, paginaHistorial, tamanoPaginaHistorial, cancellationToken);
        var presentacion = await libroElectronicoRepository.ObtenerContextoPresentacionAsync(request.IdEmpresa, request.Anio, request.Mes, libro, cancellationToken);
        var (generacionBloqueada, mensajeBloqueoGeneracion) = await ObtenerBloqueoGeneracionAsync(request, presentacion, cancellationToken);
        var asientos = await asientoRepository.ListarPorEmpresaAsync(request.IdEmpresa, periodoContable, false, cancellationToken);
        var cuentas = await planCuentaRepository.ListarPorEmpresaAsync(request.IdEmpresa, false, cancellationToken);
        var monedas = await monedaRepository.ListarActivasAsync(cancellationToken);
        var tiposDocumento = await personaRepository.ListarTiposDocumentoAsync(cancellationToken);
        var tiposComprobante = await tipoComprobanteRepository.ListarActivosAsync(false, false, cancellationToken);

        var validacion = ejecutarValidacion
            ? await pleValidationService.ValidarAsync(request, empresa, ruc, asientos, cuentas, monedas, tiposDocumento, tiposComprobante, items51, items52, items61, cancellationToken)
            : new PleValidationResultDto();

        var totalAsientos = libro switch
        {
            PleLibroElectronicoCatalogo.LibroDiario52 => items52.Select(x => x.Cuo).Distinct(StringComparer.OrdinalIgnoreCase).Count(),
            PleLibroElectronicoCatalogo.LibroMayor61 => items61.Select(x => x.Cuo).Distinct(StringComparer.OrdinalIgnoreCase).Count(),
            _ => items51.Select(x => x.Cuo).Distinct(StringComparer.OrdinalIgnoreCase).Count()
        };

        var totalDebe = libro switch
        {
            PleLibroElectronicoCatalogo.LibroDiario52 => items52.Sum(x => x.Debe),
            PleLibroElectronicoCatalogo.LibroMayor61 => items61.Sum(x => x.Debe),
            _ => items51.Sum(x => x.Debe)
        };

        var totalHaber = libro switch
        {
            PleLibroElectronicoCatalogo.LibroDiario52 => items52.Sum(x => x.Haber),
            PleLibroElectronicoCatalogo.LibroMayor61 => items61.Sum(x => x.Haber),
            _ => items51.Sum(x => x.Haber)
        };

        return new PleConsultaResultadoDto
        {
            LibroElectronico = libro,
            Resumen = new PleResumenDto
            {
                Empresa = empresa,
                Ruc = ruc,
                Libro = PleLibroElectronicoCatalogo.ObtenerNombre(libro),
                Formato = PleLibroElectronicoCatalogo.ObtenerCodigoSunat(libro),
                Periodo = PlePeriodoHelper.FormarPeriodo(request.Anio, request.Mes),
                CantidadAsientos = totalAsientos,
                CantidadMovimientos = totalMovimientos,
                TotalDebe = totalDebe,
                TotalHaber = totalHaber,
                NombreArchivo = nombreArchivo
            },
            Validacion = validacion,
            LibroDiario51Items = Paginado(items51, paginaPreview, tamanoPaginaPreview),
            LibroDiario52Items = Paginado(items52, paginaPreview, tamanoPaginaPreview),
            LibroMayor61Items = Paginado(items61, paginaPreview, tamanoPaginaPreview),
            TotalRegistrosPreview = totalMovimientos,
            PaginaPreview = paginaPreview,
            TamanoPaginaPreview = tamanoPaginaPreview,
            Historial = historial,
            Presentacion = presentacion,
            GeneracionBloqueada = generacionBloqueada,
            MensajeBloqueoGeneracion = mensajeBloqueoGeneracion
        };
    }

    private static IReadOnlyCollection<T> Paginado<T>(IReadOnlyCollection<T> items, int pagina, int tamano)
    {
        if (items.Count == 0)
        {
            return [];
        }

        var paginaTrabajo = Math.Max(1, pagina);
        var tamanoTrabajo = Math.Max(1, tamano);
        return items.Skip((paginaTrabajo - 1) * tamanoTrabajo).Take(tamanoTrabajo).ToList();
    }

    private static string CalcularHuellaPlan(IReadOnlyCollection<PlanCuentaDto> cuentas)
    {
        var contenido = string.Join('\n', cuentas
            .Where(item => item.Estado)
            .OrderBy(item => item.CodigoCuenta, StringComparer.OrdinalIgnoreCase)
            .Select(item => $"{item.CodigoCuenta.Trim()}|{item.NombreCuenta.Trim()}|01"));
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(contenido)));
    }

    private static bool EsCuentaPlanExportable(PlanCuentaDto cuenta)
    {
        var codigo = cuenta.CodigoCuenta.Trim();
        return cuenta.Estado
            && cuenta.EsUltimoNivel
            && codigo.Length is >= 3 and <= 24
            && codigo.All(char.IsDigit);
    }

    private async Task<int> ContarMovimientosAsync(LibroElectronicoConsultaRequest request, CancellationToken cancellationToken)
    {
        return request.LibroElectronico switch
        {
            PleLibroElectronicoCatalogo.LibroDiario52 => (await libroDiario52Service.ListarAsync(request, cancellationToken)).Count,
            PleLibroElectronicoCatalogo.LibroMayor61 => (await libroMayor61Service.ListarAsync(request, cancellationToken)).Count,
            _ => (await libroDiario51Service.ListarAsync(request, cancellationToken)).Count
        };
    }

    private async Task<(bool Bloqueada, string Mensaje)> ObtenerBloqueoGeneracionAsync(
        LibroElectronicoConsultaRequest request,
        PlePresentacionContextoDto presentacion,
        CancellationToken cancellationToken)
    {
        if (presentacion.Presentado)
        {
            return (true, "El periodo ya esta marcado como presentado. Desmarquelo antes de volver a generar archivos.");
        }

        if (presentacion.MesAnteriorPresentado)
        {
            return (false, string.Empty);
        }

        var fechaAnterior = new DateTime(request.Anio, request.Mes, 1).AddMonths(-1);
        var requestAnterior = new LibroElectronicoConsultaRequest
        {
            IdEmpresa = request.IdEmpresa,
            Anio = (short)fechaAnterior.Year,
            Mes = (byte)fechaAnterior.Month,
            LibroElectronico = request.LibroElectronico,
            Moneda = request.Moneda,
            Estado = request.Estado
        };
        var movimientosAnteriores = await ContarMovimientosAsync(requestAnterior, cancellationToken);
        return movimientosAnteriores > 0
            ? (true, $"No se puede generar {request.Anio:0000}-{request.Mes:00}: el periodo anterior {fechaAnterior:yyyy-MM} tiene {movimientosAnteriores} movimientos y no fue marcado como presentado.")
            : (false, string.Empty);
    }

    private static (IReadOnlyCollection<PlePlanCuentaExportItemDto> Exportacion, IReadOnlyCollection<PlePlanCuentaSnapshotItemDto> Snapshot, bool EsCompleto) PrepararPlanContable(
        IReadOnlyCollection<PlanCuentaDto> cuentas,
        string snapshotAnteriorJson,
        short anio,
        byte mes)
    {
        var periodoActual = $"{anio:0000}{mes:00}{DateTime.DaysInMonth(anio, mes):00}";
        var snapshotAnterior = string.IsNullOrWhiteSpace(snapshotAnteriorJson)
            ? []
            : JsonSerializer.Deserialize<List<PlePlanCuentaSnapshotItemDto>>(snapshotAnteriorJson) ?? [];
        var anteriores = snapshotAnterior.ToDictionary(x => x.IdPlanCuenta);
        var exportacion = new List<PlePlanCuentaExportItemDto>();
        var snapshot = new List<PlePlanCuentaSnapshotItemDto>();
        var esCompleto = snapshotAnterior.Count == 0;

        foreach (var cuenta in cuentas.OrderBy(x => x.CodigoCuenta, StringComparer.OrdinalIgnoreCase))
        {
            var codigo = cuenta.CodigoCuenta.Trim();
            var nombre = cuenta.NombreCuenta.Trim();
            var periodoInformado = periodoActual;
            var estado = "1";
            var debeExportar = esCompleto;

            if (!esCompleto && anteriores.TryGetValue(cuenta.IdPlanCuenta, out var anterior))
            {
                periodoInformado = anterior.PeriodoInformado;
                debeExportar = !string.Equals(codigo, anterior.CodigoCuenta, StringComparison.Ordinal)
                    || !string.Equals(nombre, anterior.NombreCuenta, StringComparison.Ordinal);
                estado = "9";
            }
            else if (!esCompleto)
            {
                debeExportar = true;
            }

            snapshot.Add(new PlePlanCuentaSnapshotItemDto
            {
                IdPlanCuenta = cuenta.IdPlanCuenta,
                CodigoCuenta = codigo,
                NombreCuenta = nombre,
                PeriodoInformado = periodoInformado
            });

            if (debeExportar)
            {
                exportacion.Add(new PlePlanCuentaExportItemDto
                {
                    PeriodoPle = periodoInformado,
                    CodigoCuenta = codigo,
                    NombreCuenta = nombre,
                    EstadoOperacion = estado
                });
            }
        }

        return (exportacion, snapshot, esCompleto);
    }

}
