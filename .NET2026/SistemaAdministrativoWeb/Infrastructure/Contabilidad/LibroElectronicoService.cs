namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

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

        byte[] contenido = request.LibroElectronico switch
        {
            PleLibroElectronicoCatalogo.LibroDiario52 => await pleTxtGenerator.GenerarLibroDiario52Async(consulta.LibroDiario52Items, cancellationToken),
            PleLibroElectronicoCatalogo.LibroMayor61 => await pleTxtGenerator.GenerarLibroMayor61Async(consulta.LibroMayor61Items, cancellationToken),
            _ => await pleTxtGenerator.GenerarLibroDiario51Async(consulta.LibroDiario51Items, cancellationToken)
        };

        var token = pleDownloadStore.Guardar(consulta.Resumen.NombreArchivo, contenido);
        await libroElectronicoRepository.RegistrarHistorialAsync(new PleHistorialRegistroRequest
        {
            IdEmpresa = request.IdEmpresa,
            Periodo = PlePeriodoHelper.FormarPeriodoContable(request.Anio, request.Mes),
            CodigoLibro = request.LibroElectronico,
            CodigoFormato = PleLibroElectronicoCatalogo.ObtenerCodigoSunat(request.LibroElectronico),
            NombreArchivo = consulta.Resumen.NombreArchivo,
            CantidadRegistros = consulta.Resumen.CantidadMovimientos,
            TotalDebe = consulta.Resumen.TotalDebe,
            TotalHaber = consulta.Resumen.TotalHaber,
            Estado = "GENERADO",
            Observaciones = consulta.Validacion.CantidadAdvertencias > 0 ? $"{consulta.Validacion.CantidadAdvertencias} advertencias." : string.Empty,
            UsuarioGeneracion = usuarioGeneracion
        }, cancellationToken);

        var consultaActualizada = await ConstruirResultadoAsync(request, empresa, ruc, paginaPreview, tamanoPaginaPreview, paginaHistorial, tamanoPaginaHistorial, ejecutarValidacion: true, cancellationToken);

        return new PleGenerationResultDto
        {
            Generado = true,
            Mensaje = "El archivo fue generado correctamente.",
            TokenDescarga = token,
            NombreArchivo = consulta.Resumen.NombreArchivo,
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

    private async Task<PleConsultaResultadoDto> ConstruirResultadoAsync(LibroElectronicoConsultaRequest request, string empresa, string ruc, int paginaPreview, int tamanoPaginaPreview, int paginaHistorial, int tamanoPaginaHistorial, bool ejecutarValidacion, CancellationToken cancellationToken)
    {
        var libro = PleLibroElectronicoCatalogo.Normalizar(request.LibroElectronico);
        var periodoContable = PlePeriodoHelper.FormarPeriodoContable(request.Anio, request.Mes);
        var nombreArchivo = pleFileNameService.ConstruirNombreArchivo(ruc, request.Anio, request.Mes, libro, request.Moneda);

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

        var historial = await libroElectronicoRepository.ListarHistorialAsync(request.IdEmpresa, request.Anio, request.Mes, libro, paginaHistorial, tamanoPaginaHistorial, cancellationToken);
        var asientos = await asientoRepository.ListarPorEmpresaAsync(request.IdEmpresa, periodoContable, false, cancellationToken);
        var cuentas = await planCuentaRepository.ListarPorEmpresaAsync(request.IdEmpresa, false, cancellationToken);
        var monedas = await monedaRepository.ListarActivasAsync(cancellationToken);
        var tiposDocumento = await personaRepository.ListarTiposDocumentoAsync(cancellationToken);
        var tiposComprobante = await tipoComprobanteRepository.ListarActivosAsync(true, true, cancellationToken);

        var validacion = ejecutarValidacion
            ? await pleValidationService.ValidarAsync(request, empresa, ruc, asientos, cuentas, monedas, tiposDocumento, tiposComprobante, items51, items52, items61, cancellationToken)
            : new PleValidationResultDto();

        var totalMovimientos = items51.Count + items52.Count + items61.Count;
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
            Historial = historial
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
}
