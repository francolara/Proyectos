namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

// Firma: FRANCO LARA - 03/08/2026 | Valida documentos y comprobantes de los formatos 5.1, 5.2 y 6.1, reconociendo 00 cuando no existe comprobante asociado.
// Firma: FRANCO LARA - 04/08/2026 | Permite periodos sin movimientos como libros sin información y conserva errores para inconsistencias reales entre asientos y líneas.
public sealed class PleValidationService(IPeriodoContableService periodoContableService) : IPleValidationService
{
    public async Task<PleValidationResultDto> ValidarAsync(
        LibroElectronicoConsultaRequest request,
        string empresa,
        string ruc,
        IReadOnlyCollection<AsientoResumenDto> asientos,
        IReadOnlyCollection<PlanCuentaDto> cuentas,
        IReadOnlyCollection<MonedaDto> monedas,
        IReadOnlyCollection<TipoDocumentoIdentidadDto> tiposDocumento,
        IReadOnlyCollection<TipoComprobanteDto> tiposComprobante,
        IReadOnlyCollection<LibroDiario51Dto> libroDiario51Items,
        IReadOnlyCollection<LibroDiario52Dto> libroDiario52Items,
        IReadOnlyCollection<LibroMayor61Dto> libroMayor61Items,
        CancellationToken cancellationToken = default)
    {
        var observaciones = new List<PleValidationIssueDto>();
        if (request.IdEmpresa <= 0)
        {
            observaciones.Add(CrearError("EMPRESA", "Empresa obligatoria", "Debe seleccionar una empresa para generar el libro electronico."));
        }

        if (string.IsNullOrWhiteSpace(empresa))
        {
            observaciones.Add(CrearError("EMPRESA_NOMBRE", "Empresa no encontrada", "No se pudo resolver la razon social de la empresa seleccionada."));
        }

        if (string.IsNullOrWhiteSpace(ruc) || ruc.Trim().Length != 11 || !ruc.Trim().All(char.IsDigit))
        {
            observaciones.Add(CrearError("RUC", "RUC invalido", "La empresa debe tener un RUC valido de 11 digitos."));
        }

        if (request.Mes is < 1 or > 12)
        {
            observaciones.Add(CrearError("PERIODO", "Mes invalido", "El periodo seleccionado no es valido."));
        }

        var estadoPeriodo = await periodoContableService.ObtenerEstadoAsync(request.IdEmpresa, request.Anio, request.Mes, cancellationToken);
        if (estadoPeriodo.Cerrado)
        {
            observaciones.Add(new PleValidationIssueDto
            {
                Severidad = PleValidationSeverity.Informacion,
                Codigo = "PERIODO_CERRADO",
                Titulo = "Periodo cerrado",
                Detalle = $"El periodo {request.Mes:00}/{request.Anio:0000} esta cerrado contablemente, pero se permite la exportacion del libro electronico."
            });
        }

        var lineas = MapearLineas(libroDiario51Items, libroDiario52Items, libroMayor61Items);
        if (asientos.Count == 0 && lineas.Count == 0)
        {
            observaciones.Add(new PleValidationIssueDto
            {
                Severidad = PleValidationSeverity.Informacion,
                Codigo = "SIN_MOVIMIENTOS",
                Titulo = "Periodo sin movimientos",
                Detalle = "No existen movimientos contables para el periodo. Se generará el libro sin información."
            });
        }
        else if (asientos.Count == 0)
        {
            observaciones.Add(CrearError("SIN_ASIENTOS", "Asientos no encontrados", "Existen líneas exportables, pero no se encontraron sus asientos contables."));
        }
        else if (lineas.Count == 0)
        {
            observaciones.Add(CrearError("SIN_LINEAS", "Sin detalle a exportar", "La consulta no devolvio lineas para el formato seleccionado."));
        }

        var periodoPle = PlePeriodoHelper.FormarPeriodo(request.Anio, request.Mes);
        var cuentasValidas = cuentas.ToDictionary(x => x.CodigoCuenta, StringComparer.OrdinalIgnoreCase);
        var cuentasMovimiento = cuentas.Where(x => x.AceptaMovimiento).Select(x => x.CodigoCuenta).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var monedasValidas = monedas.Select(x => x.CodigoMoneda).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var documentosValidos = tiposDocumento.Select(x => x.CodigoSunat).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var comprobantesValidos = tiposComprobante.Select(x => x.CodigoTipoComprobante).ToHashSet(StringComparer.OrdinalIgnoreCase);
        comprobantesValidos.Add("00");

        foreach (var asiento in asientos)
        {
            var totalDebe = Math.Round(asiento.TotalDebe, 2, MidpointRounding.AwayFromZero);
            var totalHaber = Math.Round(asiento.TotalHaber, 2, MidpointRounding.AwayFromZero);
            if (totalDebe != totalHaber)
            {
                observaciones.Add(new PleValidationIssueDto
                {
                    Severidad = PleValidationSeverity.Error,
                    Codigo = "ASIENTO_DESCUADRADO",
                    Titulo = "Asiento descuadrado",
                    Detalle = $"El asiento {asiento.NumeroAsiento} no esta cuadrado.",
                    Cuo = FormarCuo(asiento.IdAsiento),
                    NumeroAsiento = asiento.NumeroAsiento,
                    FechaOperacion = asiento.FechaAsiento,
                    TotalDebe = totalDebe,
                    TotalHaber = totalHaber,
                    Diferencia = Math.Abs(totalDebe - totalHaber)
                });
            }
        }

        var totalDebeGlobal = Math.Round(asientos.Sum(x => x.TotalDebe), 2, MidpointRounding.AwayFromZero);
        var totalHaberGlobal = Math.Round(asientos.Sum(x => x.TotalHaber), 2, MidpointRounding.AwayFromZero);
        if (totalDebeGlobal != totalHaberGlobal)
        {
            observaciones.Add(new PleValidationIssueDto
            {
                Severidad = PleValidationSeverity.Error,
                Codigo = "TOTAL_DESCUADRADO",
                Titulo = "Totales descuadrados",
                Detalle = $"El total Debe ({totalDebeGlobal:N2}) no coincide con el total Haber ({totalHaberGlobal:N2}).",
                TotalDebe = totalDebeGlobal,
                TotalHaber = totalHaberGlobal,
                Diferencia = Math.Abs(totalDebeGlobal - totalHaberGlobal)
            });
        }

        foreach (var duplicado in lineas.GroupBy(x => x.Cuo, StringComparer.OrdinalIgnoreCase).Where(x => x.Count() > 1 && x.Select(y => y.NumeroAsiento).Distinct().Count() > 1))
        {
            observaciones.Add(CrearError("CUO_DUPLICADO", "CUO duplicado", $"El CUO {duplicado.Key} se repite en mas de un asiento.", duplicado.Key));
        }

        foreach (var duplicado in lineas.GroupBy(x => $"{x.Cuo}|{x.Correlativo}", StringComparer.OrdinalIgnoreCase).Where(x => x.Count() > 1))
        {
            var muestra = duplicado.First();
            observaciones.Add(CrearError("CORRELATIVO_DUPLICADO", "Correlativo duplicado", $"El correlativo {muestra.Correlativo} se repite dentro del CUO {muestra.Cuo}.", muestra.Cuo, muestra.NumeroAsiento, muestra.FechaOperacion));
        }

        foreach (var linea in lineas)
        {
            if (!string.Equals(linea.Periodo, periodoPle, StringComparison.Ordinal))
            {
                observaciones.Add(CrearError("PERIODO_LINEA", "Periodo inconsistente", $"La linea CUO {linea.Cuo} no pertenece al periodo {periodoPle}.", linea.Cuo, linea.NumeroAsiento, linea.FechaOperacion));
            }

            if (!cuentasValidas.ContainsKey(linea.CodigoCuenta))
            {
                observaciones.Add(CrearError("CUENTA_INEXISTENTE", "Cuenta inexistente", $"La cuenta {linea.CodigoCuenta} no existe en el plan contable.", linea.Cuo, linea.NumeroAsiento, linea.FechaOperacion));
            }
            else if (!cuentasMovimiento.Contains(linea.CodigoCuenta))
            {
                observaciones.Add(CrearError("CUENTA_SIN_MOVIMIENTO", "Cuenta no operativa", $"La cuenta {linea.CodigoCuenta} no acepta movimiento.", linea.Cuo, linea.NumeroAsiento, linea.FechaOperacion));
            }

            if (!monedasValidas.Contains(linea.CodigoMoneda))
            {
                observaciones.Add(CrearError("MONEDA_INVALIDA", "Moneda invalida", $"El codigo de moneda {linea.CodigoMoneda} no es valido.", linea.Cuo, linea.NumeroAsiento, linea.FechaOperacion));
            }

            if (!string.IsNullOrWhiteSpace(linea.TipoDocumentoEmisor) && !documentosValidos.Contains(linea.TipoDocumentoEmisor))
            {
                observaciones.Add(CrearError("DOC_INVALIDO", "Tipo de documento invalido", $"El tipo de documento {linea.TipoDocumentoEmisor} no es valido.", linea.Cuo, linea.NumeroAsiento, linea.FechaOperacion));
            }

            if (!string.IsNullOrWhiteSpace(linea.TipoComprobante) && !comprobantesValidos.Contains(linea.TipoComprobante))
            {
                observaciones.Add(CrearError("COMP_INVALIDO", "Tipo de comprobante invalido", $"El tipo de comprobante {linea.TipoComprobante} no es valido.", linea.Cuo, linea.NumeroAsiento, linea.FechaOperacion));
            }

            if (linea.Debe < 0 || linea.Haber < 0)
            {
                observaciones.Add(CrearError("IMPORTE_NEGATIVO", "Importe invalido", $"La linea {linea.Correlativo} contiene importes negativos.", linea.Cuo, linea.NumeroAsiento, linea.FechaOperacion));
            }

            if (linea.Debe > 0 && linea.Haber > 0)
            {
                observaciones.Add(CrearError("DEBE_HABER", "Debe y Haber simultaneos", $"La linea {linea.Correlativo} tiene Debe y Haber con importe.", linea.Cuo, linea.NumeroAsiento, linea.FechaOperacion));
            }

            if (string.IsNullOrWhiteSpace(linea.Glosa))
            {
                observaciones.Add(CrearError("GLOSA_VACIA", "Glosa obligatoria", $"La linea {linea.Correlativo} no tiene glosa.", linea.Cuo, linea.NumeroAsiento, linea.FechaOperacion));
            }
            else if (linea.Glosa.Contains('\r') || linea.Glosa.Contains('\n'))
            {
                observaciones.Add(CrearError("GLOSA_SALTO", "Glosa invalida", $"La linea {linea.Correlativo} contiene saltos de linea.", linea.Cuo, linea.NumeroAsiento, linea.FechaOperacion));
            }

            if (!PleEstadoRegistroCatalogo.ValoresValidos.Contains(linea.EstadoOperacion))
            {
                observaciones.Add(CrearError("ESTADO_INVALIDO", "Estado invalido", $"El estado {linea.EstadoOperacion} no es valido para PLE.", linea.Cuo, linea.NumeroAsiento, linea.FechaOperacion));
            }
        }

        if (observaciones.Count == 0)
        {
            observaciones.Add(new PleValidationIssueDto
            {
                Severidad = PleValidationSeverity.Informacion,
                Codigo = "VALIDACION_OK",
                Titulo = "Validacion completa",
                Detalle = "No se encontraron observaciones que impidan generar el archivo."
            });
        }

        return new PleValidationResultDto
        {
            Observaciones = observaciones
        };
    }

    private static List<PleLineaValidacion> MapearLineas(
        IReadOnlyCollection<LibroDiario51Dto> libroDiario51Items,
        IReadOnlyCollection<LibroDiario52Dto> libroDiario52Items,
        IReadOnlyCollection<LibroMayor61Dto> libroMayor61Items)
    {
        if (libroDiario51Items.Count > 0)
        {
            return libroDiario51Items.Select(item => new PleLineaValidacion
            {
                Periodo = item.PeriodoPle,
                Cuo = item.Cuo,
                Correlativo = item.CorrelativoMovimiento,
                CodigoCuenta = item.CodigoCuentaContable,
                FechaOperacion = item.FechaOperacion,
                Glosa = item.Glosa,
                Debe = item.Debe,
                Haber = item.Haber,
                EstadoOperacion = item.EstadoOperacion,
                CodigoMoneda = item.CodigoMoneda,
                TipoDocumentoEmisor = item.TipoDocumentoEmisor,
                TipoComprobante = item.TipoComprobante,
                NumeroAsiento = item.NumeroAsiento
            }).ToList();
        }

        if (libroDiario52Items.Count > 0)
        {
            return libroDiario52Items.Select(item => new PleLineaValidacion
            {
                Periodo = item.PeriodoPle,
                Cuo = item.Cuo,
                Correlativo = item.CorrelativoAsiento,
                CodigoCuenta = item.CodigoCuentaContable,
                FechaOperacion = item.FechaOperacion,
                Glosa = item.Glosa,
                Debe = item.Debe,
                Haber = item.Haber,
                EstadoOperacion = item.EstadoOperacion,
                CodigoMoneda = item.CodigoMoneda,
                TipoDocumentoEmisor = item.TipoDocumentoEmisor,
                TipoComprobante = item.TipoComprobante,
                NumeroAsiento = item.NumeroAsiento
            }).ToList();
        }

        return libroMayor61Items.Select(item => new PleLineaValidacion
        {
            Periodo = item.PeriodoPle,
            Cuo = item.Cuo,
            Correlativo = item.CorrelativoMovimiento,
            CodigoCuenta = item.CodigoCuentaContable,
            FechaOperacion = item.FechaOperacion,
            Glosa = item.Glosa,
            Debe = item.Debe,
            Haber = item.Haber,
            EstadoOperacion = item.EstadoOperacion,
            CodigoMoneda = item.CodigoMoneda,
            TipoDocumentoEmisor = item.TipoDocumentoEmisor,
            TipoComprobante = item.TipoComprobante,
            NumeroAsiento = item.NumeroAsiento
        }).ToList();
    }

    private static string FormarCuo(int idAsiento)
    {
        return idAsiento <= 0
            ? string.Empty
            : idAsiento.ToString("00000000");
    }

    private static PleValidationIssueDto CrearError(string codigo, string titulo, string detalle, string cuo = "", int? numeroAsiento = null, DateOnly? fechaOperacion = null)
    {
        return new PleValidationIssueDto
        {
            Severidad = PleValidationSeverity.Error,
            Codigo = codigo,
            Titulo = titulo,
            Detalle = detalle,
            Cuo = cuo,
            NumeroAsiento = numeroAsiento,
            FechaOperacion = fechaOperacion
        };
    }

    private sealed class PleLineaValidacion
    {
        public string Periodo { get; init; } = string.Empty;
        public string Cuo { get; init; } = string.Empty;
        public string Correlativo { get; init; } = string.Empty;
        public string CodigoCuenta { get; init; } = string.Empty;
        public DateOnly FechaOperacion { get; init; }
        public string Glosa { get; init; } = string.Empty;
        public decimal Debe { get; init; }
        public decimal Haber { get; init; }
        public string EstadoOperacion { get; init; } = string.Empty;
        public string CodigoMoneda { get; init; } = string.Empty;
        public string TipoDocumentoEmisor { get; init; } = string.Empty;
        public string TipoComprobante { get; init; } = string.Empty;
        public int NumeroAsiento { get; init; }
    }
}
