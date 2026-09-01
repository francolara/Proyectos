namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class PlanCuentaMaestroDto
{
    public int IdPlanCuentaMaestro { get; init; }
    public string CodigoCuenta { get; init; } = string.Empty;
    public string? CodigoCuentaPadre { get; init; }
    public string NombreCuenta { get; init; } = string.Empty;
    public byte NivelCuenta { get; init; }
    public string ColBalance { get; init; } = string.Empty;
    public string IdMoneda { get; init; } = string.Empty;
    public string TipoCambio { get; init; } = string.Empty;
    public bool AceptaMovimiento { get; init; }
    public bool RequiereCentroCosto { get; init; }
    public bool Estado { get; init; }
    public int Orden { get; init; }
    public bool EsUltimoNivel { get; init; }
}

public sealed class GuardarPlanCuentaMaestroRequest
{
    public int? IdPlanCuentaMaestro { get; init; }
    public string CodigoCuenta { get; init; } = string.Empty;
    public string? CodigoCuentaPadre { get; init; }
    public string NombreCuenta { get; init; } = string.Empty;
    public string ColBalance { get; init; } = string.Empty;
    public string IdMoneda { get; init; } = string.Empty;
    public string TipoCambio { get; init; } = string.Empty;
    public bool AceptaMovimiento { get; init; }
    public bool RequiereCentroCosto { get; init; }
    public bool Estado { get; init; }
    public int Orden { get; init; }
    public string? UsuarioRegistro { get; init; }
}

public sealed class CuentaDestinoMaestroResumenDto
{
    public int IdCuentaDestinoReglaMaestro { get; init; }
    public string CodigoCuentaOrigen { get; init; } = string.Empty;
    public string? NombreCuentaOrigen { get; init; }
    public bool Activo { get; init; }
    public string? Observacion { get; init; }
    public int CantidadTramos { get; init; }
    public decimal PorcentajeTotal { get; init; }
}

public sealed class CuentaDestinoMaestroDto
{
    public int IdCuentaDestinoReglaMaestro { get; init; }
    public string CodigoCuentaOrigen { get; init; } = string.Empty;
    public string? NombreCuentaOrigen { get; init; }
    public bool Activo { get; init; }
    public string? Observacion { get; init; }
    public List<CuentaDestinoDetalleMaestroDto> Detalles { get; } = [];
}

public sealed class CuentaDestinoDetalleMaestroDto
{
    public int IdCuentaDestinoReglaDetalleMaestro { get; init; }
    public short Orden { get; init; }
    public string CodigoCuentaDestinoCargo { get; init; } = string.Empty;
    public string? NombreCuentaDestinoCargo { get; init; }
    public string CodigoCuentaDestinoAbono { get; init; } = string.Empty;
    public string? NombreCuentaDestinoAbono { get; init; }
    public decimal Porcentaje { get; init; }
    public bool Activo { get; init; }
}

public sealed class GuardarCuentaDestinoMaestroRequest
{
    public int? IdCuentaDestinoReglaMaestro { get; init; }
    public string CodigoCuentaOrigen { get; init; } = string.Empty;
    public bool Activo { get; init; }
    public string? Observacion { get; init; }
    public IReadOnlyCollection<GuardarCuentaDestinoDetalleMaestroRequest> Detalles { get; init; } = [];
    public string? UsuarioRegistro { get; init; }
}

public sealed class GuardarCuentaDestinoDetalleMaestroRequest
{
    public short Orden { get; init; }
    public string CodigoCuentaDestinoCargo { get; init; } = string.Empty;
    public string CodigoCuentaDestinoAbono { get; init; } = string.Empty;
    public decimal Porcentaje { get; init; }
    public bool Activo { get; init; }
}

public sealed class ParametroCuentaMaestroDto
{
    public int IdParametroMaestro { get; init; }
    public string TipoParametro { get; init; } = string.Empty;
    public string CodigoParametro { get; init; } = string.Empty;
    public string DescripcionParametro { get; init; } = string.Empty;
    public string? CodigoCuenta { get; init; }
    public string? NombreCuenta { get; init; }
    public bool Activo { get; init; }
}

public sealed class TipoImpuestoMaestroDto
{
    public int IdTipoImpuesto { get; init; }
    public string CodigoSunat { get; init; } = string.Empty;
    public string NombreImpuesto { get; init; } = string.Empty;
    public string? CodigoCuenta { get; init; }
    public string? NombreCuenta { get; init; }
    public bool Estado { get; init; }
}

public sealed class TipoComprobanteMaestroDto
{
    public int IdTipoComprobante { get; init; }
    public string CodigoTipoComprobante { get; init; } = string.Empty;
    public string Descripcion { get; init; } = string.Empty;
    public bool UsoCompras { get; init; }
    public bool UsoVentas { get; init; }
    public string? CodigoCuentaVentaSoles { get; init; }
    public string? NombreCuentaVentaSoles { get; init; }
    public string? CodigoCuentaVentaDolares { get; init; }
    public string? NombreCuentaVentaDolares { get; init; }
    public string? CodigoCuentaCompraSoles { get; init; }
    public string? NombreCuentaCompraSoles { get; init; }
    public string? CodigoCuentaCompraDolares { get; init; }
    public string? NombreCuentaCompraDolares { get; init; }
    public bool Estado { get; init; }
}

public sealed class AsignacionesMaestroDto
{
    public IReadOnlyCollection<ParametroCuentaMaestroDto> Parametros { get; init; } = [];
    public IReadOnlyCollection<TipoImpuestoMaestroDto> Impuestos { get; init; } = [];
    public IReadOnlyCollection<TipoComprobanteMaestroDto> Documentos { get; init; } = [];
}

public sealed class GuardarAsignacionMaestroRequest
{
    public string TipoAsignacion { get; init; } = string.Empty;
    public int IdRegistro { get; init; }
    public string? CodigoCuenta { get; init; }
    public string? CodigoCuentaVentaSoles { get; init; }
    public string? CodigoCuentaVentaDolares { get; init; }
    public string? CodigoCuentaCompraSoles { get; init; }
    public string? CodigoCuentaCompraDolares { get; init; }
    public string? UsuarioRegistro { get; init; }
}

public sealed class OrigenMaestroDto
{
    public int IdOrigenMaestro { get; init; }
    public string CodigoOrigen { get; init; } = string.Empty;
    public string NombreOrigen { get; init; } = string.Empty;
    public string ModuloOrigen { get; init; } = string.Empty;
    public bool PermiteRegistroManual { get; init; }
    public bool Estado { get; init; }
    public int Orden { get; init; }
}

public sealed class GuardarOrigenMaestroRequest
{
    public int? IdOrigenMaestro { get; init; }
    public string CodigoOrigen { get; init; } = string.Empty;
    public string NombreOrigen { get; init; } = string.Empty;
    public string ModuloOrigen { get; init; } = string.Empty;
    public bool PermiteRegistroManual { get; init; }
    public bool Estado { get; init; }
    public int Orden { get; init; }
    public string? UsuarioRegistro { get; init; }
}

public sealed class ConfiguracionContabilizacionMaestroDto
{
    public int IdConfiguracionContabilizacionMaestro { get; init; }
    public string ModuloOperacion { get; init; } = string.Empty;
    public string EscenarioOperacion { get; init; } = string.Empty;
    public string CodigoOrigen { get; init; } = string.Empty;
    public string? NombreOrigen { get; init; }
    public string Descripcion { get; init; } = string.Empty;
    public bool GeneraAsientoAutomatico { get; init; }
    public bool UsaTipoCambio { get; init; }
    public bool Activo { get; init; }
    public int Orden { get; init; }
}

public sealed class ValidacionMaestroIssueDto
{
    public string TipoMaestro { get; init; } = string.Empty;
    public string CodigoRegistro { get; init; } = string.Empty;
    public string Descripcion { get; init; } = string.Empty;
}
