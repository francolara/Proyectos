using System.ComponentModel.DataAnnotations;
using SistemaAdministrativoWeb.Infrastructure.Contabilidad;
using SistemaAdministrativoWeb.ViewModels.Contabilidad;

namespace SistemaAdministrativoWeb.ViewModels.Plataforma;

public sealed class MaestroContableIndexViewModel
{
    public int TotalCuentas { get; init; }
    public int TotalReglasDestino { get; init; }
    public int TotalOrigenes { get; init; }
    public int TotalConfiguraciones { get; init; }
    public IReadOnlyCollection<ValidacionMaestroIssueDto> Incidencias { get; init; } = [];
}

public sealed class PlanCuentaMaestroIndexViewModel
{
    public string TextoBusqueda { get; init; } = string.Empty;
    public byte? NivelCuenta { get; init; }
    public IReadOnlyCollection<PlanCuentaMaestroDto> Items { get; init; } = [];
    public PaginacionViewModel Paginacion { get; init; } = new();
}

public sealed class PlanCuentaMaestroFormViewModel
{
    public int? IdPlanCuentaMaestro { get; set; }

    [Required, StringLength(20)]
    [Display(Name = "Código de cuenta")]
    public string CodigoCuenta { get; set; } = string.Empty;

    [StringLength(20)]
    [Display(Name = "Cuenta padre")]
    public string? CodigoCuentaPadre { get; set; }

    [Required, StringLength(200)]
    [Display(Name = "Nombre de cuenta")]
    public string NombreCuenta { get; set; } = string.Empty;

    [Required, StringLength(1), RegularExpression("[SINFR]")]
    [Display(Name = "Columna de balance")]
    public string ColBalance { get; set; } = "S";

    [StringLength(3), RegularExpression("^$|PEN|USD")]
    [Display(Name = "Moneda")]
    public string IdMoneda { get; set; } = string.Empty;

    [StringLength(1), RegularExpression("^$|C|V")]
    [Display(Name = "Tipo de cambio")]
    public string TipoCambio { get; set; } = string.Empty;

    public bool AceptaMovimiento { get; set; }
    public bool RequiereCentroCosto { get; set; }
    public bool Estado { get; set; } = true;
    public int Orden { get; set; }
}

public sealed class CuentaDestinoMaestroIndexViewModel
{
    public string TextoBusqueda { get; init; } = string.Empty;
    public IReadOnlyCollection<CuentaDestinoMaestroResumenDto> Items { get; init; } = [];
    public PaginacionViewModel Paginacion { get; init; } = new();
}

public sealed class CuentaDestinoMaestroFormViewModel
{
    public int? IdCuentaDestinoReglaMaestro { get; set; }

    [Required, StringLength(20)]
    [Display(Name = "Cuenta origen")]
    public string CodigoCuentaOrigen { get; set; } = string.Empty;

    [StringLength(500)]
    public string? Observacion { get; set; }

    public bool Activo { get; set; } = true;
    public List<CuentaDestinoDetalleMaestroFormViewModel> Detalles { get; set; } = [new()];
}

public sealed class CuentaDestinoDetalleMaestroFormViewModel
{
    [Range(1, short.MaxValue)]
    public short Orden { get; set; } = 1;

    [Required, StringLength(20)]
    public string CodigoCuentaDestinoCargo { get; set; } = string.Empty;

    [Required, StringLength(20)]
    public string CodigoCuentaDestinoAbono { get; set; } = string.Empty;

    [Range(typeof(decimal), "0.0001", "100")]
    public decimal Porcentaje { get; set; } = 100;

    public bool Activo { get; set; } = true;
}

public sealed class AsignacionesMaestroViewModel
{
    public IReadOnlyCollection<ParametroCuentaMaestroDto> Parametros { get; init; } = [];
    public IReadOnlyCollection<TipoImpuestoMaestroDto> Impuestos { get; init; } = [];
    public IReadOnlyCollection<TipoComprobanteMaestroDto> Documentos { get; init; } = [];
}

public sealed class AsignacionCuentaMaestroFormViewModel
{
    [Required]
    public string TipoAsignacion { get; set; } = string.Empty;
    public int IdRegistro { get; set; }
    public string? CodigoCuenta { get; set; }
    public string? CodigoCuentaVentaSoles { get; set; }
    public string? CodigoCuentaVentaDolares { get; set; }
    public string? CodigoCuentaCompraSoles { get; set; }
    public string? CodigoCuentaCompraDolares { get; set; }
}

public sealed class OrigenMaestroIndexViewModel
{
    public string TextoBusqueda { get; init; } = string.Empty;
    public IReadOnlyCollection<OrigenMaestroDto> Items { get; init; } = [];
    public PaginacionViewModel Paginacion { get; init; } = new();
}

public sealed class OrigenMaestroFormViewModel
{
    public int? IdOrigenMaestro { get; set; }

    [Required, StringLength(10)]
    [Display(Name = "Código de origen")]
    public string CodigoOrigen { get; set; } = string.Empty;

    [Required, StringLength(150)]
    [Display(Name = "Nombre de origen")]
    public string NombreOrigen { get; set; } = string.Empty;

    [Required, StringLength(50)]
    [Display(Name = "Módulo")]
    public string ModuloOrigen { get; set; } = string.Empty;

    public bool PermiteRegistroManual { get; set; } = true;
    public bool Estado { get; set; } = true;
    public int Orden { get; set; }
}

public sealed class ConfiguracionMaestroIndexViewModel
{
    public string TextoBusqueda { get; init; } = string.Empty;
    public IReadOnlyCollection<ConfiguracionContabilizacionMaestroDto> Items { get; init; } = [];
    public IReadOnlyCollection<OrigenMaestroDto> OrigenesActivos { get; init; } = [];
    public PaginacionViewModel Paginacion { get; init; } = new();
}

public sealed class ConfiguracionOrigenMaestroFormViewModel
{
    public int IdConfiguracionContabilizacionMaestro { get; set; }

    [Required, StringLength(10)]
    public string CodigoOrigen { get; set; } = string.Empty;
}
