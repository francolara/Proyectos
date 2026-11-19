namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class PlanCuentaItemViewModel
{
    public int IdPlanCuenta { get; init; }
    public int? IdPlanCuentaPadre { get; init; }
    public string CodigoCuenta { get; init; } = string.Empty;
    public string NombreCuenta { get; init; } = string.Empty;
    public int NivelCuenta { get; init; }
    public string ColBalance { get; init; } = string.Empty;
    public string IdMoneda { get; init; } = string.Empty;
    public string TipoCambio { get; init; } = string.Empty;
    public bool AceptaMovimiento { get; init; }
    public bool GeneraDiferenciaPorAnalisis { get; init; }
    public bool RequiereCentroCosto { get; init; }
    public bool Estado { get; init; }

    public string ColBalanceTexto => ColBalance switch
    {
        "S" => "S - Saldo",
        "I" => "I - Inventario",
        "N" => "N - Naturaleza",
        "F" => "F - Funcion",
        "R" => "R - Resultado",
        _ => "-"
    };

    public string MonedaTexto => string.IsNullOrWhiteSpace(IdMoneda) ? "-" : IdMoneda;

    public string TipoCambioTexto => TipoCambio switch
    {
        "V" => "Venta",
        "C" => "Compra",
        _ => "-"
    };

    public string DiferenciaCambioTexto => IdMoneda == "USD"
        ? (GeneraDiferenciaPorAnalisis ? "Analisis" : "Saldo")
        : "-";
}
