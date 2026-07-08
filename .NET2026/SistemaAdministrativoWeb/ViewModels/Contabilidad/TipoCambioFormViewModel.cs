using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class TipoCambioFormViewModel
{
    private string _idMoneda = "USD";

    public int? IdTipoCambio { get; set; }

    [Required(ErrorMessage = "Ingrese la fecha del tipo de cambio.")]
    [DataType(DataType.Date)]
    public DateOnly Fecha { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [Required(ErrorMessage = "Seleccione la moneda.")]
    [RegularExpression("^(PEN|USD)$", ErrorMessage = "Seleccione una moneda valida.")]
    public string IdMoneda
    {
        get => _idMoneda;
        set => _idMoneda = NormalizarCodigoMoneda(value);
    }

    [Range(typeof(decimal), "0.000001", "999999999999", ErrorMessage = "Ingrese una compra mayor a cero.")]
    public decimal Compra { get; set; } = 1m;

    [Range(typeof(decimal), "0.000001", "999999999999", ErrorMessage = "Ingrese una venta mayor a cero.")]
    public decimal Venta { get; set; } = 1m;

    [Range(typeof(decimal), "0.000001", "999999999999", ErrorMessage = "Ingrese una compra SBS mayor a cero.")]
    public decimal CompraSbs { get; set; } = 1m;

    [Range(typeof(decimal), "0.000001", "999999999999", ErrorMessage = "Ingrese una venta SBS mayor a cero.")]
    public decimal VentaSbs { get; set; } = 1m;

    [Required(ErrorMessage = "Seleccione la fuente.")]
    [StringLength(50)]
    public string Fuente { get; set; } = "MANUAL";

    public bool Estado { get; set; } = true;

    private static string NormalizarCodigoMoneda(string? idMoneda)
    {
        var valor = (idMoneda ?? string.Empty).Trim().ToUpperInvariant();
        if (string.IsNullOrWhiteSpace(valor))
        {
            return string.Empty;
        }

        var separadores = new[] { " - ", "-", " " };
        foreach (var separador in separadores)
        {
            var indice = valor.IndexOf(separador, StringComparison.Ordinal);
            if (indice > 0)
            {
                valor = valor[..indice].Trim();
                break;
            }
        }

        return valor.Length > 3 ? valor[..3] : valor;
    }
}
