using System.ComponentModel.DataAnnotations;

namespace SistemaControlEspaciosDeportivosWeb.Models;

public class ComprobanteDetalle
{
    public int Id { get; set; }
    public int ComprobanteElectronicoId { get; set; }
    public int Item { get; set; }

    [MaxLength(250)]
    public string Descripcion { get; set; } = string.Empty;

    public decimal Cantidad { get; set; }

    [MaxLength(3)]
    public string UnidadMedidaSunat { get; set; } = "ZZ";

    public decimal ValorUnitario { get; set; }
    public decimal PrecioUnitario { get; set; }
    public decimal BaseIgv { get; set; }
    public decimal Igv { get; set; }
    public decimal Total { get; set; }

    [MaxLength(2)]
    public string AfectacionIgvSunat { get; set; } = "10";

    public ComprobanteElectronico? ComprobanteElectronico { get; set; }
}

