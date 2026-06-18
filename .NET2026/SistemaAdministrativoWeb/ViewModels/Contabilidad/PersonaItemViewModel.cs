namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class PersonaItemViewModel
{
    public int IdPersona { get; set; }
    public string TipoPersona { get; set; } = string.Empty;
    public string TipoDocumento { get; set; } = string.Empty;
    public string NombreTipoDocumento { get; set; } = string.Empty;
    public string NumeroDocumento { get; set; } = string.Empty;
    public string NombreCompleto { get; set; } = string.Empty;
    public string? CorreoElectronico { get; set; }
    public string? Telefono { get; set; }
    public string? Direccion { get; set; }
    public string? Departamento { get; set; }
    public string? Provincia { get; set; }
    public string? Distrito { get; set; }
    public bool EsCliente { get; set; }
    public bool EsProveedor { get; set; }
    public bool Estado { get; set; }
}
