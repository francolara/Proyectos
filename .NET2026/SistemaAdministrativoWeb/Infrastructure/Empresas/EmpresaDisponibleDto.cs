namespace SistemaAdministrativoWeb.Infrastructure.Empresas;

public sealed class EmpresaDisponibleDto
{
    public int IdEmpresa { get; set; }
    public string CodigoEmpresa { get; set; } = string.Empty;
    public string RazonSocial { get; set; } = string.Empty;
    public string? NombreComercial { get; set; }
    public string Ruc { get; set; } = string.Empty;
    public bool EsEmpresaPredeterminada { get; set; }
}
