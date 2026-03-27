namespace SistemaControlEspaciosDeportivosWeb.Models;

public class RolNegocioPermiso
{
    public int Id { get; set; }
    public RolNegocio RolNegocio { get; set; }
    public int ModuloSistemaId { get; set; }
    public bool PuedeVer { get; set; }
    public bool PuedeCrear { get; set; }
    public bool PuedeEditar { get; set; }
    public bool PuedeEliminar { get; set; }

    public ModuloSistema? ModuloSistema { get; set; }
}

