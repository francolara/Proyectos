using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Mvc.Rendering;

namespace SistemaControlEspaciosDeportivosWeb.ViewModels;

public class DesafiosIndexViewModel
{
    public UsuarioPublicoPerfilViewModel PerfilActual { get; set; } = new();
    public DesafioBusquedaViewModel Filtros { get; set; } = new();
    public DesafioCrearViewModel NuevoDesafio { get; set; } = new();
    public List<DesafioRivalItemViewModel> ResultadosBusqueda { get; set; } = new();
    public List<DesafioListadoItemViewModel> DesafiosEnviados { get; set; } = new();
    public List<DesafioListadoItemViewModel> DesafiosRecibidos { get; set; } = new();
    public List<DesafioListadoItemViewModel> Historial { get; set; } = new();
    public bool BusquedaEjecutada { get; set; }
    public int PaginaHistorial { get; set; } = 1;
    public int TamanoPaginaHistorial { get; set; } = 4;
    public int TotalHistorial { get; set; }
    public int TotalPaginasHistorial { get; set; } = 1;
}

public class DesafioBusquedaViewModel
{
    [StringLength(2)]
    public string? CodigoDepartamento { get; set; }

    [StringLength(4)]
    public string? CodigoProvincia { get; set; }

    [StringLength(30)]
    public string? Zona { get; set; }

    [StringLength(6)]
    public string? CodigoUbigeo { get; set; }

    public int? IdDeporte { get; set; }

    public int? IdNivel { get; set; }

    public List<SelectListItem> Departamentos { get; set; } = new();
    public List<SelectListItem> Provincias { get; set; } = new();
    public List<SelectListItem> Zonas { get; set; } = new();
    public List<SelectListItem> Distritos { get; set; } = new();
    public List<SelectListItem> Deportes { get; set; } = new();
    public List<SelectListItem> Niveles { get; set; } = new();
}

public class DesafioCrearViewModel
{
    public string RivalNombre { get; set; } = string.Empty;
    public int IdUsuarioRetadoPerfil { get; set; }

    [StringLength(2)]
    public string? CodigoDepartamento { get; set; }

    [StringLength(4)]
    public string? CodigoProvincia { get; set; }

    [Required(ErrorMessage = "Debes seleccionar un deporte.")]
    public int? IdDeporte { get; set; }

    [Required(ErrorMessage = "Debes seleccionar un nivel.")]
    public int? IdNivel { get; set; }

    [Required(ErrorMessage = "La fecha tentativa es obligatoria.")]
    public DateOnly? FechaTentativa { get; set; }

    [Required(ErrorMessage = "La hora tentativa es obligatoria.")]
    public TimeOnly? HoraTentativa { get; set; }

    [Required(ErrorMessage = "Debes seleccionar un distrito.")]
    [StringLength(6)]
    public string Distrito { get; set; } = string.Empty;

    public List<SelectListItem> Departamentos { get; set; } = new();
    public List<SelectListItem> Provincias { get; set; } = new();
    public List<SelectListItem> Distritos { get; set; } = new();

    [StringLength(150)]
    public string? CanchaSugerida { get; set; }

    [Required(ErrorMessage = "La modalidad es obligatoria.")]
    [StringLength(120)]
    public string Modalidad { get; set; } = string.Empty;

    [StringLength(500)]
    public string? Mensaje { get; set; }

    [Required(ErrorMessage = "La forma de pago es obligatoria.")]
    [StringLength(120)]
    public string FormaPago { get; set; } = string.Empty;
}

public class DesafioRivalItemViewModel
{
    public int PerfilId { get; set; }
    public string UsuarioId { get; set; } = string.Empty;
    public string NombreEquipo { get; set; } = string.Empty;
    public string ContactoNombre { get; set; } = string.Empty;
    public string ContactoUsuario { get; set; } = string.Empty;
    public string Distrito { get; set; } = string.Empty;
    public string Deporte { get; set; } = string.Empty;
    public string Nivel { get; set; } = string.Empty;
    public string? ObservacionDesafio { get; set; }
    public string? DetalleEquipo { get; set; }
    public int? IdDeporte { get; set; }
    public int? IdNivel { get; set; }
    public string CodigoUbigeo { get; set; } = string.Empty;
    public bool DisponibleParaDesafios { get; set; }
}

public class DesafioListadoItemViewModel
{
    public int Id { get; set; }
    public string CodigoDesafio => $"D-{Id:D6}";
    public string RivalNombre { get; set; } = string.Empty;
    public string ContactoNombreRival { get; set; } = string.Empty;
    public string ContactoUsuarioRival { get; set; } = string.Empty;
    public string RolVista { get; set; } = string.Empty;
    public string Deporte { get; set; } = string.Empty;
    public string Nivel { get; set; } = string.Empty;
    public string Distrito { get; set; } = string.Empty;
    public DateOnly FechaTentativa { get; set; }
    public TimeOnly HoraTentativa { get; set; }
    public string? CanchaSugerida { get; set; }
    public string Modalidad { get; set; } = string.Empty;
    public string? Mensaje { get; set; }
    public string FormaPago { get; set; } = string.Empty;
    public string Estado { get; set; } = string.Empty;
    public DateTime FechaCreacion { get; set; }
    public DateTime? FechaRespuesta { get; set; }
    public string? ObservacionDesafioRival { get; set; }
    public string? DetalleEquipoRival { get; set; }
    public string? TelefonoRival { get; set; }
    public string? WhatsappRival { get; set; }
    public bool PuedeVerContactoRival { get; set; }
    public List<DesafioMensajeItemViewModel> Mensajes { get; set; } = new();
    public bool PermiteMensajes => Estado == "Pendiente" || Estado == "Aceptado";
    public string? WhatsappUrlRival
    {
        get
        {
            var numero = (WhatsappRival ?? string.Empty)
                .Replace(" ", string.Empty)
                .Replace("+", string.Empty)
                .Replace("-", string.Empty)
                .Replace("(", string.Empty)
                .Replace(")", string.Empty);
            return string.IsNullOrWhiteSpace(numero) ? null : $"https://wa.me/{numero}";
        }
    }
}

public class DesafioMensajeItemViewModel
{
    public int IdMensaje { get; set; }
    public int IdDesafio { get; set; }
    public string UsuarioIdEmisor { get; set; } = string.Empty;
    public string NombreEmisor { get; set; } = string.Empty;
    public bool EsMio { get; set; }
    public string Mensaje { get; set; } = string.Empty;
    public DateTime FechaRegistro { get; set; }
}

public class DesafioMensajeCrearViewModel
{
    public int IdDesafio { get; set; }

    [Required(ErrorMessage = "Escribe un mensaje.")]
    [StringLength(500, ErrorMessage = "El mensaje no puede superar los 500 caracteres.")]
    public string Mensaje { get; set; } = string.Empty;
}

public class DesafioEmailContextViewModel
{
    public int DesafioId { get; set; }
    public string? CorreoRetado { get; set; }
    public string NombreRetado { get; set; } = string.Empty;
    public string EquipoRetador { get; set; } = string.Empty;
    public string ContactoRetador { get; set; } = string.Empty;
    public string UsuarioRetador { get; set; } = string.Empty;
    public string? TelefonoRetador { get; set; }
    public string Deporte { get; set; } = string.Empty;
    public string Nivel { get; set; } = string.Empty;
    public string Distrito { get; set; } = string.Empty;
    public DateOnly FechaTentativa { get; set; }
    public TimeOnly HoraTentativa { get; set; }
    public string? CanchaSugerida { get; set; }
    public string Modalidad { get; set; } = string.Empty;
    public string? Mensaje { get; set; }
    public string FormaPago { get; set; } = string.Empty;
}
