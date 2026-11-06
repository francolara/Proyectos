using System.ComponentModel.DataAnnotations;
namespace GestionServicios.Web.ViewModels;
// Firma: FRANCO LARA - 03/11/2026 | Modelo MVC del acceso privado HYT.
// Firma: FRANCO LARA - 05/11/2026 | Desafío manual de un intento y selección del respaldo en el acceso.
public sealed class AccessViewModel
{
    [Required(ErrorMessage = "Ingresa tu usuario."), StringLength(256)] public string UserName { get; set; } = "";
    [Required(ErrorMessage = "Ingresa tu contraseña."), StringLength(1024)] public string Password { get; set; } = "";
    public bool UseManualCaptcha { get; set; }
    public Guid? CaptchaId { get; set; }
    [StringLength(12, ErrorMessage="Revisa los caracteres del captcha.")] public string? CaptchaAnswer { get; set; }
    public string? ReturnUrl { get; set; }
}
