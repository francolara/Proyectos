using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using SistemaControlEspaciosDeportivosWeb.Models;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Areas.Identity.Pages.Account;

[AllowAnonymous]
public class RegisterModel(
    UserManager<ApplicationUser> userManager,
    SignInManager<ApplicationUser> signInManager,
    ISportCenterStoredProcedureService spService,
    ILogger<RegisterModel> logger) : PageModel
{
    private const string CaptchaRegistroClubSessionKey = "CaptchaRegistroClub";

    [BindProperty]
    public UsuarioInputModel Usuario { get; set; } = new();

    [BindProperty]
    public AltaClubSolicitudFormViewModel Club { get; set; } = CrearClubDefault();

    [BindProperty(SupportsGet = true)]
    public string? TipoRegistro { get; set; } = "usuario";

    [BindProperty(SupportsGet = true)]
    public string? ReturnUrl { get; set; } = string.Empty;

    public WebBannerPublicoViewModel? BannerLateral { get; set; }
    public List<SelectListItem> Departamentos { get; set; } = new();
    public List<SelectListItem> Provincias { get; set; } = new();
    public List<SelectListItem> Distritos { get; set; } = new();

    public class UsuarioInputModel
    {
        [Required(ErrorMessage = "El nombre es obligatorio.")]
        [StringLength(160)]
        public string NombreCompleto { get; set; } = string.Empty;

        [Required(ErrorMessage = "El correo es obligatorio.")]
        [EmailAddress(ErrorMessage = "Ingresa un correo valido.")]
        public string Email { get; set; } = string.Empty;

        [Phone(ErrorMessage = "Ingresa un telefono valido.")]
        [StringLength(30)]
        public string? Telefono { get; set; }

        [Required(ErrorMessage = "La contrasena es obligatoria.")]
        [StringLength(100, ErrorMessage = "La contrasena debe tener al menos {2} y como maximo {1} caracteres.", MinimumLength = 6)]
        [DataType(DataType.Password)]
        public string Password { get; set; } = string.Empty;

        [DataType(DataType.Password)]
        [Required(ErrorMessage = "La confirmacion de contrasena es obligatoria.")]
        [Compare(nameof(Password), ErrorMessage = "La contrasena y la confirmacion no coinciden.")]
        public string ConfirmPassword { get; set; } = string.Empty;
    }

    public async Task OnGetAsync(string? returnUrl = null)
    {
        ReturnUrl = returnUrl ?? Url.Content("~/");
        TipoRegistro = string.Equals(TipoRegistro, "club", StringComparison.OrdinalIgnoreCase) ? "club" : "usuario";
        Club = CrearClubDefault();
        AsignarCaptchaRegistroClub(Club);
        await CargarCombosUbigeoAsync();
        await CargarBannerLateralAsync();
    }

    public async Task<IActionResult> OnPostAsync(string? returnUrl = null)
    {
        // Fallback cuando el navegador envia Enter sin handler explicito.
        var accionForm = (Request.Form["accionRegistro"].ToString() ?? string.Empty).Trim();
        var tipoForm = string.IsNullOrWhiteSpace(accionForm)
            ? (Request.Form["TipoRegistro"].ToString() ?? TipoRegistro ?? string.Empty).Trim()
            : accionForm;
        if (string.Equals(tipoForm, "club", StringComparison.OrdinalIgnoreCase))
        {
            return await OnPostClubAsync(returnUrl);
        }

        return await OnPostUsuarioAsync(returnUrl);
    }

    public async Task<IActionResult> OnPostUsuarioAsync(string? returnUrl = null)
    {
        ReturnUrl = returnUrl ?? Url.Content("~/");
        TipoRegistro = "usuario";
        return await ProcesarRegistroUsuarioAsync();
    }

    public async Task<IActionResult> OnPostClubAsync(string? returnUrl = null)
    {
        ReturnUrl = returnUrl ?? Url.Content("~/");
        TipoRegistro = "club";
        return await ProcesarRegistroClubAsync();
    }

    private async Task<IActionResult> ProcesarRegistroUsuarioAsync()
    {
        ModelState.Remove("Club.NombreContacto");
        ModelState.Remove("Club.Telefono");
        ModelState.Remove("Club.Correo");
        ModelState.Remove("Club.Password");
        ModelState.Remove("Club.ConfirmPassword");
        ModelState.Remove("Club.RelacionClub");
        ModelState.Remove("Club.NombreClub");
        ModelState.Remove("Club.CodigoDepartamento");
        ModelState.Remove("Club.CodigoProvincia");
        ModelState.Remove("Club.CodigoUbigeo");
        ModelState.Remove("Club.Pais");
        ModelState.Remove("Club.ProvinciaEstado");
        ModelState.Remove("Club.Ciudad");
        ModelState.Remove("Club.Direccion");
        ModelState.Remove("Club.CaptchaTexto");
        ModelState.Remove("Club.CaptchaCodigo");

        if (!TryValidateModel(Usuario, nameof(Usuario)))
        {
            Club = CrearClubDefault();
            AsignarCaptchaRegistroClub(Club);
            await CargarCombosUbigeoAsync();
            await CargarBannerLateralAsync();
            return Page();
        }

        var email = (Usuario.Email ?? string.Empty).Trim();
        var existing = await userManager.FindByEmailAsync(email);
        if (existing is not null)
        {
            ModelState.AddModelError(string.Empty, "Ya existe una cuenta registrada con este correo.");
            Club = CrearClubDefault();
            AsignarCaptchaRegistroClub(Club);
            await CargarCombosUbigeoAsync();
            await CargarBannerLateralAsync();
            return Page();
        }

        var user = new ApplicationUser
        {
            UserName = email,
            Email = email,
            Nombres = (Usuario.NombreCompleto ?? string.Empty).Trim(),
            PhoneNumber = string.IsNullOrWhiteSpace(Usuario.Telefono) ? null : Usuario.Telefono.Trim()
        };

        var result = await userManager.CreateAsync(user, Usuario.Password);
        if (!result.Succeeded)
        {
            foreach (var error in result.Errors)
            {
                ModelState.AddModelError(string.Empty, TraducirErrorIdentity(error.Code, error.Description));
            }

            Club = CrearClubDefault();
            AsignarCaptchaRegistroClub(Club);
            await CargarCombosUbigeoAsync();
            await CargarBannerLateralAsync();
            return Page();
        }

        logger.LogInformation("Nuevo usuario registrado desde portal publico.");
        try
        {
            await spService.UsuariosPublicosGuardarPerfilAsync(new UsuarioPublicoPerfilViewModel
            {
                UsuarioId = user.Id,
                TipoDocumento = "0",
                Nombres = (Usuario.NombreCompleto ?? string.Empty).Trim(),
                Apellidos = string.Empty,
                Telefono = string.IsNullOrWhiteSpace(Usuario.Telefono) ? null : Usuario.Telefono.Trim(),
                Correo = email
            }, email);
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "No se pudo sincronizar el perfil publico inicial para usuario {Email}.", email);
        }

        await signInManager.SignInAsync(user, isPersistent: false);
        return LocalRedirect(ReturnUrl ?? Url.Content("~/"));
    }

    private async Task<IActionResult> ProcesarRegistroClubAsync()
    {
        ModelState.Remove("Usuario.NombreCompleto");
        ModelState.Remove("Usuario.Email");
        ModelState.Remove("Usuario.Telefono");
        ModelState.Remove("Usuario.Password");
        ModelState.Remove("Usuario.ConfirmPassword");

        if (!TryValidateModel(Club, nameof(Club)))
        {
            AsignarCaptchaRegistroClub(Club);
            await CargarCombosUbigeoAsync();
            await CargarBannerLateralAsync();
            return Page();
        }

        var captchaEsperado = HttpContext.Session.GetString(CaptchaRegistroClubSessionKey);
        if (string.IsNullOrWhiteSpace(captchaEsperado) ||
            !string.Equals(Club.CaptchaCodigo?.Trim(), captchaEsperado, StringComparison.OrdinalIgnoreCase))
        {
            ModelState.AddModelError("Club.CaptchaCodigo", "El codigo CAPTCHA no es valido.");
            AsignarCaptchaRegistroClub(Club);
            await CargarCombosUbigeoAsync();
            await CargarBannerLateralAsync();
            return Page();
        }

        try
        {
            var ubigeo = await spService.UbigeoObtenerPorCodigoAsync((Club.CodigoUbigeo ?? string.Empty).Trim());
            if (ubigeo is null)
            {
                ModelState.AddModelError("Club.CodigoUbigeo", "Selecciona un distrito valido.");
                AsignarCaptchaRegistroClub(Club);
                await CargarCombosUbigeoAsync();
                await CargarBannerLateralAsync();
                return Page();
            }

            Club.Pais = "Peru";
            Club.ProvinciaEstado = ubigeo.Provincia;
            Club.Ciudad = ubigeo.Distrito;

            var correo = (Club.Correo ?? string.Empty).Trim();
            var existe = await userManager.FindByEmailAsync(correo);
            if (existe is not null)
            {
                ModelState.AddModelError("Club.Correo", "Ya existe una cuenta con este correo.");
                AsignarCaptchaRegistroClub(Club);
                await CargarCombosUbigeoAsync();
                await CargarBannerLateralAsync();
                return Page();
            }

            var nuevoUsuario = new ApplicationUser
            {
                UserName = correo,
                Email = correo,
                Nombres = (Club.NombreContacto ?? string.Empty).Trim()
            };

            var resultadoCreacion = await userManager.CreateAsync(nuevoUsuario, Club.Password);
            if (!resultadoCreacion.Succeeded)
            {
                foreach (var error in resultadoCreacion.Errors)
                {
                    ModelState.AddModelError(string.Empty, TraducirErrorIdentity(error.Code, error.Description));
                }

                AsignarCaptchaRegistroClub(Club);
                await CargarCombosUbigeoAsync();
                await CargarBannerLateralAsync();
                return Page();
            }

            try
            {
                await spService.HomeSolicitarAltaClubAsync(Club);
                TempData["SuccessMessage"] = "Registro completado correctamente. Tu solicitud fue recibida y sera evaluada por el equipo de plataforma. Te contactaremos por WhatsApp o correo para la activacion.";
                return RedirectToPage("./Login", new { ReturnUrl });
            }
            catch
            {
                await userManager.DeleteAsync(nuevoUsuario);
                throw;
            }
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            AsignarCaptchaRegistroClub(Club);
            await CargarCombosUbigeoAsync();
            await CargarBannerLateralAsync();
            return Page();
        }
    }

    private void AsignarCaptchaRegistroClub(AltaClubSolicitudFormViewModel model)
    {
        const string chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
        var captcha = new string(Enumerable.Range(0, 5)
            .Select(_ => chars[Random.Shared.Next(chars.Length)])
            .ToArray());

        HttpContext.Session.SetString(CaptchaRegistroClubSessionKey, captcha);
        model.CaptchaTexto = captcha;
        model.CaptchaCodigo = string.Empty;
    }

    private static AltaClubSolicitudFormViewModel CrearClubDefault()
    {
        return new AltaClubSolicitudFormViewModel
        {
            Pais = "Peru",
            RelacionClub = "Dueno"
        };
    }

    private async Task CargarCombosUbigeoAsync()
    {
        Departamentos = await spService.UbigeoDepartamentosListarAsync();
        Provincias = !string.IsNullOrWhiteSpace(Club.CodigoDepartamento) && Club.CodigoDepartamento.Length == 2
            ? await spService.UbigeoProvinciasListarAsync(Club.CodigoDepartamento)
            : new List<SelectListItem>();
        Distritos = !string.IsNullOrWhiteSpace(Club.CodigoProvincia) && Club.CodigoProvincia.Length == 4
            ? await spService.UbigeoDistritosListarAsync(Club.CodigoProvincia)
            : new List<SelectListItem>();
    }

    private async Task CargarBannerLateralAsync()
    {
        try
        {
            BannerLateral = await spService.WebBannersObtenerFijoPorTipoAsync((int)BannerTipo.Registro);
        }
        catch
        {
            BannerLateral = null;
        }
    }

    private static string TraducirErrorIdentity(string code, string fallback)
    {
        return code switch
        {
            "PasswordRequiresNonAlphanumeric" => "La contrasena debe incluir al menos un simbolo (por ejemplo: !, @, #).",
            "PasswordRequiresLower" => "La contrasena debe incluir al menos una letra minuscula (a-z).",
            "PasswordRequiresUpper" => "La contrasena debe incluir al menos una letra mayuscula (A-Z).",
            "PasswordRequiresDigit" => "La contrasena debe incluir al menos un numero (0-9).",
            "PasswordRequiresUniqueChars" => "La contrasena debe incluir mas caracteres distintos.",
            "PasswordTooShort" => "La contrasena es muy corta. Usa al menos 6 caracteres.",
            "DuplicateEmail" => "Ya existe una cuenta registrada con este correo.",
            "DuplicateUserName" => "Ese correo/usuario ya esta en uso.",
            "InvalidEmail" => "El correo ingresado no tiene un formato valido.",
            "InvalidUserName" => "El correo/usuario contiene caracteres no permitidos.",
            _ => "No se pudo completar el registro. Revisa los datos ingresados e intenta nuevamente."
        };
    }
}
