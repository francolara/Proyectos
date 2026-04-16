using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
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

    [BindProperty]
    public string TipoRegistro { get; set; } = "usuario";

    [BindProperty(SupportsGet = true)]
    public string ReturnUrl { get; set; } = string.Empty;

    public WebBannerPublicoViewModel? BannerLateral { get; set; }

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
        [Compare(nameof(Password), ErrorMessage = "La contrasena y la confirmacion no coinciden.")]
        public string ConfirmPassword { get; set; } = string.Empty;
    }

    public async Task OnGetAsync(string? returnUrl = null)
    {
        ReturnUrl = returnUrl ?? Url.Content("~/");
        TipoRegistro = "usuario";
        Club = CrearClubDefault();
        AsignarCaptchaRegistroClub(Club);
        await CargarBannerLateralAsync();
    }

    public async Task<IActionResult> OnPostAsync(string? returnUrl = null)
    {
        ReturnUrl = returnUrl ?? Url.Content("~/");
        TipoRegistro = string.Equals(TipoRegistro, "club", StringComparison.OrdinalIgnoreCase)
            ? "club"
            : "usuario";

        if (TipoRegistro == "club")
        {
            return await ProcesarRegistroClubAsync();
        }

        return await ProcesarRegistroUsuarioAsync();
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
        ModelState.Remove("Club.Pais");
        ModelState.Remove("Club.ProvinciaEstado");
        ModelState.Remove("Club.Ciudad");
        ModelState.Remove("Club.Direccion");
        ModelState.Remove("Club.CaptchaCodigo");

        if (!TryValidateModel(Usuario, nameof(Usuario)))
        {
            Club = CrearClubDefault();
            AsignarCaptchaRegistroClub(Club);
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
            await CargarBannerLateralAsync();
            return Page();
        }

        logger.LogInformation("Nuevo usuario registrado desde portal publico.");
        await signInManager.SignInAsync(user, isPersistent: false);
        return LocalRedirect(ReturnUrl);
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
            await CargarBannerLateralAsync();
            return Page();
        }

        var captchaEsperado = HttpContext.Session.GetString(CaptchaRegistroClubSessionKey);
        if (string.IsNullOrWhiteSpace(captchaEsperado) ||
            !string.Equals(Club.CaptchaCodigo?.Trim(), captchaEsperado, StringComparison.OrdinalIgnoreCase))
        {
            ModelState.AddModelError("Club.CaptchaCodigo", "El codigo CAPTCHA no es valido.");
            AsignarCaptchaRegistroClub(Club);
            await CargarBannerLateralAsync();
            return Page();
        }

        try
        {
            var correo = (Club.Correo ?? string.Empty).Trim();
            var existe = await userManager.FindByEmailAsync(correo);
            if (existe is not null)
            {
                ModelState.AddModelError("Club.Correo", "Ya existe una cuenta con este correo.");
                AsignarCaptchaRegistroClub(Club);
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
                await CargarBannerLateralAsync();
                return Page();
            }

            try
            {
                await spService.HomeRegistrarClubConPruebaAsync(Club, nuevoUsuario.Id);
                await signInManager.SignInAsync(nuevoUsuario, isPersistent: false);

                var negocios = await spService.PanelListarNegociosUsuarioAsync(nuevoUsuario.Id);
                var negocioId = negocios.FirstOrDefault()?.NegocioId;

                TempData["MensajeSolicitudClub"] = "Registro completado. Tu prueba de 1 mes ya esta activa.";
                if (negocioId.HasValue)
                {
                    return RedirectToAction("Create", "Sedes", new { negocioId = negocioId.Value });
                }

                return RedirectToAction("Index", "Panel");
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
            "PasswordTooShort" => "La contrasena es muy corta. Usa al menos 6 caracteres.",
            "DuplicateEmail" => "Ya existe una cuenta registrada con este correo.",
            "DuplicateUserName" => "Ese correo/usuario ya esta en uso.",
            "InvalidEmail" => "El correo ingresado no tiene un formato valido.",
            "InvalidUserName" => "El correo/usuario contiene caracteres no permitidos.",
            _ => fallback
        };
    }
}
