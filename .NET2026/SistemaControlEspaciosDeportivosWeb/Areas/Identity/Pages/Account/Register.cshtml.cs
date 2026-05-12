using System.ComponentModel.DataAnnotations;
using System.Text;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.AspNetCore.Mvc.ModelBinding;
using Microsoft.AspNetCore.Mvc.ModelBinding.Validation;
using Microsoft.AspNetCore.WebUtilities;
using SistemaControlEspaciosDeportivosWeb.Models;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Areas.Identity.Pages.Account;

[AllowAnonymous]
public class RegisterModel(
    UserManager<ApplicationUser> userManager,
    SignInManager<ApplicationUser> signInManager,
    ISportCenterStoredProcedureService spService,
    IAccountEmailService accountEmailService,
    IClubRegistrationNotificationService clubRegistrationNotificationService,
    ITurnstileValidationService turnstileValidationService,
    Microsoft.Extensions.Options.IOptions<CloudflareTurnstileSettings> turnstileOptions,
    ILogger<RegisterModel> logger) : PageModel
{
    [BindProperty]
    [ValidateNever]
    public UsuarioInputModel Usuario { get; set; } = new();

    [BindProperty]
    [ValidateNever]
    public AltaClubSolicitudFormViewModel Club { get; set; } = CrearClubDefault();

    [BindProperty(SupportsGet = true)]
    public string? TipoRegistro { get; set; } = "usuario";

    [BindProperty(SupportsGet = true)]
    public string? ReturnUrl { get; set; } = string.Empty;

    public WebBannerPublicoViewModel? BannerLateral { get; set; }
    public IList<AuthenticationScheme> ExternalLogins { get; set; } = new List<AuthenticationScheme>();
    public List<SelectListItem> Departamentos { get; set; } = new();
    public List<SelectListItem> Provincias { get; set; } = new();
    public List<SelectListItem> Distritos { get; set; } = new();
    public string TurnstileSiteKey { get; private set; } = string.Empty;

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
        [StringLength(100, ErrorMessage = "La contrasena debe tener al menos {2} y como maximo {1} caracteres.", MinimumLength = 8)]
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
        ExternalLogins = (await signInManager.GetExternalAuthenticationSchemesAsync()).ToList();
        TurnstileSiteKey = turnstileOptions.Value.SiteKey;
        Club = CrearClubDefault();
        await CargarCombosUbigeoAsync();
        await CargarBannerLateralAsync();
    }

    public async Task<IActionResult> OnPostAsync(string? returnUrl = null)
    {
        ReturnUrl = returnUrl ?? ReturnUrl ?? Url.Content("~/");
        ExternalLogins = (await signInManager.GetExternalAuthenticationSchemesAsync()).ToList();
        TurnstileSiteKey = turnstileOptions.Value.SiteKey;
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
        ReturnUrl = returnUrl ?? ReturnUrl ?? Url.Content("~/");
        ExternalLogins = (await signInManager.GetExternalAuthenticationSchemesAsync()).ToList();
        TurnstileSiteKey = turnstileOptions.Value.SiteKey;
        TipoRegistro = "usuario";
        return await ProcesarRegistroUsuarioAsync();
    }

    public async Task<IActionResult> OnPostClubAsync(string? returnUrl = null)
    {
        ReturnUrl = returnUrl ?? ReturnUrl ?? Url.Content("~/");
        ExternalLogins = (await signInManager.GetExternalAuthenticationSchemesAsync()).ToList();
        TurnstileSiteKey = turnstileOptions.Value.SiteKey;
        TipoRegistro = "club";
        return await ProcesarRegistroClubAsync();
    }

    private async Task<IActionResult> ProcesarRegistroUsuarioAsync()
    {
        ModelState.Clear();
        ModelState.ClearValidationState(nameof(Usuario));

        if (!TryValidateModel(Usuario, nameof(Usuario)))
        {
            logger.LogWarning("Registro usuario invalido. Detalle: {Detalle}",
                string.Join(" | ", ModelState
                    .Where(x => x.Value?.Errors?.Count > 0)
                    .SelectMany(x => x.Value!.Errors.Select(e => $"{(string.IsNullOrWhiteSpace(x.Key) ? "<sin-campo>" : x.Key)}: {e.ErrorMessage}"))));
            Club = CrearClubDefault();
            await CargarCombosUbigeoAsync();
            await CargarBannerLateralAsync();
            return Page();
        }

        if (!await ValidarTurnstileAsync())
        {
            Club = CrearClubDefault();
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
            await CargarCombosUbigeoAsync();
            await CargarBannerLateralAsync();
            return Page();
        }

        logger.LogInformation("Nuevo usuario registrado desde portal publico.");
        var (nombresPerfil, apellidosPerfil) = SepararNombreCompleto(Usuario.NombreCompleto);
        try
        {
            await spService.UsuariosPublicosGuardarPerfilAsync(new UsuarioPublicoPerfilViewModel
            {
                UsuarioId = user.Id,
                TipoDocumento = "0",
                Nombres = nombresPerfil,
                Apellidos = apellidosPerfil,
                Telefono = string.IsNullOrWhiteSpace(Usuario.Telefono) ? null : Usuario.Telefono.Trim(),
                Correo = email
            }, email);
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "No se pudo sincronizar el perfil publico inicial para usuario {Email}.", email);
        }

        var correoEnviado = await IntentarEnviarCorreoConfirmacionAsync(user, email, Usuario.NombreCompleto);
        TempData["SuccessMessage"] = correoEnviado
            ? "Usuario creado satisfactoriamente. Te enviamos un correo para confirmar tu cuenta."
            : "Usuario creado satisfactoriamente, pero no pudimos enviar el correo de confirmacion. Usa la opcion de reenvio en el login.";
        return RedirectToPage("./Login", new { ReturnUrl });
    }

    private async Task<IActionResult> ProcesarRegistroClubAsync()
    {
        ModelState.Clear();
        ModelState.ClearValidationState(nameof(Club));

        // Campo legado del captcha interno: en este flujo se valida con Turnstile.
        // Se completa para evitar error generico "Este campo es obligatorio." sin campo visible.
        if (string.IsNullOrWhiteSpace(Club.CaptchaCodigo))
            Club.CaptchaCodigo = "TURNSTILE";

        if (!TryValidateModel(Club, nameof(Club)))
        {
            logger.LogWarning("Registro club invalido. Detalle: {Detalle}",
                string.Join(" | ", ModelState
                    .Where(x => x.Value?.Errors?.Count > 0)
                    .SelectMany(x => x.Value!.Errors.Select(e => $"{(string.IsNullOrWhiteSpace(x.Key) ? "<sin-campo>" : x.Key)}: {e.ErrorMessage}"))));
            await CargarCombosUbigeoAsync();
            await CargarBannerLateralAsync();
            return Page();
        }

        if (!await ValidarTurnstileAsync())
        {
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
                await CargarCombosUbigeoAsync();
                await CargarBannerLateralAsync();
                return Page();
            }

            var nuevoUsuario = new ApplicationUser
            {
                UserName = correo,
                Email = correo,
                Nombres = (Club.NombreContacto ?? string.Empty).Trim(),
                PhoneNumber = string.IsNullOrWhiteSpace(Club.Telefono) ? null : Club.Telefono.Trim()
            };

            var resultadoCreacion = await userManager.CreateAsync(nuevoUsuario, Club.Password);
            if (!resultadoCreacion.Succeeded)
            {
                foreach (var error in resultadoCreacion.Errors)
                {
                    ModelState.AddModelError(string.Empty, TraducirErrorIdentity(error.Code, error.Description));
                }

                await CargarCombosUbigeoAsync();
                await CargarBannerLateralAsync();
                return Page();
            }

            var (nombresPerfilClub, apellidosPerfilClub) = SepararNombreCompleto(Club.NombreContacto);
            try
            {
                await spService.UsuariosPublicosGuardarPerfilAsync(new UsuarioPublicoPerfilViewModel
                {
                    UsuarioId = nuevoUsuario.Id,
                    TipoDocumento = "0",
                    Nombres = nombresPerfilClub,
                    Apellidos = apellidosPerfilClub,
                    Telefono = string.IsNullOrWhiteSpace(Club.Telefono) ? null : Club.Telefono.Trim(),
                    Correo = correo
                }, correo);
            }
            catch (Exception ex)
            {
                logger.LogWarning(ex, "No se pudo sincronizar el perfil publico inicial para club {Correo}.", correo);
            }

            string codigoSolicitud;
            try
            {
                codigoSolicitud = await spService.HomeSolicitarAltaClubAsync(Club);
            }
            catch
            {
                await userManager.DeleteAsync(nuevoUsuario);
                throw;
            }

            try
            {
                await clubRegistrationNotificationService.NotifyNewClubRegistrationAsync(Club, codigoSolicitud);
            }
            catch (Exception ex)
            {
                logger.LogWarning(ex, "No se pudo enviar notificacion interna por alta de club para {Correo}.", correo);
            }

            var correoEnviado = await IntentarEnviarCorreoConfirmacionAsync(nuevoUsuario, correo, Club.NombreContacto);
            TempData["SuccessMessage"] = correoEnviado
                ? "Registro completado correctamente. Tu solicitud fue recibida y te enviamos un correo para confirmar tu cuenta."
                : "Registro completado correctamente. Tu solicitud fue recibida, pero no pudimos enviar el correo de confirmacion. Usa la opcion de reenvio en el login.";
            return RedirectToPage("./Login", new { ReturnUrl });
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            await CargarCombosUbigeoAsync();
            await CargarBannerLateralAsync();
            return Page();
        }
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

    private static (string Nombres, string Apellidos) SepararNombreCompleto(string? nombreCompleto)
    {
        var valor = (nombreCompleto ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(valor))
            return (string.Empty, string.Empty);

        var partes = valor
            .Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (partes.Length == 1)
            return (partes[0], partes[0]);

        return (partes[0], string.Join(' ', partes.Skip(1)));
    }

    private static void RemoverModelStatePorPrefijo(ModelStateDictionary modelState, string prefijo)
    {
        var keys = modelState.Keys
            .Where(k => string.Equals(k, prefijo, StringComparison.OrdinalIgnoreCase)
                     || k.StartsWith(prefijo + ".", StringComparison.OrdinalIgnoreCase)
                     || k.StartsWith(prefijo + "[", StringComparison.OrdinalIgnoreCase))
            .ToList();

        foreach (var key in keys)
            modelState.Remove(key);
    }

    private static void RemoverModelStatePorNombres(ModelStateDictionary modelState, params string[] nombres)
    {
        if (nombres is null || nombres.Length == 0) return;
        var set = new HashSet<string>(nombres, StringComparer.OrdinalIgnoreCase);
        var keys = modelState.Keys
            .Where(k => set.Contains(k))
            .ToList();
        foreach (var key in keys)
            modelState.Remove(key);
    }

    private async Task EnviarCorreoConfirmacionAsync(ApplicationUser user, string email, string? nombre)
    {
        var code = await userManager.GenerateEmailConfirmationTokenAsync(user);
        code = WebEncoders.Base64UrlEncode(Encoding.UTF8.GetBytes(code));
        var callbackUrl = Url.Page(
            "/Account/ConfirmEmail",
            pageHandler: null,
            values: new { area = "Identity", userId = user.Id, code, returnUrl = ReturnUrl },
            protocol: Request.Scheme);

        if (string.IsNullOrWhiteSpace(callbackUrl))
        {
            throw new InvalidOperationException("No se pudo construir la URL de confirmacion.");
        }

        await accountEmailService.SendConfirmationEmailAsync(email, nombre, callbackUrl);
    }

    private async Task<bool> IntentarEnviarCorreoConfirmacionAsync(ApplicationUser user, string email, string? nombre)
    {
        try
        {
            await EnviarCorreoConfirmacionAsync(user, email, nombre);
            return true;
        }
        catch (EmailDeliveryException ex)
        {
            logger.LogWarning(ex, "No se pudo enviar correo de confirmacion para {Email}.", email);
            return false;
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Error no controlado al intentar enviar correo de confirmacion para {Email}.", email);
            return false;
        }
    }

    private async Task<bool> ValidarTurnstileAsync()
    {
        var valoresToken = Request.Form["cf-turnstile-response"];
        var token = valoresToken
            .SelectMany(v => (v ?? string.Empty)
                .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            .LastOrDefault(v => !string.IsNullOrWhiteSpace(v)) ?? string.Empty;

        if (string.IsNullOrWhiteSpace(token))
        {
            ModelState.AddModelError(string.Empty, "Completa la verificacion de seguridad.");
            return false;
        }

        var remoteIp = HttpContext.Connection.RemoteIpAddress?.ToString();
        var resultado = await turnstileValidationService.VerifyAsync(token, remoteIp, HttpContext.RequestAborted);
        if (resultado.Success)
            return true;

        logger.LogWarning(
            "Turnstile rechazo registro. Errores: {Errores}. TokensRecibidos: {TotalTokens}",
            string.Join(",", resultado.ErrorCodes ?? Array.Empty<string>()),
            valoresToken.Count);
        ModelState.AddModelError(string.Empty, "No se pudo validar la verificacion de seguridad. Intenta nuevamente.");
        return false;
    }
}
