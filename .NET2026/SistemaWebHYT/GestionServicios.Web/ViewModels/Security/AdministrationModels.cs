using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Text.Json;
using GestionServicios.Application.Security;

namespace GestionServicios.Web.ViewModels.Security;

public sealed record UserListModel(SecurityPage Data, string? Search, bool? Active, string? Situation);
public sealed record ProfileListModel(SecurityPage Data, string? Search);
public sealed record UserFormModel(UserInput Input, IReadOnlyList<JsonElement> Roles, JsonElement? Record, bool ReadOnly);
public sealed record ProfileFormModel(ProfileInput Input, IReadOnlyList<JsonElement> Catalog, JsonElement? Record, bool ReadOnly);
public sealed record UserActionModel(string Id, string Name, string Email, string Action, UserActionInput Input);
public sealed record TemporaryPasswordModel(string Email, string Password);
public sealed record AuditListModel(SecurityAuditPage Result, DateOnly From, DateOnly To, string? Search, string? Actor, string? Entity, string? Action);

public sealed class UserInput
{
    [StringLength(128)] public string? Version { get; set; }
    [Required, StringLength(150)] public string Names { get; set; } = "";
    [Required, StringLength(150)] public string Surnames { get; set; } = "";
    [Required, EmailAddress, StringLength(256)] public string Email { get; set; } = "";
    [Required, StringLength(128)] public string RoleId { get; set; } = "";
    [StringLength(128, MinimumLength = 12, ErrorMessage = "La contraseña temporal debe tener al menos 12 caracteres.")]
    public string? TemporaryPassword { get; set; }
    public bool Active { get; set; } = true;
}

public sealed class UserActionInput
{
    [Required, StringLength(128)] public string Version { get; set; } = "";
    [Required(ErrorMessage = "Indica el motivo de la acción."), StringLength(1000)] public string Reason { get; set; } = "";
    [StringLength(128, MinimumLength = 12, ErrorMessage = "La contraseña temporal debe tener al menos 12 caracteres.")]
    public string? TemporaryPassword { get; set; }
}

public sealed class ProfileInput
{
    [StringLength(128)] public string? Version { get; set; }
    [Required, StringLength(256)] public string Name { get; set; } = "";
    public string[] Permissions { get; set; } = [];
}

public static class SecurityDisplay
{
    public static string Text(this JsonElement row, string name, string fallback = "—") =>
        row.ValueKind == JsonValueKind.Object && row.TryGetProperty(name, out var value) &&
        value.ValueKind != JsonValueKind.Null ? value.ToString() : fallback;

    public static bool Flag(this JsonElement row, string name) =>
        row.ValueKind == JsonValueKind.Object && row.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.True;

    public static string Date(this JsonElement row, string name, string fallback = "—") =>
        DateTimeOffset.TryParse(row.Text(name), CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var value)
            ? value.ToOffset(TimeSpan.FromHours(-5)).ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture) : fallback;

    public static string Name(JsonElement row) => string.Join(" ", new[] { row.Text("Nombres", ""), row.Text("Apellidos", "") }.Where(s => s.Length > 0));

    public static string Module(string code) => code.ToUpperInvariant() switch
    {
        "ASPNETUSERS" or "USUARIOS" or "USUARIO" => "Usuarios",
        "ASPNETROLES" or "PERFILES" or "PERFIL" => "Perfiles de acceso",
        "CLIENTES" or "CLIENTE" => "Clientes y contactos",
        "TIENDAS" or "TIENDA" => "Tiendas",
        "MATERIALES" or "MATERIAL" or "SERVICIOMATERIAL" => "Materiales y servicios",
        "UNIDADES" => "Unidades de medida",
        "ZONAS" => "Zonas",
        "TARIFARIOS" => "Tarifarios",
        "PRESUPUESTO" or "PRESUPUESTOS" or "PRESUPUESTOREVISION" => "Presupuestos",
        "ORDENCOMPRACLIENTE" or "OC_CLIENTE" => "Orden de Compra - Cliente",
        "GUIACONFORMIDAD" or "GUIAS" => "Guía",
        "CONFIRMACIONES" or "CONFIRMACIONCLIENTE" => "Confirmación del cliente",
        "EJECUCION" or "EJECUCIONTRABAJO" => "Ejecución del trabajo",
        "FACTURADO" or "REFERENCIAFACTURACION" => "Facturado",
        "CONFIGURACION" or "CONFIGURACIONHYT" => "Configuración HYT",
        "AUDITORIA" => "Auditoría",
        "REPORTES" => "Reportes",
        "INICIO" => "Inicio",
        "ACCESO" or "ACCESOUSUARIO" => "Inicio de sesión",
        _ => Humanize(code)
    };

    public static string Permission(string code) => code.Split('.').Last() switch
    {
        "CONSULTAR" => "Consultar",
        "CREAR" => "Crear",
        "MODIFICAR" or "EDITAR" => "Editar",
        "ACTIVAR" => "Activar o desactivar",
        "CAMBIAR_ESTADO" => "Cambiar estado",
        "IMPRIMIR" => "Imprimir",
        "DESCARGAR" => "Descargar",
        "RESTABLECER" => "Restablecer contraseña",
        "DESBLOQUEAR" => "Desbloquear acceso",
        "SESIONES" => "Cerrar sesiones",
        "MARCAR" => "Marcar como facturado",
        var action => Humanize(action)
    };

    public static string Action(string code) => code switch
    {
        "CREAR_USUARIO" => "Crear usuario y emitir contraseña temporal",
        "CAMBIAR_CLAVE_OBLIGATORIA" => "Cambiar contraseña temporal",
        "CAMBIAR_CLAVE" => "Cambiar contraseña propia",
        "USUARIO_ACTIVAR" => "Activar usuario",
        "USUARIO_DESACTIVAR" => "Desactivar usuario",
        "USUARIO_RESTABLECER" => "Restablecer contraseña",
        "USUARIO_DESBLOQUEAR" => "Desbloquear acceso",
        "USUARIO_CERRAR_SESIONES" => "Cerrar sesiones",
        "BLOQUEO_AUTOMATICO" => "Bloqueo por intentos",
        "INICIO_SESION" => "Iniciar sesión",
        "CIERRE_SESION" => "Cerrar sesión",
        "ACCESO_RECHAZADO" => "Intento de acceso rechazado",
        "RESTABLECER_ACCESO" => "Restablecer contraseña",
        "CAMBIAR_ESTADO_AUTOMATICO" => "Cambio automático de estado",
        _ => Humanize(code)
    };

    public static string Humanize(string value) =>
        CultureInfo.GetCultureInfo("es-PE").TextInfo.ToTitleCase(value.Replace('_', ' ').ToLowerInvariant());
}
