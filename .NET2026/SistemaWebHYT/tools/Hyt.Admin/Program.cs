using System.Data;
using System.Text.RegularExpressions;
using GestionServicios.Application.Security;
using GestionServicios.Infrastructure;
using Microsoft.AspNetCore.Identity;
using Microsoft.Data.SqlClient;

if (args.Length != 1 || args[0] is "--help" or "-h")
{
    Console.WriteLine("HYT: herramienta administrativa para ejecución MANUAL por el usuario.");
    Console.WriteLine("Comandos: crear-admin | restablecer");
    Console.WriteLine("Variables: ConnectionStrings__HYT, HYT_ADMIN_USERNAME, HYT_ADMIN_PASSWORD.");
    Console.WriteLine("Crear: HYT_ADMIN_NAME. Restablecer: HYT_ACTOR_ID y HYT_RESET_REASON.");
    Console.WriteLine("No muestra contraseñas ni conexiones. Sin comando válido no abre SQL Server.");
    return 0;
}
if (args[0] is not ("crear-admin" or "restablecer")) { Console.Error.WriteLine("Comando no reconocido. Usa --help."); return 1; }

string Required(string name) => Environment.GetEnvironmentVariable(name) is { Length: > 0 } value ? value : throw new InvalidOperationException("Falta configurar " + name);
try
{
    var userName = Required("HYT_ADMIN_USERNAME");
    var password = Required("HYT_ADMIN_PASSWORD");
    if (!Regex.IsMatch(userName, @"\A[a-zA-Z0-9._-]{3,128}\z"))
        throw new InvalidOperationException("El usuario debe tener 3–128 letras sin tildes, dígitos, punto, guion o guion bajo.");
    if (password.Length is < 12 or > 1024 || !password.Any(char.IsLower) || !password.Any(char.IsUpper) || !password.Any(char.IsDigit) || !password.Any(character => !char.IsLetterOrDigit(character)))
        throw new InvalidOperationException("La contraseña debe tener al menos 12 caracteres, mayúscula, minúscula, dígito y símbolo.");
    var displayName = args[0] == "crear-admin" ? Required("HYT_ADMIN_NAME") : userName;
    if (displayName.Length > 150) throw new InvalidOperationException("El nombre no puede exceder 150 caracteres.");
    var actor = args[0] == "restablecer" ? Required("HYT_ACTOR_ID") : null;
    var reason = args[0] == "restablecer" ? Required("HYT_RESET_REASON") : null;
    if (actor?.Length > 128 || reason?.Length > 1000) throw new InvalidOperationException("Actor o motivo fuera del tamaño permitido.");
    var id = Guid.NewGuid().ToString("N");
    var stamp = Guid.NewGuid().ToString("N");
    var account = new AccessAccount(id, userName, displayName, null, stamp, true, null, [], []);
    var hash = new PasswordHasher<AccessAccount>().HashPassword(account, password);

    // Solo este comando explícito abre conexión. La aplicación web no ejecuta bootstrap.
    await using var connection = new SqlConnectionFactory(Required("ConnectionStrings__HYT")).CreateConnection();
    await connection.OpenAsync();
    await using var command = new SqlCommand(args[0] == "crear-admin" ? "dbo.spu_CrearAdministradorInicial" : "dbo.spu_RestablecerAcceso", connection) { CommandType = CommandType.StoredProcedure };
    command.Parameters.Add("@PasswordHash", SqlDbType.NVarChar, -1).Value = hash;
    command.Parameters.Add("@SecurityStamp", SqlDbType.NVarChar, -1).Value = stamp;
    command.Parameters.Add("@CorrelacionId", SqlDbType.UniqueIdentifier).Value = Guid.NewGuid();
    if (args[0] == "crear-admin")
    {
        command.Parameters.Add("@UsuarioId", SqlDbType.NVarChar, 128).Value = id;
        command.Parameters.Add("@Usuario", SqlDbType.NVarChar, 256).Value = userName;
        command.Parameters.Add("@NombreNormalizado", SqlDbType.NVarChar, 256).Value = userName.ToUpperInvariant();
        command.Parameters.Add("@Nombre", SqlDbType.NVarChar, 150).Value = displayName;
    }
    else
    {
        command.Parameters.Add("@ActorId", SqlDbType.NVarChar, 128).Value = actor!;
        command.Parameters.Add("@UsuarioNormalizado", SqlDbType.NVarChar, 256).Value = userName.ToUpperInvariant();
        command.Parameters.Add("@Motivo", SqlDbType.NVarChar, 1000).Value = reason!;
    }
    await command.ExecuteNonQueryAsync();
    Console.WriteLine(args[0] == "crear-admin" ? "Administrador inicial creado. Identificador: " + id : "Acceso restablecido. Las sesiones anteriores se invalidarán al revalidarse.");
    return 0;
}
catch (InvalidOperationException exception) { Console.Error.WriteLine(exception.Message); return 1; }
catch (SqlException) { Console.Error.WriteLine("No se pudo completar la operación SQL. Revisa conexión, scripts, permisos y duplicados con el administrador; no se muestran detalles sensibles."); return 1; }

catch (ArgumentException) { Console.Error.WriteLine("Configuración de conexión no válida. Revisa los valores sin compartir credenciales."); return 1; }
