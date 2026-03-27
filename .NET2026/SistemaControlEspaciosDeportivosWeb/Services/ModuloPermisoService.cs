using System.Data;
using System.Security.Claims;
using Microsoft.Data.SqlClient;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public interface IModuloPermisoService
{
    Task<ModuloPermisoContexto> ObtenerContextoAsync(ClaimsPrincipal user, int negocioId, string moduloCodigo);
}

public class ModuloPermisoService(IConfiguration configuration) : IModuloPermisoService
{
    private readonly string _connectionString = configuration.GetConnectionString("DefaultConnection")
                                               ?? throw new InvalidOperationException("No se encontro DefaultConnection.");

    public async Task<ModuloPermisoContexto> ObtenerContextoAsync(ClaimsPrincipal user, int negocioId, string moduloCodigo)
    {
        var usuarioId = user.FindFirstValue(ClaimTypes.NameIdentifier);
        if (string.IsNullOrWhiteSpace(usuarioId))
        {
            return ModuloPermisoContexto.SinAcceso("No se pudo identificar al usuario.");
        }

        await using var cn = new SqlConnection(_connectionString);
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Seguridad_ObtenerContextoModulo", cn)
        {
            CommandType = CommandType.StoredProcedure
        };
        cmd.Parameters.Add("@UsuarioId", SqlDbType.NVarChar).Value = usuarioId;
        cmd.Parameters.Add("@NegocioId", SqlDbType.Int).Value = negocioId;
        cmd.Parameters.Add("@ModuloCodigo", SqlDbType.NVarChar).Value = moduloCodigo;

        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync())
        {
            return ModuloPermisoContexto.SinAcceso("No se encontro contexto de acceso.");
        }

        var autorizado = dr.GetBoolean(0);
        if (!autorizado)
        {
            return ModuloPermisoContexto.SinAcceso(dr.IsDBNull(10) ? "No autorizado." : dr.GetString(10));
        }

        return new ModuloPermisoContexto
        {
            Autorizado = dr.GetBoolean(0),
            NegocioId = dr.GetInt32(1),
            NegocioNombre = dr.GetString(2),
            ModuloCodigo = dr.GetString(3),
            ModuloNombre = dr.GetString(4),
            RolActual = dr.GetString(5),
            PuedeVer = dr.GetBoolean(6),
            PuedeCrear = dr.GetBoolean(7),
            PuedeEditar = dr.GetBoolean(8),
            PuedeEliminar = dr.GetBoolean(9),
            Mensaje = dr.IsDBNull(10) ? string.Empty : dr.GetString(10)
        };
    }
}

public class ModuloPermisoContexto
{
    public bool Autorizado { get; set; }
    public string Mensaje { get; set; } = string.Empty;
    public int NegocioId { get; set; }
    public string NegocioNombre { get; set; } = string.Empty;
    public string ModuloCodigo { get; set; } = string.Empty;
    public string ModuloNombre { get; set; } = string.Empty;
    public string RolActual { get; set; } = string.Empty;
    public bool PuedeVer { get; set; }
    public bool PuedeCrear { get; set; }
    public bool PuedeEditar { get; set; }
    public bool PuedeEliminar { get; set; }

    public static ModuloPermisoContexto SinAcceso(string mensaje) => new()
    {
        Autorizado = false,
        Mensaje = mensaje
    };
}
