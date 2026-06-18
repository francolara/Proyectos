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

        var moduloNormalizado = (moduloCodigo ?? string.Empty).Trim().ToUpperInvariant();
        var exentoPorSuscripcion = moduloNormalizado is "DASHBOARD" or "SUSCRIPCION";
        if (!exentoPorSuscripcion && await DebeBloquearPorSuscripcionAsync(negocioId))
        {
            return ModuloPermisoContexto.SinAcceso("Tu suscripcion no esta activa. Solo puedes acceder a Dashboard y Mi suscripcion.");
        }

        if (await DebeBloquearPorPlanBasicoAsync(negocioId, moduloNormalizado))
        {
            return ModuloPermisoContexto.SinAccesoPorPlanBasico("Esta funcionalidad esta disponible en el plan Full. Para usar esta opcion debes migrar de plan.");
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
            Mensaje = dr.IsDBNull(10) ? string.Empty : dr.GetString(10),
            SedeIdAsignada = dr.FieldCount > 11 && !dr.IsDBNull(11) ? dr.GetInt32(11) : null,
            EsAdministrador = dr.FieldCount > 12 && !dr.IsDBNull(12) && dr.GetBoolean(12)
        };
    }

    private async Task<bool> DebeBloquearPorSuscripcionAsync(int negocioId)
    {
        await using var cn = new SqlConnection(_connectionString);
        await cn.OpenAsync();
        await using var cmd = new SqlCommand(
            @"SELECT TOP 1
                    CAST(COALESCE(ns.EstadoSuscripcion, 0) AS INT) AS EstadoSuscripcion,
                    CAST(COALESCE(ns.EsPrueba, 0) AS BIT) AS EsPrueba,
                    ns.FechaFinPrueba,
                    ns.FechaFinPlan
              FROM dbo.NegociosSuscripcion ns
              WHERE ns.NegocioId = @NegocioId;", cn);
        cmd.Parameters.Add("@NegocioId", SqlDbType.Int).Value = negocioId;

        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync()) return true;

        var estado = dr.IsDBNull(0) ? 0 : Convert.ToInt32(dr.GetValue(0));
        var esPrueba = !dr.IsDBNull(1) && Convert.ToBoolean(dr.GetValue(1));
        var fechaFinPrueba = dr.IsDBNull(2) ? (DateTime?)null : dr.GetDateTime(2);
        var fechaFinPlan = dr.IsDBNull(3) ? (DateTime?)null : dr.GetDateTime(3);

        // Estados habilitados: 1 (activa/prueba) y 2 (contrato activo).
        if (estado is not 1 and not 2) return true;

        // Si esta en prueba, valida vigencia de la fecha fin de prueba.
        if (estado == 1 && esPrueba && fechaFinPrueba.HasValue && fechaFinPrueba.Value.Date < DateTime.Today)
            return true;

        // Si es contrato, valida vigencia de la fecha fin de plan.
        if (estado == 2 && fechaFinPlan.HasValue && fechaFinPlan.Value.Date < DateTime.Today)
            return true;

        return false;
    }

    private async Task<bool> DebeBloquearPorPlanBasicoAsync(int negocioId, string moduloCodigo)
    {
        if (moduloCodigo is not ("PROMOCIONES" or "CUPONES" or "COMPROBANTES"))
            return false;

        await using var cn = new SqlConnection(_connectionString);
        await cn.OpenAsync();
        await using var cmd = new SqlCommand(
            @"SELECT TOP 1 CAST(COALESCE(n.TipoPlan, N'Basico') AS NVARCHAR(20))
              FROM dbo.Negocios n
              WHERE n.Id = @NegocioId;", cn);
        cmd.Parameters.Add("@NegocioId", SqlDbType.Int).Value = negocioId;

        var tipoPlan = (await cmd.ExecuteScalarAsync() as string)?.Trim();
        return string.Equals(tipoPlan, "Basico", StringComparison.OrdinalIgnoreCase);
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
    public int? SedeIdAsignada { get; set; }
    public bool EsAdministrador { get; set; }
    public bool PuedeVer { get; set; }
    public bool PuedeCrear { get; set; }
    public bool PuedeEditar { get; set; }
    public bool PuedeEliminar { get; set; }
    public bool BloqueadoPorPlanBasico { get; set; }

    public static ModuloPermisoContexto SinAcceso(string mensaje) => new()
    {
        Autorizado = false,
        Mensaje = mensaje
    };

    public static ModuloPermisoContexto SinAccesoPorPlanBasico(string mensaje) => new()
    {
        Autorizado = false,
        Mensaje = mensaje,
        BloqueadoPorPlanBasico = true
    };
}
