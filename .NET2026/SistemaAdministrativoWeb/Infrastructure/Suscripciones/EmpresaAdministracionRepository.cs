using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Suscripciones;

public sealed class CuentaAdministradoraRepository(IDbConnectionFactory connectionFactory) : ICuentaAdministradoraRepository
{
    public async Task GuardarPerfilUsuarioAsync(UsuarioPerfilRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_SEG_GuardarUsuarioPerfil", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@AspNetUserId", request.AspNetUserId);
        command.Parameters.AddWithValue("@NombreCompleto", request.NombreCompleto);
        command.Parameters.AddWithValue("@Telefono", (object?)request.Telefono ?? DBNull.Value);
        command.Parameters.AddWithValue("@CorreoReferencia", (object?)request.CorreoReferencia ?? DBNull.Value);
        command.Parameters.AddWithValue("@UsuarioRegistro", (object?)request.UsuarioRegistro ?? DBNull.Value);

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task<RegistroCuentaAdministradoraConEmpresaResult> RegistrarCuentaConEmpresaAsync(RegistroCuentaAdministradoraConEmpresaRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_SEG_RegistrarCuentaAdministradoraConEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@AspNetUserId", request.AspNetUserId);
        command.Parameters.AddWithValue("@NombreCompleto", request.NombreCompleto);
        command.Parameters.AddWithValue("@Telefono", (object?)request.Telefono ?? DBNull.Value);
        command.Parameters.AddWithValue("@CorreoReferencia", request.CorreoReferencia);
        command.Parameters.AddWithValue("@CodigoCuenta", request.CodigoCuenta);
        command.Parameters.AddWithValue("@NombreCuenta", request.NombreCuenta);
        command.Parameters.AddWithValue("@CodigoEmpresa", request.CodigoEmpresa);
        command.Parameters.AddWithValue("@RazonSocial", request.RazonSocial);
        command.Parameters.AddWithValue("@NombreComercial", (object?)request.NombreComercial ?? DBNull.Value);
        command.Parameters.AddWithValue("@Ruc", request.Ruc);
        command.Parameters.AddWithValue("@DiasPrueba", request.DiasPrueba);
        command.Parameters.AddWithValue("@UsuarioRegistro", (object?)request.UsuarioRegistro ?? DBNull.Value);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        if (!await reader.ReadAsync(cancellationToken))
        {
            throw new InvalidOperationException("No se pudo obtener la cuenta registrada.");
        }

        return new RegistroCuentaAdministradoraConEmpresaResult
        {
            IdCuentaAdministradora = reader.GetInt32(reader.GetOrdinal("IdCuentaAdministradora")),
            IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
            FechaInicioPrueba = reader.IsDBNull(reader.GetOrdinal("FechaInicioPrueba"))
                ? null
                : DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaInicioPrueba"))),
            FechaFinPrueba = reader.IsDBNull(reader.GetOrdinal("FechaFinPrueba"))
                ? null
                : DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaFinPrueba")))
        };
    }

    public async Task<IReadOnlyCollection<CuentaSuscripcionResumenDto>> ListarCuentasSuscripcionAsync(CancellationToken cancellationToken = default)
    {
        var result = new List<CuentaSuscripcionResumenDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_SEG_ListarCuentasAdministradorasSuscripcion", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new CuentaSuscripcionResumenDto
            {
                IdCuentaAdministradora = reader.GetInt32(reader.GetOrdinal("IdCuentaAdministradora")),
                CodigoCuenta = reader.GetString(reader.GetOrdinal("CodigoCuenta")),
                NombreCuenta = reader.GetString(reader.GetOrdinal("NombreCuenta")),
                CorreoPrincipal = reader.GetString(reader.GetOrdinal("CorreoPrincipal")),
                TelefonoPrincipal = reader.IsDBNull(reader.GetOrdinal("TelefonoPrincipal")) ? null : reader.GetString(reader.GetOrdinal("TelefonoPrincipal")),
                EstadoCuenta = reader.GetBoolean(reader.GetOrdinal("EstadoCuenta")),
                IdEmpresaPrincipal = reader.IsDBNull(reader.GetOrdinal("IdEmpresaPrincipal")) ? null : reader.GetInt32(reader.GetOrdinal("IdEmpresaPrincipal")),
                CodigoEmpresaPrincipal = reader.IsDBNull(reader.GetOrdinal("CodigoEmpresaPrincipal")) ? null : reader.GetString(reader.GetOrdinal("CodigoEmpresaPrincipal")),
                RazonSocialEmpresaPrincipal = reader.IsDBNull(reader.GetOrdinal("RazonSocialEmpresaPrincipal")) ? null : reader.GetString(reader.GetOrdinal("RazonSocialEmpresaPrincipal")),
                NombreComercialEmpresaPrincipal = reader.IsDBNull(reader.GetOrdinal("NombreComercialEmpresaPrincipal")) ? null : reader.GetString(reader.GetOrdinal("NombreComercialEmpresaPrincipal")),
                RucEmpresaPrincipal = reader.IsDBNull(reader.GetOrdinal("RucEmpresaPrincipal")) ? null : reader.GetString(reader.GetOrdinal("RucEmpresaPrincipal")),
                CantidadEmpresas = reader.IsDBNull(reader.GetOrdinal("CantidadEmpresas")) ? 0 : reader.GetInt32(reader.GetOrdinal("CantidadEmpresas")),
                IdCuentaAdministradoraSuscripcion = reader.IsDBNull(reader.GetOrdinal("IdCuentaAdministradoraSuscripcion")) ? null : reader.GetInt32(reader.GetOrdinal("IdCuentaAdministradoraSuscripcion")),
                TipoPlan = reader.IsDBNull(reader.GetOrdinal("TipoPlan")) ? null : reader.GetString(reader.GetOrdinal("TipoPlan")),
                EstadoSuscripcion = reader.IsDBNull(reader.GetOrdinal("EstadoSuscripcion")) ? null : reader.GetString(reader.GetOrdinal("EstadoSuscripcion")),
                EsPrueba = !reader.IsDBNull(reader.GetOrdinal("EsPrueba")) && reader.GetBoolean(reader.GetOrdinal("EsPrueba")),
                FechaInicioPrueba = reader.IsDBNull(reader.GetOrdinal("FechaInicioPrueba")) ? null : DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaInicioPrueba"))),
                FechaFinPrueba = reader.IsDBNull(reader.GetOrdinal("FechaFinPrueba")) ? null : DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaFinPrueba"))),
                FechaInicioPlan = reader.IsDBNull(reader.GetOrdinal("FechaInicioPlan")) ? null : DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaInicioPlan"))),
                FechaFinPlan = reader.IsDBNull(reader.GetOrdinal("FechaFinPlan")) ? null : DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaFinPlan"))),
                EmpresasPermitidas = reader.IsDBNull(reader.GetOrdinal("EmpresasPermitidas")) ? null : reader.GetInt32(reader.GetOrdinal("EmpresasPermitidas")),
                UsuariosPermitidos = reader.IsDBNull(reader.GetOrdinal("UsuariosPermitidos")) ? null : reader.GetInt32(reader.GetOrdinal("UsuariosPermitidos")),
                Activo = !reader.IsDBNull(reader.GetOrdinal("Activo")) && reader.GetBoolean(reader.GetOrdinal("Activo")),
                Observacion = reader.IsDBNull(reader.GetOrdinal("Observacion")) ? null : reader.GetString(reader.GetOrdinal("Observacion")),
                AspNetUserId = reader.IsDBNull(reader.GetOrdinal("AspNetUserId")) ? null : reader.GetString(reader.GetOrdinal("AspNetUserId")),
                NombreCompleto = reader.IsDBNull(reader.GetOrdinal("NombreCompleto")) ? null : reader.GetString(reader.GetOrdinal("NombreCompleto")),
                Telefono = reader.IsDBNull(reader.GetOrdinal("Telefono")) ? null : reader.GetString(reader.GetOrdinal("Telefono")),
                Email = reader.IsDBNull(reader.GetOrdinal("Email")) ? null : reader.GetString(reader.GetOrdinal("Email"))
            });
        }

        return result;
    }

    public async Task ActualizarSuscripcionCuentaAsync(ActualizarSuscripcionCuentaRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_SEG_ActualizarSuscripcionCuentaAdministradora", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdCuentaAdministradora", request.IdCuentaAdministradora);
        command.Parameters.AddWithValue("@TipoPlan", request.TipoPlan);
        command.Parameters.AddWithValue("@EstadoSuscripcion", request.EstadoSuscripcion);
        command.Parameters.AddWithValue("@EsPrueba", request.EsPrueba);
        command.Parameters.AddWithValue("@FechaInicioPrueba", request.FechaInicioPrueba.HasValue ? request.FechaInicioPrueba.Value.ToDateTime(TimeOnly.MinValue) : DBNull.Value);
        command.Parameters.AddWithValue("@FechaFinPrueba", request.FechaFinPrueba.HasValue ? request.FechaFinPrueba.Value.ToDateTime(TimeOnly.MinValue) : DBNull.Value);
        command.Parameters.AddWithValue("@FechaInicioPlan", request.FechaInicioPlan.HasValue ? request.FechaInicioPlan.Value.ToDateTime(TimeOnly.MinValue) : DBNull.Value);
        command.Parameters.AddWithValue("@FechaFinPlan", request.FechaFinPlan.HasValue ? request.FechaFinPlan.Value.ToDateTime(TimeOnly.MinValue) : DBNull.Value);
        command.Parameters.AddWithValue("@EmpresasPermitidas", (object?)request.EmpresasPermitidas ?? DBNull.Value);
        command.Parameters.AddWithValue("@UsuariosPermitidos", (object?)request.UsuariosPermitidos ?? DBNull.Value);
        command.Parameters.AddWithValue("@Activo", request.Activo);
        command.Parameters.AddWithValue("@EstadoCuenta", request.EstadoCuenta);
        command.Parameters.AddWithValue("@Observacion", (object?)request.Observacion ?? DBNull.Value);
        command.Parameters.AddWithValue("@UsuarioRegistro", (object?)request.UsuarioRegistro ?? DBNull.Value);

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task<ContextoSuscripcionEmpresaDto?> ObtenerContextoSuscripcionPorEmpresaAsync(int idEmpresa, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_SEG_ObtenerContextoSuscripcionPorEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        if (!await reader.ReadAsync(cancellationToken))
        {
            return null;
        }

        return new ContextoSuscripcionEmpresaDto
        {
            IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
            IdCuentaAdministradora = reader.GetInt32(reader.GetOrdinal("IdCuentaAdministradora")),
            CodigoEmpresa = reader.GetString(reader.GetOrdinal("CodigoEmpresa")),
            RazonSocial = reader.GetString(reader.GetOrdinal("RazonSocial")),
            NombreComercial = reader.IsDBNull(reader.GetOrdinal("NombreComercial")) ? null : reader.GetString(reader.GetOrdinal("NombreComercial")),
            Ruc = reader.GetString(reader.GetOrdinal("Ruc")),
            EstadoEmpresa = reader.GetBoolean(reader.GetOrdinal("EstadoEmpresa")),
            CodigoCuenta = reader.GetString(reader.GetOrdinal("CodigoCuenta")),
            NombreCuenta = reader.GetString(reader.GetOrdinal("NombreCuenta")),
            CorreoPrincipal = reader.GetString(reader.GetOrdinal("CorreoPrincipal")),
            TelefonoPrincipal = reader.IsDBNull(reader.GetOrdinal("TelefonoPrincipal")) ? null : reader.GetString(reader.GetOrdinal("TelefonoPrincipal")),
            EstadoCuenta = reader.GetBoolean(reader.GetOrdinal("EstadoCuenta")),
            IdCuentaAdministradoraSuscripcion = reader.IsDBNull(reader.GetOrdinal("IdCuentaAdministradoraSuscripcion")) ? null : reader.GetInt32(reader.GetOrdinal("IdCuentaAdministradoraSuscripcion")),
            TipoPlan = reader.IsDBNull(reader.GetOrdinal("TipoPlan")) ? null : reader.GetString(reader.GetOrdinal("TipoPlan")),
            EstadoSuscripcion = reader.IsDBNull(reader.GetOrdinal("EstadoSuscripcion")) ? null : reader.GetString(reader.GetOrdinal("EstadoSuscripcion")),
            EsPrueba = !reader.IsDBNull(reader.GetOrdinal("EsPrueba")) && reader.GetBoolean(reader.GetOrdinal("EsPrueba")),
            FechaInicioPrueba = reader.IsDBNull(reader.GetOrdinal("FechaInicioPrueba")) ? null : DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaInicioPrueba"))),
            FechaFinPrueba = reader.IsDBNull(reader.GetOrdinal("FechaFinPrueba")) ? null : DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaFinPrueba"))),
            FechaInicioPlan = reader.IsDBNull(reader.GetOrdinal("FechaInicioPlan")) ? null : DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaInicioPlan"))),
            FechaFinPlan = reader.IsDBNull(reader.GetOrdinal("FechaFinPlan")) ? null : DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaFinPlan"))),
            EmpresasPermitidas = reader.IsDBNull(reader.GetOrdinal("EmpresasPermitidas")) ? null : reader.GetInt32(reader.GetOrdinal("EmpresasPermitidas")),
            UsuariosPermitidos = reader.IsDBNull(reader.GetOrdinal("UsuariosPermitidos")) ? null : reader.GetInt32(reader.GetOrdinal("UsuariosPermitidos")),
            Activo = !reader.IsDBNull(reader.GetOrdinal("Activo")) && reader.GetBoolean(reader.GetOrdinal("Activo")),
            Observacion = reader.IsDBNull(reader.GetOrdinal("Observacion")) ? null : reader.GetString(reader.GetOrdinal("Observacion"))
        };
    }
}
