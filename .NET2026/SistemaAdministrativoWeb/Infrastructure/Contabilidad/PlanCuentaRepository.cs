using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class PlanCuentaRepository(IDbConnectionFactory connectionFactory) : IPlanCuentaRepository
{
    public async Task<IReadOnlyCollection<PlanCuentaDto>> ListarPorEmpresaAsync(int idEmpresa, bool soloMovimiento = false, CancellationToken cancellationToken = default)
    {
        var result = new List<PlanCuentaDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ListarPlanCuentaPorEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@SoloMovimiento", soloMovimiento);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(MapPlanCuenta(reader));
        }

        return result;
    }

    public async Task<PagedResult<PlanCuentaDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, string? textoBusqueda, byte? nivelCuenta, int numeroPagina, int tamanoPagina, bool soloMovimiento = false, bool soloUltimoNivel = false, CancellationToken cancellationToken = default)
    {
        var result = new List<PlanCuentaDto>();
        var totalRegistros = 0;

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ListarPlanCuentaPorEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@SoloMovimiento", soloMovimiento);
        command.Parameters.AddWithValue("@TextoBusqueda", string.IsNullOrWhiteSpace(textoBusqueda) ? (object)DBNull.Value : textoBusqueda.Trim());
        command.Parameters.AddWithValue("@NivelCuenta", (object?)nivelCuenta ?? DBNull.Value);
        command.Parameters.AddWithValue("@NumeroPagina", numeroPagina);
        command.Parameters.AddWithValue("@TamanoPagina", tamanoPagina);
        command.Parameters.AddWithValue("@SoloUltimoNivel", soloUltimoNivel);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            totalRegistros = reader.GetInt32(reader.GetOrdinal("TotalRegistros"));
            result.Add(MapPlanCuenta(reader));
        }

        return new PagedResult<PlanCuentaDto>
        {
            Items = result,
            TotalRecords = totalRegistros,
            PageNumber = numeroPagina,
            PageSize = tamanoPagina
        };
    }

    public async Task<PlanCuentaDto> GuardarAsync(GuardarPlanCuentaRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_GuardarPlanCuentaPorEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdPlanCuenta", (object?)request.IdPlanCuenta ?? DBNull.Value);
        command.Parameters.AddWithValue("@IdEmpresa", request.IdEmpresa);
        command.Parameters.AddWithValue("@IdPlanCuentaPadre", (object?)request.IdPlanCuentaPadre ?? DBNull.Value);
        command.Parameters.AddWithValue("@CodigoCuenta", request.CodigoCuenta);
        command.Parameters.AddWithValue("@NombreCuenta", request.NombreCuenta);
        command.Parameters.AddWithValue("@ColBalance", request.ColBalance);
        command.Parameters.AddWithValue("@IdMoneda", request.IdMoneda);
        command.Parameters.AddWithValue("@TipoCambio", request.TipoCambio);
        command.Parameters.AddWithValue("@AceptaMovimiento", request.AceptaMovimiento);
        command.Parameters.AddWithValue("@GeneraDiferenciaPorAnalisis", request.GeneraDiferenciaPorAnalisis);
        command.Parameters.AddWithValue("@RequiereCentroCosto", request.RequiereCentroCosto);
        command.Parameters.AddWithValue("@Estado", request.Estado);
        command.Parameters.AddWithValue("@UsuarioRegistro", (object?)request.UsuarioRegistro ?? DBNull.Value);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        if (!await reader.ReadAsync(cancellationToken))
        {
            throw new InvalidOperationException("No se pudo obtener la cuenta grabada.");
        }

        return MapPlanCuenta(reader);
    }

    public async Task CargarDefaultAsync(int idEmpresa, string? usuarioRegistro, CancellationToken cancellationToken = default, int? idEmpresaBase = null)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_CargarPlanCuentaDefaultEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@IdEmpresaBase", (object?)idEmpresaBase ?? DBNull.Value);
        command.Parameters.AddWithValue("@UsuarioRegistro", (object?)usuarioRegistro ?? DBNull.Value);

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    private static PlanCuentaDto MapPlanCuenta(SqlDataReader reader)
    {
        return new PlanCuentaDto
        {
            IdPlanCuenta = reader.GetInt32(reader.GetOrdinal("IdPlanCuenta")),
            IdPlanCuentaPadre = reader.IsDBNull(reader.GetOrdinal("IdPlanCuentaPadre")) ? null : reader.GetInt32(reader.GetOrdinal("IdPlanCuentaPadre")),
            CodigoCuenta = reader.GetString(reader.GetOrdinal("CodigoCuenta")),
            NombreCuenta = reader.GetString(reader.GetOrdinal("NombreCuenta")),
            NivelCuenta = reader.GetByte(reader.GetOrdinal("NivelCuenta")),
            ColBalance = reader.GetString(reader.GetOrdinal("ColBalance")),
            IdMoneda = reader.GetString(reader.GetOrdinal("IdMoneda")),
            TipoCambio = reader.GetString(reader.GetOrdinal("TipoCambio")),
            AceptaMovimiento = reader.GetBoolean(reader.GetOrdinal("AceptaMovimiento")),
            GeneraDiferenciaPorAnalisis = reader.GetBoolean(reader.GetOrdinal("GeneraDiferenciaPorAnalisis")),
            EsUltimoNivel = reader.GetBoolean(reader.GetOrdinal("EsUltimoNivel")),
            RequiereCentroCosto = reader.GetBoolean(reader.GetOrdinal("RequiereCentroCosto")),
            Estado = reader.GetBoolean(reader.GetOrdinal("Estado"))
        };
    }
}
