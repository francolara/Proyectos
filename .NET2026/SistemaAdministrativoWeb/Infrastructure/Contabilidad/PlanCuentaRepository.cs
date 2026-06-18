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
            result.Add(new PlanCuentaDto
            {
                IdPlanCuenta = reader.GetInt32(reader.GetOrdinal("IdPlanCuenta")),
                IdPlanCuentaPadre = reader.IsDBNull(reader.GetOrdinal("IdPlanCuentaPadre")) ? null : reader.GetInt32(reader.GetOrdinal("IdPlanCuentaPadre")),
                CodigoCuenta = reader.GetString(reader.GetOrdinal("CodigoCuenta")),
                NombreCuenta = reader.GetString(reader.GetOrdinal("NombreCuenta")),
                NivelCuenta = reader.GetByte(reader.GetOrdinal("NivelCuenta")),
                NaturalezaSaldo = reader.GetString(reader.GetOrdinal("NaturalezaSaldo")),
                AceptaMovimiento = reader.GetBoolean(reader.GetOrdinal("AceptaMovimiento")),
                RequiereCentroCosto = reader.GetBoolean(reader.GetOrdinal("RequiereCentroCosto")),
                Estado = reader.GetBoolean(reader.GetOrdinal("Estado"))
            });
        }

        return result;
    }

    public async Task<PagedResult<PlanCuentaDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, string? textoBusqueda, int numeroPagina, int tamanoPagina, bool soloMovimiento = false, CancellationToken cancellationToken = default)
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
        command.Parameters.AddWithValue("@NumeroPagina", numeroPagina);
        command.Parameters.AddWithValue("@TamanoPagina", tamanoPagina);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            totalRegistros = reader.GetInt32(reader.GetOrdinal("TotalRegistros"));
            result.Add(new PlanCuentaDto
            {
                IdPlanCuenta = reader.GetInt32(reader.GetOrdinal("IdPlanCuenta")),
                IdPlanCuentaPadre = reader.IsDBNull(reader.GetOrdinal("IdPlanCuentaPadre")) ? null : reader.GetInt32(reader.GetOrdinal("IdPlanCuentaPadre")),
                CodigoCuenta = reader.GetString(reader.GetOrdinal("CodigoCuenta")),
                NombreCuenta = reader.GetString(reader.GetOrdinal("NombreCuenta")),
                NivelCuenta = reader.GetByte(reader.GetOrdinal("NivelCuenta")),
                NaturalezaSaldo = reader.GetString(reader.GetOrdinal("NaturalezaSaldo")),
                AceptaMovimiento = reader.GetBoolean(reader.GetOrdinal("AceptaMovimiento")),
                RequiereCentroCosto = reader.GetBoolean(reader.GetOrdinal("RequiereCentroCosto")),
                Estado = reader.GetBoolean(reader.GetOrdinal("Estado"))
            });
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
        command.Parameters.AddWithValue("@NaturalezaSaldo", request.NaturalezaSaldo);
        command.Parameters.AddWithValue("@AceptaMovimiento", request.AceptaMovimiento);
        command.Parameters.AddWithValue("@RequiereCentroCosto", request.RequiereCentroCosto);
        command.Parameters.AddWithValue("@Estado", request.Estado);
        command.Parameters.AddWithValue("@UsuarioRegistro", (object?)request.UsuarioRegistro ?? DBNull.Value);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        if (!await reader.ReadAsync(cancellationToken))
        {
            throw new InvalidOperationException("No se pudo obtener la cuenta grabada.");
        }

        return new PlanCuentaDto
        {
            IdPlanCuenta = reader.GetInt32(reader.GetOrdinal("IdPlanCuenta")),
            IdPlanCuentaPadre = reader.IsDBNull(reader.GetOrdinal("IdPlanCuentaPadre")) ? null : reader.GetInt32(reader.GetOrdinal("IdPlanCuentaPadre")),
            CodigoCuenta = reader.GetString(reader.GetOrdinal("CodigoCuenta")),
            NombreCuenta = reader.GetString(reader.GetOrdinal("NombreCuenta")),
            NivelCuenta = reader.GetByte(reader.GetOrdinal("NivelCuenta")),
            NaturalezaSaldo = reader.GetString(reader.GetOrdinal("NaturalezaSaldo")),
            AceptaMovimiento = reader.GetBoolean(reader.GetOrdinal("AceptaMovimiento")),
            RequiereCentroCosto = reader.GetBoolean(reader.GetOrdinal("RequiereCentroCosto")),
            Estado = reader.GetBoolean(reader.GetOrdinal("Estado"))
        };
    }
}
