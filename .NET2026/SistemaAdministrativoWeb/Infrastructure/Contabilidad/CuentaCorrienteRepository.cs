using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class CuentaCorrienteRepository(IDbConnectionFactory connectionFactory) : ICuentaCorrienteRepository
{
    public async Task<PagedResult<BancoConfiguracionEmpresaDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, string? textoBusqueda, int numeroPagina, int tamanoPagina, bool soloActivos = false, CancellationToken cancellationToken = default)
    {
        var result = new List<BancoConfiguracionEmpresaDto>();
        var totalRegistros = 0;

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ListarBancosConfiguracionEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@SoloActivos", soloActivos);
        command.Parameters.AddWithValue("@TextoBusqueda", string.IsNullOrWhiteSpace(textoBusqueda) ? (object)DBNull.Value : textoBusqueda.Trim());
        command.Parameters.AddWithValue("@NumeroPagina", numeroPagina);
        command.Parameters.AddWithValue("@TamanoPagina", tamanoPagina);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            totalRegistros = reader.GetInt32(reader.GetOrdinal("TotalRegistros"));
            result.Add(Map(reader));
        }

        return new PagedResult<BancoConfiguracionEmpresaDto>
        {
            Items = result,
            TotalRecords = totalRegistros,
            PageNumber = numeroPagina,
            PageSize = tamanoPagina
        };
    }

    public async Task<BancoConfiguracionEmpresaDto?> ObtenerPorIdAsync(int idEmpresa, int idBancoConfiguracionEmpresa, CancellationToken cancellationToken = default)
    {
        var result = await ListarPaginadoPorEmpresaAsync(idEmpresa, null, 1, int.MaxValue, false, cancellationToken);
        return result.Items.FirstOrDefault(x => x.IdBancoConfiguracionEmpresa == idBancoConfiguracionEmpresa);
    }

    public async Task<IReadOnlyCollection<BancoConfiguracionEmpresaDto>> ListarPorEmpresaAsync(int idEmpresa, bool soloActivos = false, CancellationToken cancellationToken = default)
    {
        var result = await ListarPaginadoPorEmpresaAsync(idEmpresa, null, 1, int.MaxValue, soloActivos, cancellationToken);
        return result.Items
            .OrderBy(x => x.NombreBanco)
            .ThenBy(x => x.NroCuentaCorriente)
            .ToList();
    }

    public async Task<BancoConfiguracionEmpresaDto> GuardarAsync(GuardarBancoConfiguracionEmpresaRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_GuardarBancoConfiguracionEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdBancoConfiguracionEmpresa", (object?)request.IdBancoConfiguracionEmpresa ?? DBNull.Value);
        command.Parameters.AddWithValue("@IdEmpresa", request.IdEmpresa);
        command.Parameters.AddWithValue("@IdBanco", request.IdBanco);
        command.Parameters.AddWithValue("@NroCuentaCorriente", request.NroCuentaCorriente);
        command.Parameters.AddWithValue("@Titular", request.Titular);
        command.Parameters.AddWithValue("@IdMoneda", request.IdMoneda);
        command.Parameters.AddWithValue("@IdPlanCuenta", request.IdPlanCuenta);
        command.Parameters.AddWithValue("@PeriodoSaldoInicial", request.PeriodoSaldoInicial);
        command.Parameters.AddWithValue("@SaldoInicialDebe", request.SaldoInicialDebe);
        command.Parameters.AddWithValue("@SaldoInicialHaber", request.SaldoInicialHaber);
        command.Parameters.AddWithValue("@Activo", request.Activo);
        command.Parameters.AddWithValue("@UsuarioRegistro", string.IsNullOrWhiteSpace(request.UsuarioRegistro) ? (object)DBNull.Value : request.UsuarioRegistro.Trim());

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        if (!await reader.ReadAsync(cancellationToken))
        {
            throw new InvalidOperationException("No se pudo obtener la cuenta corriente grabada.");
        }

        return Map(reader);
    }

    private static BancoConfiguracionEmpresaDto Map(SqlDataReader reader)
    {
        return new BancoConfiguracionEmpresaDto
        {
            IdBancoConfiguracionEmpresa = reader.GetInt32(reader.GetOrdinal("IdBancoConfiguracionEmpresa")),
            IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
            IdBanco = reader.GetInt32(reader.GetOrdinal("IdBanco")),
            CodigoBanco = reader.GetString(reader.GetOrdinal("CodigoBanco")),
            NombreBanco = reader.GetString(reader.GetOrdinal("NombreBanco")),
            NroCuentaCorriente = reader.GetString(reader.GetOrdinal("NroCuentaCorriente")),
            Titular = reader.IsDBNull(reader.GetOrdinal("Titular"))
                ? string.Empty
                : reader.GetString(reader.GetOrdinal("Titular")),
            IdMoneda = reader.IsDBNull(reader.GetOrdinal("IdMoneda"))
                ? null
                : reader.GetInt32(reader.GetOrdinal("IdMoneda")),
            CodigoMoneda = reader.IsDBNull(reader.GetOrdinal("CodigoMoneda"))
                ? string.Empty
                : reader.GetString(reader.GetOrdinal("CodigoMoneda")),
            NombreMoneda = reader.IsDBNull(reader.GetOrdinal("NombreMoneda"))
                ? string.Empty
                : reader.GetString(reader.GetOrdinal("NombreMoneda")),
            IdPlanCuenta = reader.GetInt32(reader.GetOrdinal("IdPlanCuenta")),
            CodigoCuenta = reader.GetString(reader.GetOrdinal("CodigoCuenta")),
            NombreCuenta = reader.GetString(reader.GetOrdinal("NombreCuenta")),
            PeriodoSaldoInicial = reader.IsDBNull(reader.GetOrdinal("PeriodoSaldoInicial"))
                ? string.Empty
                : reader.GetString(reader.GetOrdinal("PeriodoSaldoInicial")),
            SaldoInicialDebe = reader.IsDBNull(reader.GetOrdinal("SaldoInicialDebe"))
                ? 0m
                : reader.GetDecimal(reader.GetOrdinal("SaldoInicialDebe")),
            SaldoInicialHaber = reader.IsDBNull(reader.GetOrdinal("SaldoInicialHaber"))
                ? 0m
                : reader.GetDecimal(reader.GetOrdinal("SaldoInicialHaber")),
            Activo = reader.GetBoolean(reader.GetOrdinal("Activo")),
            FechaRegistro = reader.GetDateTime(reader.GetOrdinal("FechaRegistro")),
            UsuarioRegistro = reader.IsDBNull(reader.GetOrdinal("UsuarioRegistro"))
                ? string.Empty
                : reader.GetString(reader.GetOrdinal("UsuarioRegistro"))
        };
    }
}
