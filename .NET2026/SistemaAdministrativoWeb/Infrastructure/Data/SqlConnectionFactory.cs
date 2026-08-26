using Microsoft.Data.SqlClient;

namespace SistemaAdministrativoWeb.Infrastructure.Data;

public sealed class SqlConnectionFactory(IConfiguration configuration) : IDbConnectionFactory
{
    private readonly string _connectionString =
        configuration["FRALSECONT_ConnectionStrings:DefaultConnection"]
        ?? throw new InvalidOperationException("Configuration 'FRALSECONT_ConnectionStrings:DefaultConnection' not found.");

    public SqlConnection CreateConnection()
    {
        return new SqlConnection(_connectionString);
    }
}
