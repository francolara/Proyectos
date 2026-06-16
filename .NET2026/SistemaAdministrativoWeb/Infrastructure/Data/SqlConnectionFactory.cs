using Microsoft.Data.SqlClient;

namespace SistemaAdministrativoWeb.Infrastructure.Data;

public sealed class SqlConnectionFactory(IConfiguration configuration) : IDbConnectionFactory
{
    private readonly string _connectionString =
        configuration.GetConnectionString("DefaultConnection")
        ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");

    public SqlConnection CreateConnection()
    {
        return new SqlConnection(_connectionString);
    }
}
