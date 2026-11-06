using Microsoft.Data.SqlClient;

namespace GestionServicios.Infrastructure;

// Crear una conexión no la abre ni ejecuta scripts de despliegue.
public sealed class SqlConnectionFactory(string connectionString)
{
    public SqlConnection CreateConnection()
    {
        if (string.IsNullOrWhiteSpace(connectionString))
            throw new InvalidOperationException("Configure la conexión de HYT mediante un almacén de secretos.");
        return new SqlConnection(connectionString);
    }
}
