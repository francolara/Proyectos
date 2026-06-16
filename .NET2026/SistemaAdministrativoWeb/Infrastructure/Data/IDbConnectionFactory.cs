using Microsoft.Data.SqlClient;

namespace SistemaAdministrativoWeb.Infrastructure.Data;

public interface IDbConnectionFactory
{
    SqlConnection CreateConnection();
}
