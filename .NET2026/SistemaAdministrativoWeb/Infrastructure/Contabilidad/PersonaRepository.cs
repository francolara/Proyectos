using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class PersonaRepository(IDbConnectionFactory connectionFactory) : IPersonaRepository
{
    public async Task<PagedResult<PersonaDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, string? textoBusqueda, string? tipoPersona, bool soloClientes, bool soloProveedores, int numeroPagina, int tamanoPagina, CancellationToken cancellationToken = default)
    {
        var result = new List<PersonaDto>();
        var totalRegistros = 0;

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_ADM_ListarPersonasPorEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@TextoBusqueda", string.IsNullOrWhiteSpace(textoBusqueda) ? (object)DBNull.Value : textoBusqueda.Trim());
        command.Parameters.AddWithValue("@TipoPersona", string.IsNullOrWhiteSpace(tipoPersona) ? (object)DBNull.Value : tipoPersona.Trim().ToUpperInvariant());
        command.Parameters.AddWithValue("@SoloClientes", soloClientes);
        command.Parameters.AddWithValue("@SoloProveedores", soloProveedores);
        command.Parameters.AddWithValue("@NumeroPagina", numeroPagina);
        command.Parameters.AddWithValue("@TamanoPagina", tamanoPagina);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            totalRegistros = reader.GetInt32(reader.GetOrdinal("TotalRegistros"));
            result.Add(new PersonaDto
            {
                IdPersona = reader.GetInt32(reader.GetOrdinal("IdPersona")),
                IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
                TipoPersona = reader.GetString(reader.GetOrdinal("TipoPersona")),
                TipoDocumento = reader.GetString(reader.GetOrdinal("TipoDocumento")),
                NombreTipoDocumento = reader.GetString(reader.GetOrdinal("NombreTipoDocumento")),
                NumeroDocumento = reader.GetString(reader.GetOrdinal("NumeroDocumento")),
                NombreCompleto = reader.GetString(reader.GetOrdinal("NombreCompleto")),
                CorreoElectronico = reader.IsDBNull(reader.GetOrdinal("CorreoElectronico")) ? null : reader.GetString(reader.GetOrdinal("CorreoElectronico")),
                Telefono = reader.IsDBNull(reader.GetOrdinal("Telefono")) ? null : reader.GetString(reader.GetOrdinal("Telefono")),
                Direccion = reader.IsDBNull(reader.GetOrdinal("Direccion")) ? null : reader.GetString(reader.GetOrdinal("Direccion")),
                CodigoUbigeo = reader.IsDBNull(reader.GetOrdinal("CodigoUbigeo")) ? null : reader.GetString(reader.GetOrdinal("CodigoUbigeo")),
                Departamento = reader.IsDBNull(reader.GetOrdinal("Departamento")) ? null : reader.GetString(reader.GetOrdinal("Departamento")),
                Provincia = reader.IsDBNull(reader.GetOrdinal("Provincia")) ? null : reader.GetString(reader.GetOrdinal("Provincia")),
                Distrito = reader.IsDBNull(reader.GetOrdinal("Distrito")) ? null : reader.GetString(reader.GetOrdinal("Distrito")),
                EsCliente = reader.GetBoolean(reader.GetOrdinal("EsCliente")),
                EsProveedor = reader.GetBoolean(reader.GetOrdinal("EsProveedor")),
                Estado = reader.GetBoolean(reader.GetOrdinal("Estado"))
            });
        }

        return new PagedResult<PersonaDto>
        {
            Items = result,
            TotalRecords = totalRegistros,
            PageNumber = numeroPagina,
            PageSize = tamanoPagina
        };
    }

    public async Task<PersonaDetalleDto?> ObtenerPorIdAsync(int idEmpresa, int idPersona, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_ADM_ObtenerPersona", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@IdPersona", idPersona);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        if (!await reader.ReadAsync(cancellationToken))
        {
            return null;
        }

        return MapearDetalle(reader);
    }

    public async Task<IReadOnlyCollection<TipoDocumentoIdentidadDto>> ListarTiposDocumentoAsync(CancellationToken cancellationToken = default)
    {
        var result = new List<TipoDocumentoIdentidadDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_ADM_ListarTiposDocumentoIdentidadSunat", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new TipoDocumentoIdentidadDto
            {
                CodigoSunat = reader.GetString(reader.GetOrdinal("CodigoSunat")),
                CodigoInterno = reader.GetString(reader.GetOrdinal("CodigoInterno")),
                Nombre = reader.GetString(reader.GetOrdinal("Nombre")),
                Orden = reader.GetByte(reader.GetOrdinal("Orden"))
            });
        }

        return result;
    }

    public async Task<IReadOnlyCollection<UbigeoDepartamentoDto>> ListarDepartamentosAsync(CancellationToken cancellationToken = default)
    {
        var result = new List<UbigeoDepartamentoDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_ADM_ListarUbigeoDepartamentos", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new UbigeoDepartamentoDto
            {
                CodigoDepartamento = reader.GetString(reader.GetOrdinal("CodigoDepartamento")),
                Nombre = reader.GetString(reader.GetOrdinal("Nombre"))
            });
        }

        return result;
    }

    public async Task<IReadOnlyCollection<UbigeoProvinciaDto>> ListarProvinciasAsync(string codigoDepartamento, CancellationToken cancellationToken = default)
    {
        var result = new List<UbigeoProvinciaDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_ADM_ListarUbigeoProvincias", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@CodigoDepartamento", codigoDepartamento);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new UbigeoProvinciaDto
            {
                CodigoProvincia = reader.GetString(reader.GetOrdinal("CodigoProvincia")),
                CodigoDepartamento = reader.GetString(reader.GetOrdinal("CodigoDepartamento")),
                Nombre = reader.GetString(reader.GetOrdinal("Nombre"))
            });
        }

        return result;
    }

    public async Task<IReadOnlyCollection<UbigeoDistritoDto>> ListarDistritosAsync(string codigoProvincia, CancellationToken cancellationToken = default)
    {
        var result = new List<UbigeoDistritoDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_ADM_ListarUbigeoDistritos", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@CodigoProvincia", codigoProvincia);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new UbigeoDistritoDto
            {
                CodigoUbigeo = reader.GetString(reader.GetOrdinal("CodigoUbigeo")),
                CodigoDepartamento = reader.GetString(reader.GetOrdinal("CodigoDepartamento")),
                CodigoProvincia = reader.GetString(reader.GetOrdinal("CodigoProvincia")),
                Nombre = reader.GetString(reader.GetOrdinal("Nombre")),
                Zona = reader.IsDBNull(reader.GetOrdinal("Zona")) ? null : reader.GetString(reader.GetOrdinal("Zona"))
            });
        }

        return result;
    }

    public async Task<PersonaDetalleDto> GuardarAsync(GuardarPersonaRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_ADM_GuardarPersona", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdPersona", (object?)request.IdPersona ?? DBNull.Value);
        command.Parameters.AddWithValue("@IdEmpresa", request.IdEmpresa);
        command.Parameters.AddWithValue("@TipoPersona", request.TipoPersona);
        command.Parameters.AddWithValue("@TipoDocumento", request.TipoDocumento);
        command.Parameters.AddWithValue("@NumeroDocumento", request.NumeroDocumento);
        command.Parameters.AddWithValue("@ApellidoPaterno", (object?)request.ApellidoPaterno ?? DBNull.Value);
        command.Parameters.AddWithValue("@ApellidoMaterno", (object?)request.ApellidoMaterno ?? DBNull.Value);
        command.Parameters.AddWithValue("@Nombres", (object?)request.Nombres ?? DBNull.Value);
        command.Parameters.AddWithValue("@RazonSocial", (object?)request.RazonSocial ?? DBNull.Value);
        command.Parameters.AddWithValue("@CorreoElectronico", (object?)request.CorreoElectronico ?? DBNull.Value);
        command.Parameters.AddWithValue("@Telefono", (object?)request.Telefono ?? DBNull.Value);
        command.Parameters.AddWithValue("@Direccion", (object?)request.Direccion ?? DBNull.Value);
        command.Parameters.AddWithValue("@CodigoUbigeo", (object?)request.CodigoUbigeo ?? DBNull.Value);
        command.Parameters.AddWithValue("@EsCliente", request.EsCliente);
        command.Parameters.AddWithValue("@EsProveedor", request.EsProveedor);
        command.Parameters.AddWithValue("@Estado", request.Estado);
        command.Parameters.AddWithValue("@UsuarioRegistro", (object?)request.UsuarioRegistro ?? DBNull.Value);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        if (!await reader.ReadAsync(cancellationToken))
        {
            throw new InvalidOperationException("No se pudo obtener la persona grabada.");
        }

        return MapearDetalle(reader);
    }

    private static PersonaDetalleDto MapearDetalle(SqlDataReader reader)
    {
        return new PersonaDetalleDto
        {
            IdPersona = reader.GetInt32(reader.GetOrdinal("IdPersona")),
            IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
            TipoPersona = reader.GetString(reader.GetOrdinal("TipoPersona")),
            TipoDocumento = reader.GetString(reader.GetOrdinal("TipoDocumento")),
            NumeroDocumento = reader.GetString(reader.GetOrdinal("NumeroDocumento")),
            ApellidoPaterno = reader.IsDBNull(reader.GetOrdinal("ApellidoPaterno")) ? null : reader.GetString(reader.GetOrdinal("ApellidoPaterno")),
            ApellidoMaterno = reader.IsDBNull(reader.GetOrdinal("ApellidoMaterno")) ? null : reader.GetString(reader.GetOrdinal("ApellidoMaterno")),
            Nombres = reader.IsDBNull(reader.GetOrdinal("Nombres")) ? null : reader.GetString(reader.GetOrdinal("Nombres")),
            RazonSocial = reader.IsDBNull(reader.GetOrdinal("RazonSocial")) ? null : reader.GetString(reader.GetOrdinal("RazonSocial")),
            NombreCompleto = reader.GetString(reader.GetOrdinal("NombreCompleto")),
            CorreoElectronico = reader.IsDBNull(reader.GetOrdinal("CorreoElectronico")) ? null : reader.GetString(reader.GetOrdinal("CorreoElectronico")),
            Telefono = reader.IsDBNull(reader.GetOrdinal("Telefono")) ? null : reader.GetString(reader.GetOrdinal("Telefono")),
            Direccion = reader.IsDBNull(reader.GetOrdinal("Direccion")) ? null : reader.GetString(reader.GetOrdinal("Direccion")),
            CodigoUbigeo = reader.IsDBNull(reader.GetOrdinal("CodigoUbigeo")) ? null : reader.GetString(reader.GetOrdinal("CodigoUbigeo")),
            CodigoDepartamento = reader.IsDBNull(reader.GetOrdinal("CodigoDepartamento")) ? null : reader.GetString(reader.GetOrdinal("CodigoDepartamento")),
            CodigoProvincia = reader.IsDBNull(reader.GetOrdinal("CodigoProvincia")) ? null : reader.GetString(reader.GetOrdinal("CodigoProvincia")),
            EsCliente = reader.GetBoolean(reader.GetOrdinal("EsCliente")),
            EsProveedor = reader.GetBoolean(reader.GetOrdinal("EsProveedor")),
            Estado = reader.GetBoolean(reader.GetOrdinal("Estado"))
        };
    }
}
