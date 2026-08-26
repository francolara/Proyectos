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

    public async Task<RegistroEmpresaCuentaAdministradoraResult> RegistrarEmpresaCuentaAsync(RegistroEmpresaCuentaAdministradoraRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_SEG_RegistrarEmpresaCuentaAdministradora", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdCuentaAdministradora", request.IdCuentaAdministradora);
        command.Parameters.AddWithValue("@AspNetUserId", request.AspNetUserId);
        command.Parameters.AddWithValue("@CodigoEmpresa", request.CodigoEmpresa);
        command.Parameters.AddWithValue("@RazonSocial", request.RazonSocial);
        command.Parameters.AddWithValue("@NombreComercial", (object?)request.NombreComercial ?? DBNull.Value);
        command.Parameters.AddWithValue("@Ruc", request.Ruc);
        command.Parameters.AddWithValue("@EsEmpresaPredeterminada", request.EsEmpresaPredeterminada);
        command.Parameters.AddWithValue("@UsuarioRegistro", (object?)request.UsuarioRegistro ?? DBNull.Value);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        if (!await reader.ReadAsync(cancellationToken))
        {
            throw new InvalidOperationException("No se pudo obtener la empresa registrada.");
        }

        return new RegistroEmpresaCuentaAdministradoraResult
        {
            IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa"))
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
            result.Add(MapearCuentaSuscripcionResumen(reader));
        }

        return result;
    }

    public async Task<CuentaSuscripcionPaginaDto> ListarCuentasSuscripcionPaginadasAsync(
        string? textoBusqueda,
        string estadoFiltro,
        int pagina,
        int tamanoPagina,
        CancellationToken cancellationToken = default)
    {
        var cuentas = new List<CuentaSuscripcionResumenDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_SEG_ListarCuentasAdministradorasSuscripcionPaginado", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.Add("@TextoBusqueda", SqlDbType.NVarChar, 200).Value =
            string.IsNullOrWhiteSpace(textoBusqueda) ? DBNull.Value : textoBusqueda.Trim();
        command.Parameters.Add("@EstadoFiltro", SqlDbType.NVarChar, 20).Value = estadoFiltro;
        command.Parameters.Add("@NumeroPagina", SqlDbType.Int).Value = pagina;
        command.Parameters.Add("@TamanoPagina", SqlDbType.Int).Value = tamanoPagina;

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        var paginaActual = 1;
        var tamanoPaginaEfectivo = tamanoPagina;
        var totalFiltrado = 0;
        var totalPaginas = 1;
        var totalCuentas = 0;
        var cuentasActivas = 0;
        var cuentasEnPrueba = 0;
        var cuentasSuspendidasOBaja = 0;
        var cobrosRegistrados = 0;
        var cobrosPendientesAplicacion = 0;
        var montoCobradoMes = 0m;

        if (await reader.ReadAsync(cancellationToken))
        {
            paginaActual = reader.GetInt32(reader.GetOrdinal("PaginaActual"));
            tamanoPaginaEfectivo = reader.GetInt32(reader.GetOrdinal("TamanoPagina"));
            totalFiltrado = reader.GetInt32(reader.GetOrdinal("TotalFiltrado"));
            totalPaginas = reader.GetInt32(reader.GetOrdinal("TotalPaginas"));
            totalCuentas = reader.GetInt32(reader.GetOrdinal("TotalCuentas"));
            cuentasActivas = reader.GetInt32(reader.GetOrdinal("CuentasActivas"));
            cuentasEnPrueba = reader.GetInt32(reader.GetOrdinal("CuentasEnPrueba"));
            cuentasSuspendidasOBaja = reader.GetInt32(reader.GetOrdinal("CuentasSuspendidasOBaja"));
            cobrosRegistrados = reader.GetInt32(reader.GetOrdinal("CobrosRegistrados"));
            cobrosPendientesAplicacion = reader.GetInt32(reader.GetOrdinal("CobrosPendientesAplicacion"));
            montoCobradoMes = reader.GetDecimal(reader.GetOrdinal("MontoCobradoMes"));
        }

        if (await reader.NextResultAsync(cancellationToken))
        {
            while (await reader.ReadAsync(cancellationToken))
            {
                cuentas.Add(MapearCuentaSuscripcionResumen(reader));
            }
        }

        return new CuentaSuscripcionPaginaDto
        {
            PaginaActual = paginaActual,
            TamanoPagina = tamanoPaginaEfectivo,
            TotalFiltrado = totalFiltrado,
            TotalPaginas = totalPaginas,
            TotalCuentas = totalCuentas,
            CuentasActivas = cuentasActivas,
            CuentasEnPrueba = cuentasEnPrueba,
            CuentasSuspendidasOBaja = cuentasSuspendidasOBaja,
            CobrosRegistrados = cobrosRegistrados,
            CobrosPendientesAplicacion = cobrosPendientesAplicacion,
            MontoCobradoMes = montoCobradoMes,
            Cuentas = cuentas
        };
    }

    public async Task SincronizarVencimientoSuscripcionCuentaAsync(
        int idCuentaAdministradora,
        string? usuarioRegistro,
        CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_SEG_SincronizarVencimientoSuscripcionCuentaAdministradora", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.Add("@IdCuentaAdministradora", SqlDbType.Int).Value = idCuentaAdministradora;
        command.Parameters.Add("@UsuarioRegistro", SqlDbType.NVarChar, 450).Value =
            string.IsNullOrWhiteSpace(usuarioRegistro) ? DBNull.Value : usuarioRegistro.Trim();

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
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
        command.Parameters.AddWithValue("@TipoCobro", (object?)request.TipoCobro ?? DBNull.Value);
        command.Parameters.AddWithValue("@DiasGracia", request.DiasGracia <= 0 ? 5 : request.DiasGracia);
        command.Parameters.AddWithValue("@EmpresasPermitidas", (object?)request.EmpresasPermitidas ?? DBNull.Value);
        command.Parameters.AddWithValue("@UsuariosPermitidos", (object?)request.UsuariosPermitidos ?? DBNull.Value);
        command.Parameters.AddWithValue("@Activo", request.Activo);
        command.Parameters.AddWithValue("@EstadoCuenta", request.EstadoCuenta);
        command.Parameters.AddWithValue("@Observacion", (object?)request.Observacion ?? DBNull.Value);
        command.Parameters.AddWithValue("@UsuarioRegistro", (object?)request.UsuarioRegistro ?? DBNull.Value);

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task ActivarContratoCuentaAsync(ActivarContratoCuentaRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_SEG_ActivarContratoCuentaAdministradora", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdCuentaAdministradora", request.IdCuentaAdministradora);
        command.Parameters.AddWithValue("@TipoPlan", request.TipoPlan);
        command.Parameters.AddWithValue("@TipoCobro", request.TipoCobro);
        command.Parameters.AddWithValue("@FechaInicioPlan", request.FechaInicioPlan.ToDateTime(TimeOnly.MinValue));
        command.Parameters.AddWithValue("@FechaFinPlan", request.FechaFinPlan.ToDateTime(TimeOnly.MinValue));
        command.Parameters.AddWithValue("@DiasGracia", request.DiasGracia <= 0 ? 5 : request.DiasGracia);
        command.Parameters.AddWithValue("@Observacion", (object?)request.Observacion ?? DBNull.Value);
        command.Parameters.AddWithValue("@UsuarioRegistro", (object?)request.UsuarioRegistro ?? DBNull.Value);

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task RegistrarPagoSuscripcionCuentaAsync(RegistrarPagoSuscripcionCuentaRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_SEG_RegistrarPagoSuscripcionCuentaAdministradora", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdCuentaAdministradora", request.IdCuentaAdministradora);
        command.Parameters.AddWithValue("@TipoPago", request.TipoPago);
        command.Parameters.AddWithValue("@EstadoPago", request.EstadoPago);
        command.Parameters.AddWithValue("@Monto", request.Monto);
        command.Parameters["@Monto"].Precision = 12;
        command.Parameters["@Monto"].Scale = 2;
        command.Parameters.AddWithValue("@Moneda", request.Moneda);
        command.Parameters.AddWithValue("@FechaPago", request.FechaPago);
        command.Parameters.AddWithValue("@FechaVencimiento", request.FechaVencimiento.HasValue ? request.FechaVencimiento.Value.ToDateTime(TimeOnly.MinValue) : DBNull.Value);
        command.Parameters.AddWithValue("@OperacionNumero", (object?)request.OperacionNumero ?? DBNull.Value);
        command.Parameters.AddWithValue("@EntidadFinanciera", (object?)request.EntidadFinanciera ?? DBNull.Value);
        command.Parameters.AddWithValue("@ReferenciaExterna", (object?)request.ReferenciaExterna ?? DBNull.Value);
        command.Parameters.AddWithValue("@ProveedorPasarela", (object?)request.ProveedorPasarela ?? DBNull.Value);
        command.Parameters.AddWithValue("@TransaccionPasarelaId", (object?)request.TransaccionPasarelaId ?? DBNull.Value);
        command.Parameters.AddWithValue("@PagoPasarelaId", (object?)request.PagoPasarelaId ?? DBNull.Value);
        command.Parameters.AddWithValue("@EstadoPasarela", (object?)request.EstadoPasarela ?? DBNull.Value);
        command.Parameters.AddWithValue("@PayloadPasarela", (object?)request.PayloadPasarela ?? DBNull.Value);
        command.Parameters.AddWithValue("@Observacion", (object?)request.Observacion ?? DBNull.Value);
        command.Parameters.AddWithValue("@AccionAplicacion", (object?)request.AccionAplicacion ?? DBNull.Value);
        command.Parameters.AddWithValue("@AplicarAlConfirmar", request.AplicarAlConfirmar);
        command.Parameters.AddWithValue("@TipoCobroObjetivo", (object?)request.TipoCobroObjetivo ?? DBNull.Value);
        command.Parameters.AddWithValue("@FechaInicioPlanObjetivo", request.FechaInicioPlanObjetivo.HasValue ? request.FechaInicioPlanObjetivo.Value.ToDateTime(TimeOnly.MinValue) : DBNull.Value);
        command.Parameters.AddWithValue("@DiasGraciaObjetivo", (object?)request.DiasGraciaObjetivo ?? DBNull.Value);
        command.Parameters.AddWithValue("@UsuarioRegistro", (object?)request.UsuarioRegistro ?? DBNull.Value);

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task ConfirmarPagoSuscripcionCuentaAsync(int idCuentaAdministradora, int idCuentaAdministradoraSuscripcionPago, string? usuarioRegistro, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_SEG_ConfirmarPagoSuscripcionCuentaAdministradora", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdCuentaAdministradora", idCuentaAdministradora);
        command.Parameters.AddWithValue("@IdCuentaAdministradoraSuscripcionPago", idCuentaAdministradoraSuscripcionPago);
        command.Parameters.AddWithValue("@UsuarioRegistro", (object?)usuarioRegistro ?? DBNull.Value);

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task<IReadOnlyCollection<CuentaSuscripcionMovimientoDto>> ListarMovimientosSuscripcionCuentaAsync(int idCuentaAdministradora, int top = 20, CancellationToken cancellationToken = default)
    {
        var result = new List<CuentaSuscripcionMovimientoDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_SEG_ListarMovimientosSuscripcionCuentaAdministradora", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdCuentaAdministradora", idCuentaAdministradora);
        command.Parameters.AddWithValue("@Top", top <= 0 ? 20 : top);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new CuentaSuscripcionMovimientoDto
            {
                IdCuentaAdministradoraSuscripcionMovimiento = reader.GetInt32(reader.GetOrdinal("IdCuentaAdministradoraSuscripcionMovimiento")),
                TipoMovimiento = reader.GetString(reader.GetOrdinal("TipoMovimiento")),
                TipoPlanAnterior = reader.IsDBNull(reader.GetOrdinal("TipoPlanAnterior")) ? null : reader.GetString(reader.GetOrdinal("TipoPlanAnterior")),
                TipoPlanNuevo = reader.IsDBNull(reader.GetOrdinal("TipoPlanNuevo")) ? null : reader.GetString(reader.GetOrdinal("TipoPlanNuevo")),
                EstadoSuscripcionAnterior = reader.IsDBNull(reader.GetOrdinal("EstadoSuscripcionAnterior")) ? null : reader.GetString(reader.GetOrdinal("EstadoSuscripcionAnterior")),
                EstadoSuscripcionNuevo = reader.GetString(reader.GetOrdinal("EstadoSuscripcionNuevo")),
                EsPruebaAnterior = reader.IsDBNull(reader.GetOrdinal("EsPruebaAnterior")) ? null : reader.GetBoolean(reader.GetOrdinal("EsPruebaAnterior")),
                EsPruebaNuevo = reader.GetBoolean(reader.GetOrdinal("EsPruebaNuevo")),
                TipoCobroAnterior = reader.IsDBNull(reader.GetOrdinal("TipoCobroAnterior")) ? null : reader.GetString(reader.GetOrdinal("TipoCobroAnterior")),
                TipoCobroNuevo = reader.IsDBNull(reader.GetOrdinal("TipoCobroNuevo")) ? null : reader.GetString(reader.GetOrdinal("TipoCobroNuevo")),
                FechaInicioReferencia = reader.IsDBNull(reader.GetOrdinal("FechaInicioReferencia")) ? null : DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaInicioReferencia"))),
                FechaFinReferencia = reader.IsDBNull(reader.GetOrdinal("FechaFinReferencia")) ? null : DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaFinReferencia"))),
                DiasGracia = reader.IsDBNull(reader.GetOrdinal("DiasGracia")) ? 0 : reader.GetInt32(reader.GetOrdinal("DiasGracia")),
                DiasExtra = reader.IsDBNull(reader.GetOrdinal("DiasExtra")) ? 0 : reader.GetInt32(reader.GetOrdinal("DiasExtra")),
                EmpresasPermitidasAnterior = reader.IsDBNull(reader.GetOrdinal("EmpresasPermitidasAnterior")) ? null : reader.GetInt32(reader.GetOrdinal("EmpresasPermitidasAnterior")),
                EmpresasPermitidasNuevo = reader.IsDBNull(reader.GetOrdinal("EmpresasPermitidasNuevo")) ? null : reader.GetInt32(reader.GetOrdinal("EmpresasPermitidasNuevo")),
                UsuariosPermitidosAnterior = reader.IsDBNull(reader.GetOrdinal("UsuariosPermitidosAnterior")) ? null : reader.GetInt32(reader.GetOrdinal("UsuariosPermitidosAnterior")),
                UsuariosPermitidosNuevo = reader.IsDBNull(reader.GetOrdinal("UsuariosPermitidosNuevo")) ? null : reader.GetInt32(reader.GetOrdinal("UsuariosPermitidosNuevo")),
                Observacion = reader.IsDBNull(reader.GetOrdinal("Observacion")) ? null : reader.GetString(reader.GetOrdinal("Observacion")),
                FechaRegistro = reader.GetDateTime(reader.GetOrdinal("FechaRegistro")),
                UsuarioRegistro = reader.IsDBNull(reader.GetOrdinal("UsuarioRegistro")) ? null : reader.GetString(reader.GetOrdinal("UsuarioRegistro"))
            });
        }

        return result;
    }

    public async Task<IReadOnlyCollection<CuentaSuscripcionPagoDto>> ListarPagosSuscripcionCuentaAsync(int idCuentaAdministradora, int top = 20, CancellationToken cancellationToken = default)
    {
        var result = new List<CuentaSuscripcionPagoDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_SEG_ListarPagosSuscripcionCuentaAdministradora", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdCuentaAdministradora", idCuentaAdministradora);
        command.Parameters.AddWithValue("@Top", top <= 0 ? 20 : top);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new CuentaSuscripcionPagoDto
            {
                IdCuentaAdministradoraSuscripcionPago = reader.GetInt32(reader.GetOrdinal("IdCuentaAdministradoraSuscripcionPago")),
                TipoPago = reader.GetString(reader.GetOrdinal("TipoPago")),
                EstadoPago = reader.GetString(reader.GetOrdinal("EstadoPago")),
                Monto = reader.GetDecimal(reader.GetOrdinal("Monto")),
                Moneda = reader.GetString(reader.GetOrdinal("Moneda")),
                FechaPago = reader.GetDateTime(reader.GetOrdinal("FechaPago")),
                FechaVencimiento = reader.IsDBNull(reader.GetOrdinal("FechaVencimiento")) ? null : DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaVencimiento"))),
                OperacionNumero = reader.IsDBNull(reader.GetOrdinal("OperacionNumero")) ? null : reader.GetString(reader.GetOrdinal("OperacionNumero")),
                EntidadFinanciera = reader.IsDBNull(reader.GetOrdinal("EntidadFinanciera")) ? null : reader.GetString(reader.GetOrdinal("EntidadFinanciera")),
                ReferenciaExterna = reader.IsDBNull(reader.GetOrdinal("ReferenciaExterna")) ? null : reader.GetString(reader.GetOrdinal("ReferenciaExterna")),
                ProveedorPasarela = reader.IsDBNull(reader.GetOrdinal("ProveedorPasarela")) ? null : reader.GetString(reader.GetOrdinal("ProveedorPasarela")),
                TransaccionPasarelaId = reader.IsDBNull(reader.GetOrdinal("TransaccionPasarelaId")) ? null : reader.GetString(reader.GetOrdinal("TransaccionPasarelaId")),
                PagoPasarelaId = reader.IsDBNull(reader.GetOrdinal("PagoPasarelaId")) ? null : reader.GetString(reader.GetOrdinal("PagoPasarelaId")),
                EstadoPasarela = reader.IsDBNull(reader.GetOrdinal("EstadoPasarela")) ? null : reader.GetString(reader.GetOrdinal("EstadoPasarela")),
                AccionAplicacion = reader.IsDBNull(reader.GetOrdinal("AccionAplicacion")) ? null : reader.GetString(reader.GetOrdinal("AccionAplicacion")),
                AplicarAlConfirmar = !reader.IsDBNull(reader.GetOrdinal("AplicarAlConfirmar")) && reader.GetBoolean(reader.GetOrdinal("AplicarAlConfirmar")),
                AplicadoSuscripcion = !reader.IsDBNull(reader.GetOrdinal("AplicadoSuscripcion")) && reader.GetBoolean(reader.GetOrdinal("AplicadoSuscripcion")),
                FechaAplicacion = reader.IsDBNull(reader.GetOrdinal("FechaAplicacion")) ? null : reader.GetDateTime(reader.GetOrdinal("FechaAplicacion")),
                UsuarioAplicacion = reader.IsDBNull(reader.GetOrdinal("UsuarioAplicacion")) ? null : reader.GetString(reader.GetOrdinal("UsuarioAplicacion")),
                TipoCobroObjetivo = reader.IsDBNull(reader.GetOrdinal("TipoCobroObjetivo")) ? null : reader.GetString(reader.GetOrdinal("TipoCobroObjetivo")),
                FechaInicioPlanObjetivo = reader.IsDBNull(reader.GetOrdinal("FechaInicioPlanObjetivo")) ? null : DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaInicioPlanObjetivo"))),
                DiasGraciaObjetivo = reader.IsDBNull(reader.GetOrdinal("DiasGraciaObjetivo")) ? null : reader.GetInt32(reader.GetOrdinal("DiasGraciaObjetivo")),
                Observacion = reader.IsDBNull(reader.GetOrdinal("Observacion")) ? null : reader.GetString(reader.GetOrdinal("Observacion")),
                FechaRegistro = reader.GetDateTime(reader.GetOrdinal("FechaRegistro")),
                UsuarioRegistro = reader.IsDBNull(reader.GetOrdinal("UsuarioRegistro")) ? null : reader.GetString(reader.GetOrdinal("UsuarioRegistro"))
            });
        }

        return result;
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
            TipoCobro = reader.IsDBNull(reader.GetOrdinal("TipoCobro")) ? null : reader.GetString(reader.GetOrdinal("TipoCobro")),
            DiasGracia = reader.IsDBNull(reader.GetOrdinal("DiasGracia")) ? 5 : reader.GetInt32(reader.GetOrdinal("DiasGracia")),
            FechaFinGracia = reader.IsDBNull(reader.GetOrdinal("FechaFinGracia")) ? null : DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaFinGracia"))),
            EmpresasPermitidas = reader.IsDBNull(reader.GetOrdinal("EmpresasPermitidas")) ? null : reader.GetInt32(reader.GetOrdinal("EmpresasPermitidas")),
            UsuariosPermitidos = reader.IsDBNull(reader.GetOrdinal("UsuariosPermitidos")) ? null : reader.GetInt32(reader.GetOrdinal("UsuariosPermitidos")),
            Activo = !reader.IsDBNull(reader.GetOrdinal("Activo")) && reader.GetBoolean(reader.GetOrdinal("Activo")),
            Observacion = reader.IsDBNull(reader.GetOrdinal("Observacion")) ? null : reader.GetString(reader.GetOrdinal("Observacion"))
        };
    }

    public async Task<ContextoLoginUsuarioDto?> ObtenerContextoLoginUsuarioAsync(string aspNetUserId, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_SEG_ObtenerContextoLoginUsuario", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@AspNetUserId", aspNetUserId);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        if (!await reader.ReadAsync(cancellationToken))
        {
            return null;
        }

        return new ContextoLoginUsuarioDto
        {
            TieneAcceso = reader.GetBoolean(reader.GetOrdinal("TieneAcceso")),
            EsSuperAdmin = reader.GetBoolean(reader.GetOrdinal("EsSuperAdmin")),
            IdCuentaAdministradora = GetNullableInt(reader, "IdCuentaAdministradora"),
            CodigoCuenta = GetNullableString(reader, "CodigoCuenta"),
            NombreCuenta = GetNullableString(reader, "NombreCuenta"),
            CorreoPrincipal = GetNullableString(reader, "CorreoPrincipal"),
            TelefonoPrincipal = GetNullableString(reader, "TelefonoPrincipal"),
            EstadoCuenta = GetNullableBool(reader, "EstadoCuenta"),
            RolCuenta = GetNullableString(reader, "RolCuenta"),
            CantidadEmpresasAsignadas = GetNullableInt(reader, "CantidadEmpresasAsignadas") ?? 0,
            IdEmpresaPredeterminada = GetNullableInt(reader, "IdEmpresaPredeterminada"),
            RazonSocialEmpresaPredeterminada = GetNullableString(reader, "RazonSocialEmpresaPredeterminada"),
            DebeSeleccionarEmpresa = reader.GetBoolean(reader.GetOrdinal("DebeSeleccionarEmpresa")),
            SoloModulosCuenta = reader.GetBoolean(reader.GetOrdinal("SoloModulosCuenta")),
            IdCuentaAdministradoraSuscripcion = GetNullableInt(reader, "IdCuentaAdministradoraSuscripcion"),
            TipoPlan = GetNullableString(reader, "TipoPlan"),
            EstadoSuscripcion = GetNullableString(reader, "EstadoSuscripcion"),
            EsPrueba = GetNullableBool(reader, "EsPrueba"),
            FechaInicioPrueba = GetNullableDateOnly(reader, "FechaInicioPrueba"),
            FechaFinPrueba = GetNullableDateOnly(reader, "FechaFinPrueba"),
            FechaInicioPlan = GetNullableDateOnly(reader, "FechaInicioPlan"),
            FechaFinPlan = GetNullableDateOnly(reader, "FechaFinPlan"),
            DiasGracia = GetNullableInt(reader, "DiasGracia"),
            FechaFinGracia = GetNullableDateOnly(reader, "FechaFinGracia"),
            EmpresasPermitidas = GetNullableInt(reader, "EmpresasPermitidas"),
            UsuariosPermitidos = GetNullableInt(reader, "UsuariosPermitidos"),
            ActivoSuscripcion = GetNullableBool(reader, "ActivoSuscripcion"),
            ObservacionSuscripcion = GetNullableString(reader, "ObservacionSuscripcion"),
            Mensaje = GetNullableString(reader, "Mensaje")
        };
    }

    public async Task<CuentaAdministradoraConfiguracionDto?> ObtenerConfiguracionCuentaAdministradoraAsync(int idCuentaAdministradora, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_SEG_ObtenerConfiguracionCuentaAdministradora", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdCuentaAdministradora", idCuentaAdministradora);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        if (!await reader.ReadAsync(cancellationToken))
        {
            return null;
        }

        return new CuentaAdministradoraConfiguracionDto
        {
            IdCuentaAdministradora = reader.GetInt32(reader.GetOrdinal("IdCuentaAdministradora")),
            CodigoCuenta = reader.GetString(reader.GetOrdinal("CodigoCuenta")),
            NombreCuenta = reader.GetString(reader.GetOrdinal("NombreCuenta")),
            CorreoPrincipal = GetNullableString(reader, "CorreoPrincipal"),
            TelefonoPrincipal = GetNullableString(reader, "TelefonoPrincipal"),
            IdCuentaAdministradoraConfiguracion = GetNullableInt(reader, "IdCuentaAdministradoraConfiguracion"),
            NombreResponsablePrincipal = GetNullableString(reader, "NombreResponsablePrincipal"),
            CorreoAdministrativo = GetNullableString(reader, "CorreoAdministrativo"),
            TelefonoAdministrativo = GetNullableString(reader, "TelefonoAdministrativo"),
            IdEmpresaPredeterminada = GetNullableInt(reader, "IdEmpresaPredeterminada"),
            CodigoEmpresaPredeterminada = GetNullableString(reader, "CodigoEmpresaPredeterminada"),
            RazonSocialEmpresaPredeterminada = GetNullableString(reader, "RazonSocialEmpresaPredeterminada"),
            ObservacionAdministrativa = GetNullableString(reader, "ObservacionAdministrativa"),
            IdCuentaAdministradoraFacturacion = GetNullableInt(reader, "IdCuentaAdministradoraFacturacion"),
            TipoComprobantePreferido = GetNullableString(reader, "TipoComprobantePreferido"),
            TipoDocumentoFacturacion = GetNullableString(reader, "TipoDocumentoFacturacion"),
            NumeroDocumento = GetNullableString(reader, "NumeroDocumento"),
            NombreFacturacion = GetNullableString(reader, "NombreFacturacion"),
            RazonSocialFacturacion = GetNullableString(reader, "RazonSocialFacturacion"),
            CorreoFacturacion = GetNullableString(reader, "CorreoFacturacion"),
            TelefonoFacturacion = GetNullableString(reader, "TelefonoFacturacion"),
            DireccionFiscal = GetNullableString(reader, "DireccionFiscal"),
            Ubigeo = GetNullableString(reader, "Ubigeo"),
            Distrito = GetNullableString(reader, "Distrito"),
            Provincia = GetNullableString(reader, "Provincia"),
            Departamento = GetNullableString(reader, "Departamento"),
            ObservacionFacturacion = GetNullableString(reader, "ObservacionFacturacion")
        };
    }

    public async Task GuardarConfiguracionCuentaAdministradoraAsync(GuardarConfiguracionCuentaAdministradoraRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_SEG_GuardarConfiguracionCuentaAdministradora", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdCuentaAdministradora", request.IdCuentaAdministradora);
        command.Parameters.AddWithValue("@NombreResponsablePrincipal", DBNullIfNull(request.NombreResponsablePrincipal));
        command.Parameters.AddWithValue("@CorreoAdministrativo", DBNullIfNull(request.CorreoAdministrativo));
        command.Parameters.AddWithValue("@TelefonoAdministrativo", DBNullIfNull(request.TelefonoAdministrativo));
        command.Parameters.AddWithValue("@IdEmpresaPredeterminada", request.IdEmpresaPredeterminada ?? (object)DBNull.Value);
        command.Parameters.AddWithValue("@ObservacionAdministrativa", DBNullIfNull(request.ObservacionAdministrativa));
        command.Parameters.AddWithValue("@TipoComprobantePreferido", request.TipoComprobantePreferido);
        command.Parameters.AddWithValue("@TipoDocumentoFacturacion", request.TipoDocumentoFacturacion);
        command.Parameters.AddWithValue("@NumeroDocumento", DBNullIfNull(request.NumeroDocumento));
        command.Parameters.AddWithValue("@NombreFacturacion", DBNullIfNull(request.NombreFacturacion));
        command.Parameters.AddWithValue("@RazonSocialFacturacion", DBNullIfNull(request.RazonSocialFacturacion));
        command.Parameters.AddWithValue("@CorreoFacturacion", DBNullIfNull(request.CorreoFacturacion));
        command.Parameters.AddWithValue("@TelefonoFacturacion", DBNullIfNull(request.TelefonoFacturacion));
        command.Parameters.AddWithValue("@DireccionFiscal", DBNullIfNull(request.DireccionFiscal));
        command.Parameters.AddWithValue("@Ubigeo", DBNullIfNull(request.Ubigeo));
        command.Parameters.AddWithValue("@Distrito", DBNullIfNull(request.Distrito));
        command.Parameters.AddWithValue("@Provincia", DBNullIfNull(request.Provincia));
        command.Parameters.AddWithValue("@Departamento", DBNullIfNull(request.Departamento));
        command.Parameters.AddWithValue("@ObservacionFacturacion", DBNullIfNull(request.ObservacionFacturacion));
        command.Parameters.AddWithValue("@UsuarioRegistro", DBNullIfNull(request.UsuarioRegistro));

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task<IReadOnlyCollection<CuentaAdministradoraEmpresaDto>> ListarEmpresasCuentaAdministradoraAsync(int idCuentaAdministradora, CancellationToken cancellationToken = default)
    {
        var result = new List<CuentaAdministradoraEmpresaDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_SEG_ListarEmpresasCuentaAdministradora", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdCuentaAdministradora", idCuentaAdministradora);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new CuentaAdministradoraEmpresaDto
            {
                IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
                CodigoEmpresa = reader.GetString(reader.GetOrdinal("CodigoEmpresa")),
                RazonSocial = reader.GetString(reader.GetOrdinal("RazonSocial")),
                NombreComercial = GetNullableString(reader, "NombreComercial"),
                Ruc = GetNullableString(reader, "Ruc"),
                Estado = reader.GetBoolean(reader.GetOrdinal("Estado"))
            });
        }

        return result;
    }

    public async Task<IReadOnlyCollection<CuentaAdministradoraUsuarioDto>> ListarUsuariosCuentaAdministradoraAsync(int idCuentaAdministradora, CancellationToken cancellationToken = default)
    {
        var result = new List<CuentaAdministradoraUsuarioDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_SEG_ListarUsuariosCuentaAdministradora", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdCuentaAdministradora", idCuentaAdministradora);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new CuentaAdministradoraUsuarioDto
            {
                IdUsuarioCuentaAdministradora = reader.GetInt32(reader.GetOrdinal("IdUsuarioCuentaAdministradora")),
                AspNetUserId = reader.GetString(reader.GetOrdinal("AspNetUserId")),
                CorreoUsuario = reader.GetString(reader.GetOrdinal("CorreoUsuario")),
                NombreCompleto = GetNullableString(reader, "NombreCompleto"),
                Telefono = GetNullableString(reader, "Telefono"),
                RolCuenta = reader.GetString(reader.GetOrdinal("RolCuenta")),
                EsCuentaPredeterminada = reader.GetBoolean(reader.GetOrdinal("EsCuentaPredeterminada")),
                Estado = reader.GetBoolean(reader.GetOrdinal("Estado")),
                CantidadEmpresasAsignadas = GetNullableInt(reader, "CantidadEmpresasAsignadas") ?? 0,
                EmpresasAsignadas = GetNullableString(reader, "EmpresasAsignadas")
            });
        }

        return result;
    }

    public async Task AsignarUsuarioCuentaAdministradoraAsync(AsignarUsuarioCuentaAdministradoraRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_SEG_AsignarUsuarioCuentaAdministradora", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@AspNetUserId", request.AspNetUserId);
        command.Parameters.AddWithValue("@IdCuentaAdministradora", request.IdCuentaAdministradora);
        command.Parameters.AddWithValue("@RolCuenta", request.RolCuenta);
        command.Parameters.AddWithValue("@EsCuentaPredeterminada", request.EsCuentaPredeterminada);
        command.Parameters.AddWithValue("@UsuarioRegistro", DBNullIfNull(request.UsuarioRegistro));

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task AsignarUsuarioEmpresaAsync(AsignarUsuarioEmpresaRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_SEG_AsignarUsuarioEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@AspNetUserId", request.AspNetUserId);
        command.Parameters.AddWithValue("@IdEmpresa", request.IdEmpresa);
        command.Parameters.AddWithValue("@EsEmpresaPredeterminada", request.EsEmpresaPredeterminada);
        command.Parameters.AddWithValue("@UsuarioRegistro", request.UsuarioRegistro);

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task DesactivarUsuarioCuentaAdministradoraAsync(string aspNetUserId, int idCuentaAdministradora, string? usuarioRegistro, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_SEG_DesactivarUsuarioCuentaAdministradora", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@AspNetUserId", aspNetUserId);
        command.Parameters.AddWithValue("@IdCuentaAdministradora", idCuentaAdministradora);
        command.Parameters.AddWithValue("@UsuarioRegistro", DBNullIfNull(usuarioRegistro));

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task DesactivarUsuarioEmpresaAsync(string aspNetUserId, int idEmpresa, string? usuarioRegistro, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_SEG_DesactivarUsuarioEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@AspNetUserId", aspNetUserId);
        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@UsuarioRegistro", DBNullIfNull(usuarioRegistro));

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task<IReadOnlyCollection<UsuarioCuentaEmpresaAsignadaDto>> ListarEmpresasUsuarioCuentaAdministradoraAsync(int idCuentaAdministradora, string aspNetUserId, CancellationToken cancellationToken = default)
    {
        var result = new List<UsuarioCuentaEmpresaAsignadaDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_SEG_ListarEmpresasUsuarioCuentaAdministradora", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdCuentaAdministradora", idCuentaAdministradora);
        command.Parameters.AddWithValue("@AspNetUserId", aspNetUserId);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new UsuarioCuentaEmpresaAsignadaDto
            {
                IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
                CodigoEmpresa = reader.GetString(reader.GetOrdinal("CodigoEmpresa")),
                RazonSocial = reader.GetString(reader.GetOrdinal("RazonSocial")),
                NombreComercial = GetNullableString(reader, "NombreComercial"),
                Ruc = GetNullableString(reader, "Ruc"),
                Asignado = reader.GetBoolean(reader.GetOrdinal("Asignado")),
                EsEmpresaPredeterminada = reader.GetBoolean(reader.GetOrdinal("EsEmpresaPredeterminada")),
                IdUsuarioEmpresa = GetNullableInt(reader, "IdUsuarioEmpresa")
            });
        }

        return result;
    }

    public async Task<IReadOnlyCollection<UsuarioCuentaPermisoDto>> ListarPermisosUsuarioCuentaAsync(int idUsuarioCuentaAdministradora, CancellationToken cancellationToken = default)
    {
        return await ListarPermisosAsync(
            "dbo.usp_SEG_ListarPermisosUsuarioCuenta",
            "@IdUsuarioCuentaAdministradora",
            idUsuarioCuentaAdministradora,
            cancellationToken);
    }

    public async Task GuardarUsuarioCuentaPermisoAsync(GuardarUsuarioCuentaPermisoRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_SEG_GuardarUsuarioCuentaPermiso", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdUsuarioCuentaAdministradora", request.IdUsuarioCuentaAdministradora);
        command.Parameters.AddWithValue("@IdModuloSistema", request.IdModuloSistema);
        command.Parameters.AddWithValue("@PuedeVer", request.PuedeVer ?? (object)DBNull.Value);
        command.Parameters.AddWithValue("@PuedeCrear", request.PuedeCrear ?? (object)DBNull.Value);
        command.Parameters.AddWithValue("@PuedeEditar", request.PuedeEditar ?? (object)DBNull.Value);
        command.Parameters.AddWithValue("@PuedeEliminar", request.PuedeEliminar ?? (object)DBNull.Value);
        command.Parameters.AddWithValue("@UsuarioRegistro", DBNullIfNull(request.UsuarioRegistro));

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task<IReadOnlyCollection<UsuarioCuentaPermisoDto>> ListarPermisosUsuarioEmpresaAsync(int idUsuarioEmpresa, CancellationToken cancellationToken = default)
    {
        return await ListarPermisosAsync(
            "dbo.usp_SEG_ListarPermisosUsuarioEmpresa",
            "@IdUsuarioEmpresa",
            idUsuarioEmpresa,
            cancellationToken);
    }

    public async Task GuardarUsuarioEmpresaPermisoAsync(GuardarUsuarioEmpresaPermisoRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_SEG_GuardarUsuarioEmpresaPermiso", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdUsuarioEmpresa", request.IdUsuarioEmpresa);
        command.Parameters.AddWithValue("@IdModuloSistema", request.IdModuloSistema);
        command.Parameters.AddWithValue("@PuedeVer", request.PuedeVer ?? (object)DBNull.Value);
        command.Parameters.AddWithValue("@PuedeCrear", request.PuedeCrear ?? (object)DBNull.Value);
        command.Parameters.AddWithValue("@PuedeEditar", request.PuedeEditar ?? (object)DBNull.Value);
        command.Parameters.AddWithValue("@PuedeEliminar", request.PuedeEliminar ?? (object)DBNull.Value);
        command.Parameters.AddWithValue("@UsuarioRegistro", DBNullIfNull(request.UsuarioRegistro));

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    private async Task<IReadOnlyCollection<UsuarioCuentaPermisoDto>> ListarPermisosAsync(string storedProcedure, string parameterName, int parameterValue, CancellationToken cancellationToken)
    {
        var result = new List<UsuarioCuentaPermisoDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand(storedProcedure, connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue(parameterName, parameterValue);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new UsuarioCuentaPermisoDto
            {
                IdModuloSistema = reader.GetInt32(reader.GetOrdinal("IdModuloSistema")),
                CodigoModulo = reader.GetString(reader.GetOrdinal("CodigoModulo")),
                NombreModulo = reader.GetString(reader.GetOrdinal("NombreModulo")),
                GrupoMenu = GetNullableString(reader, "GrupoMenu"),
                RolCuenta = GetNullableString(reader, "RolCuenta"),
                PuedeVerRol = GetNullableBool(reader, "PuedeVerRol"),
                PuedeCrearRol = GetNullableBool(reader, "PuedeCrearRol"),
                PuedeEditarRol = GetNullableBool(reader, "PuedeEditarRol"),
                PuedeEliminarRol = GetNullableBool(reader, "PuedeEliminarRol"),
                PuedeVerOverride = GetNullableBool(reader, "PuedeVerOverride"),
                PuedeCrearOverride = GetNullableBool(reader, "PuedeCrearOverride"),
                PuedeEditarOverride = GetNullableBool(reader, "PuedeEditarOverride"),
                PuedeEliminarOverride = GetNullableBool(reader, "PuedeEliminarOverride"),
                PuedeVerEfectivo = reader.GetBoolean(reader.GetOrdinal("PuedeVerEfectivo")),
                PuedeCrearEfectivo = reader.GetBoolean(reader.GetOrdinal("PuedeCrearEfectivo")),
                PuedeEditarEfectivo = reader.GetBoolean(reader.GetOrdinal("PuedeEditarEfectivo")),
                PuedeEliminarEfectivo = reader.GetBoolean(reader.GetOrdinal("PuedeEliminarEfectivo"))
            });
        }

        return result;
    }

    private static CuentaSuscripcionResumenDto MapearCuentaSuscripcionResumen(SqlDataReader reader)
    {
        return new CuentaSuscripcionResumenDto
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
            TipoCobro = reader.IsDBNull(reader.GetOrdinal("TipoCobro")) ? null : reader.GetString(reader.GetOrdinal("TipoCobro")),
            DiasGracia = reader.IsDBNull(reader.GetOrdinal("DiasGracia")) ? 5 : reader.GetInt32(reader.GetOrdinal("DiasGracia")),
            FechaFinGracia = reader.IsDBNull(reader.GetOrdinal("FechaFinGracia")) ? null : DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaFinGracia"))),
            EmpresasPermitidas = reader.IsDBNull(reader.GetOrdinal("EmpresasPermitidas")) ? null : reader.GetInt32(reader.GetOrdinal("EmpresasPermitidas")),
            UsuariosPermitidos = reader.IsDBNull(reader.GetOrdinal("UsuariosPermitidos")) ? null : reader.GetInt32(reader.GetOrdinal("UsuariosPermitidos")),
            Activo = !reader.IsDBNull(reader.GetOrdinal("Activo")) && reader.GetBoolean(reader.GetOrdinal("Activo")),
            Observacion = reader.IsDBNull(reader.GetOrdinal("Observacion")) ? null : reader.GetString(reader.GetOrdinal("Observacion")),
            AspNetUserId = reader.IsDBNull(reader.GetOrdinal("AspNetUserId")) ? null : reader.GetString(reader.GetOrdinal("AspNetUserId")),
            NombreCompleto = reader.IsDBNull(reader.GetOrdinal("NombreCompleto")) ? null : reader.GetString(reader.GetOrdinal("NombreCompleto")),
            Telefono = reader.IsDBNull(reader.GetOrdinal("Telefono")) ? null : reader.GetString(reader.GetOrdinal("Telefono")),
            Email = reader.IsDBNull(reader.GetOrdinal("Email")) ? null : reader.GetString(reader.GetOrdinal("Email"))
        };
    }

    private static object DBNullIfNull(string? value)
        => string.IsNullOrWhiteSpace(value) ? DBNull.Value : value.Trim();

    private static string? GetNullableString(SqlDataReader reader, string columnName)
    {
        var ordinal = reader.GetOrdinal(columnName);
        return reader.IsDBNull(ordinal) ? null : reader.GetString(ordinal);
    }

    private static int? GetNullableInt(SqlDataReader reader, string columnName)
    {
        var ordinal = reader.GetOrdinal(columnName);
        return reader.IsDBNull(ordinal) ? null : reader.GetInt32(ordinal);
    }

    private static bool? GetNullableBool(SqlDataReader reader, string columnName)
    {
        var ordinal = reader.GetOrdinal(columnName);
        return reader.IsDBNull(ordinal) ? null : reader.GetBoolean(ordinal);
    }

    private static DateOnly? GetNullableDateOnly(SqlDataReader reader, string columnName)
    {
        var ordinal = reader.GetOrdinal(columnName);
        return reader.IsDBNull(ordinal)
            ? null
            : DateOnly.FromDateTime(reader.GetDateTime(ordinal));
    }
}
