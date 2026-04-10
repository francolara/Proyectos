/*
Firma: Codex - 09/04/2026
Descripcion: Ajusta modulo Pagos (listado por reserva, busqueda incremental de reserva, resumen referencial, validaciones de politica al crear pago, validacion de 2do pago = saldo restante, monto de reserva + saldo pendiente con simbolo de moneda desde Monedas/MonedasSuperMaestro, banderas PagadaCompleta/TieneComprobanteActivo para emision de comprobantes y eliminacion total de pagos por reserva con cancelacion).
*/
USE [DbSportCenter]
GO

CREATE OR ALTER PROCEDURE [dbo].[Sp_Pagos_Listar]
    @NegocioId INT,
    @SedeId INT = NULL,
    @Buscar NVARCHAR(120) = NULL,
    @Pagina INT = 1,
    @TamanoPagina INT = 20,
    @TotalRegistros INT OUTPUT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        IF @Pagina < 1 SET @Pagina = 1;
        IF @TamanoPagina < 1 SET @TamanoPagina = 20;

        DECLARE @Offset INT = (@Pagina - 1) * @TamanoPagina;
        DECLARE @BuscarTrim NVARCHAR(120) = NULLIF(LTRIM(RTRIM(@Buscar)), N'');

        CREATE TABLE #ReservasFiltradas
        (
            ReservaId INT NOT NULL,
            ReservaCodigo NVARCHAR(25) NOT NULL,
            Sede NVARCHAR(200) NOT NULL,
            Espacio NVARCHAR(200) NOT NULL,
            Cliente NVARCHAR(200) NOT NULL,
            Fecha DATE NOT NULL,
            MontoTotal DECIMAL(10,2) NOT NULL,
            SaldoPendiente DECIMAL(10,2) NOT NULL,
            FormaPagoResumen NVARCHAR(500) NOT NULL,
            CantidadPagos INT NOT NULL,
            MonedaSimbolo NVARCHAR(10) NOT NULL,
            PagadaCompleta BIT NOT NULL,
            TieneComprobanteActivo BIT NOT NULL
        );

        ;WITH ReservasConPago AS
        (
            SELECT
                r.Id AS ReservaId,
                s.Nombre AS Sede,
                e.Nombre AS Espacio,
                c.NombresORazonSocial AS Cliente,
                r.Fecha,
                CAST(r.Total AS DECIMAL(10,2)) AS MontoTotal,
                CAST(r.Total - SUM(p.Monto) AS DECIMAL(10,2)) AS SaldoPendiente,
                CAST(CASE WHEN r.Estado = 4 AND (r.Total - SUM(p.Monto)) <= 0 THEN 1 ELSE 0 END AS BIT) AS PagadaCompleta,
                COUNT(p.Id) AS CantidadPagos,
                STRING_AGG(fp.Nombre, N', ') WITHIN GROUP (ORDER BY fp.Nombre) AS FormaPagoResumen,
                COALESCE(ms.Simbolo, N'S/') AS MonedaSimbolo
            FROM dbo.Reservas r
            INNER JOIN dbo.Pagos p ON p.ReservaId = r.Id
            INNER JOIN dbo.FormasPago fp ON fp.Id = p.FormaPago
            INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
            INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
            LEFT JOIN dbo.Monedas m ON m.Id = n.MonedaId
            LEFT JOIN dbo.MonedasSuperMaestro ms ON ms.Id = m.MonedaSuperId
            INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
            WHERE s.NegocioId = @NegocioId
              AND (@SedeId IS NULL OR s.Id = @SedeId)
            GROUP BY r.Id, s.Nombre, e.Nombre, c.NombresORazonSocial, r.Fecha, r.Total, r.Estado, ms.Simbolo
        )
        INSERT INTO #ReservasFiltradas
        (
            ReservaId,
            ReservaCodigo,
            Sede,
            Espacio,
            Cliente,
            Fecha,
            MontoTotal,
            SaldoPendiente,
            FormaPagoResumen,
            CantidadPagos,
            MonedaSimbolo,
            PagadaCompleta,
            TieneComprobanteActivo
        )
        SELECT
            x.ReservaId,
            CONCAT(N'#', CONVERT(NVARCHAR(20), x.ReservaId)) AS ReservaCodigo,
            x.Sede,
            x.Espacio,
            x.Cliente,
            x.Fecha,
            x.MontoTotal,
            x.SaldoPendiente,
            x.FormaPagoResumen,
            x.CantidadPagos,
            x.MonedaSimbolo,
            x.PagadaCompleta,
            CAST(CASE
                WHEN EXISTS
                (
                    SELECT 1
                    FROM dbo.ComprobantesElectronicos ce
                    WHERE ce.NegocioId = @NegocioId
                      AND ce.ReservaId = x.ReservaId
                      AND ce.Estado <> 5
                ) THEN 1 ELSE 0
            END AS BIT) AS TieneComprobanteActivo
        FROM ReservasConPago x
        WHERE @BuscarTrim IS NULL
           OR CONVERT(NVARCHAR(20), x.ReservaId) LIKE N'%' + @BuscarTrim + N'%'
           OR x.Sede LIKE N'%' + @BuscarTrim + N'%'
           OR x.Espacio LIKE N'%' + @BuscarTrim + N'%'
           OR x.Cliente LIKE N'%' + @BuscarTrim + N'%'
           OR x.FormaPagoResumen LIKE N'%' + @BuscarTrim + N'%'
           OR CONVERT(NVARCHAR(10), x.Fecha, 103) LIKE N'%' + @BuscarTrim + N'%';

        SELECT @TotalRegistros = COUNT(1)
        FROM #ReservasFiltradas;

        SELECT
            ReservaId,
            ReservaCodigo,
            Sede,
            Espacio,
            Cliente,
            Fecha,
            MontoTotal,
            SaldoPendiente,
            FormaPagoResumen,
            CantidadPagos,
            MonedaSimbolo,
            PagadaCompleta,
            TieneComprobanteActivo
        FROM #ReservasFiltradas
        ORDER BY Fecha DESC, ReservaId DESC
        OFFSET @Offset ROWS FETCH NEXT @TamanoPagina ROWS ONLY;

        DROP TABLE #ReservasFiltradas;
    END TRY
    BEGIN CATCH
        IF OBJECT_ID('tempdb..#ReservasFiltradas') IS NOT NULL
            DROP TABLE #ReservasFiltradas;

        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE [dbo].[Sp_Pagos_ObtenerPorId]
    @NegocioId INT,
    @Id INT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            r.Id AS ReservaId,
            CONCAT(N'#', CONVERT(NVARCHAR(20), r.Id)) AS ReservaCodigo,
            s.Nombre AS Sede,
            e.Nombre AS Espacio,
            c.NombresORazonSocial AS Cliente,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin,
            r.Total AS TotalReserva,
            COALESCE(SUM(p.Monto), 0) AS TotalPagado,
            (r.Total - COALESCE(SUM(p.Monto), 0)) AS SaldoPendiente,
            COALESCE(ms.Simbolo, N'S/') AS MonedaSimbolo,
            CAST(ISNULL(n.PoliticaConfirmacionPago, 0) AS INT) AS PoliticaConfirmacionPago,
            n.PorcentajeAdelantoMinimo
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
        INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
        LEFT JOIN dbo.Monedas m ON m.Id = n.MonedaId
        LEFT JOIN dbo.MonedasSuperMaestro ms ON ms.Id = m.MonedaSuperId
        LEFT JOIN dbo.Pagos p ON p.ReservaId = r.Id
        WHERE r.Id = @Id
          AND s.NegocioId = @NegocioId
        GROUP BY
            r.Id,
            s.Nombre,
            e.Nombre,
            c.NombresORazonSocial,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin,
            r.Total,
            ms.Simbolo,
            n.PoliticaConfirmacionPago,
            n.PorcentajeAdelantoMinimo;

        SELECT
            p.Id,
            p.FechaPago,
            p.Monto,
            p.FormaPago,
            fp.Nombre AS FormaPagoNombre,
            p.NumeroOperacion,
            p.Observacion
        FROM dbo.Pagos p
        INNER JOIN dbo.FormasPago fp ON fp.Id = p.FormaPago
        INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE r.Id = @Id
          AND s.NegocioId = @NegocioId
        ORDER BY p.FechaPago, p.Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE [dbo].[Sp_Pagos_Actualizar]
    @Id INT,
    @NegocioId INT,
    @Observacion NVARCHAR(300) = NULL,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.Pagos p
            INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
            INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
            WHERE p.Id = @Id
              AND s.NegocioId = @NegocioId
        )
            RAISERROR('No se encontro el pago para actualizar en el negocio.', 16, 1);

        UPDATE dbo.Pagos
        SET Observacion = NULLIF(LTRIM(RTRIM(@Observacion)), N''),
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id;

        DECLARE @EntidadIdAudit NVARCHAR(80) = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar
            @NegocioId = @NegocioId,
            @Modulo = N'PAGOS',
            @Accion = N'EDIT',
            @Entidad = N'Pago',
            @EntidadId = @EntidadIdAudit,
            @Usuario = @Usuario,
            @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE [dbo].[Sp_Pagos_Eliminar]
    @NegocioId INT,
    @Id INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @ReservaId INT;

        SELECT @ReservaId = p.ReservaId
        FROM dbo.Pagos p
        INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE p.Id = @Id
          AND s.NegocioId = @NegocioId;

        IF @ReservaId IS NULL
            RAISERROR('No se encontro el pago para eliminar en el negocio.', 16, 1);

        BEGIN TRANSACTION;

        DELETE FROM dbo.Pagos
        WHERE Id = @Id;

        DECLARE @PagadoRestante DECIMAL(10,2);
        DECLARE @CantidadPagosRestante INT;

        SELECT
            @PagadoRestante = COALESCE(SUM(p.Monto), 0),
            @CantidadPagosRestante = COUNT(1)
        FROM dbo.Pagos p
        WHERE p.ReservaId = @ReservaId;

        UPDATE dbo.Reservas
        SET Adelanto = @PagadoRestante,
            Saldo = (Total - @PagadoRestante),
            Estado = CASE WHEN @CantidadPagosRestante = 0 THEN 5 ELSE 2 END,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @ReservaId;

        DECLARE @EntidadIdAudit NVARCHAR(80) = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar
            @NegocioId = @NegocioId,
            @Modulo = N'PAGOS',
            @Accion = N'DELETE',
            @Entidad = N'Pago',
            @EntidadId = @EntidadIdAudit,
            @Usuario = @Usuario,
            @DetalleJson = NULL;

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0
            ROLLBACK TRANSACTION;

        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE [dbo].[Sp_Pagos_EliminarPorReserva]
    @NegocioId INT,
    @ReservaId INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.Reservas r
            INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
            WHERE r.Id = @ReservaId
              AND s.NegocioId = @NegocioId
        )
            RAISERROR('No se encontro la reserva para eliminar pagos en el negocio.', 16, 1);

        BEGIN TRANSACTION;

        DELETE FROM dbo.Pagos
        WHERE ReservaId = @ReservaId;

        UPDATE dbo.Reservas
        SET Adelanto = 0,
            Saldo = Total,
            Estado = 5,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @ReservaId;

        DECLARE @EntidadIdAudit NVARCHAR(80) = CONVERT(NVARCHAR(80), @ReservaId);
        EXEC dbo.Sp_Auditoria_Registrar
            @NegocioId = @NegocioId,
            @Modulo = N'PAGOS',
            @Accion = N'DELETE',
            @Entidad = N'ReservaPago',
            @EntidadId = @EntidadIdAudit,
            @Usuario = @Usuario,
            @DetalleJson = NULL;

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0
            ROLLBACK TRANSACTION;

        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE [dbo].[Sp_Combos_Reservas_Buscar]
    @NegocioId INT,
    @Buscar NVARCHAR(150) = NULL,
    @ReservaId INT = NULL,
    @Top INT = 40
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @BuscarNorm NVARCHAR(150) = NULLIF(LTRIM(RTRIM(@Buscar)), N'');
        SET @Top = CASE WHEN ISNULL(@Top, 0) < 1 THEN 40 WHEN @Top > 100 THEN 100 ELSE @Top END;

        ;WITH Fuente AS
        (
            SELECT
                r.Id,
                CONCAT(
                    N'#', r.Id,
                    N' - ',
                    c.NombresORazonSocial,
                    CASE
                        WHEN NULLIF(LTRIM(RTRIM(c.NombreEquipo)), N'') IS NULL THEN N''
                        ELSE CONCAT(N' [', LTRIM(RTRIM(c.NombreEquipo)), N']')
                    END,
                    N' | ',
                    CONVERT(NVARCHAR(10), r.Fecha, 103),
                    N' ',
                    CONVERT(NVARCHAR(5), r.HoraInicio),
                    N'-',
                    CONVERT(NVARCHAR(5), r.HoraFin),
                    N' | Saldo: ',
                    CONVERT(NVARCHAR(32), CAST((r.Total - COALESCE(r.Adelanto, 0)) AS DECIMAL(10,2)))
                ) AS ReservaTexto,
                r.Fecha,
                r.HoraInicio
            FROM dbo.Reservas r
            INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
            INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
            WHERE s.NegocioId = @NegocioId
              AND
              (
                  (@BuscarNorm IS NOT NULL AND
                   (
                       CONVERT(NVARCHAR(20), r.Id) LIKE N'%' + @BuscarNorm + N'%'
                       OR c.NombresORazonSocial LIKE N'%' + @BuscarNorm + N'%'
                       OR ISNULL(c.NombreEquipo, N'') LIKE N'%' + @BuscarNorm + N'%'
                       OR s.Nombre LIKE N'%' + @BuscarNorm + N'%'
                       OR e.Nombre LIKE N'%' + @BuscarNorm + N'%'
                       OR CONVERT(NVARCHAR(10), r.Fecha, 103) LIKE N'%' + @BuscarNorm + N'%'
                   ))
                  OR (@ReservaId IS NOT NULL AND r.Id = @ReservaId)
              )
        )
        SELECT TOP (@Top)
            f.Id,
            f.ReservaTexto
        FROM Fuente f
        ORDER BY
            CASE WHEN @ReservaId IS NOT NULL AND f.Id = @ReservaId THEN 0 ELSE 1 END,
            f.Fecha DESC,
            f.HoraInicio DESC,
            f.Id DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE [dbo].[Sp_Pagos_Crear]
    @NegocioId INT,
    @ReservaId INT,
    @FechaPago DATETIME2,
    @Monto DECIMAL(10,2),
    @FormaPago INT,
    @NumeroOperacion NVARCHAR(50) = NULL,
    @Observacion NVARCHAR(300) = NULL,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        IF @Monto <= 0
            RAISERROR('El monto debe ser mayor que cero.', 16, 1);

        IF NOT EXISTS (SELECT 1 FROM dbo.FormasPago WHERE Id = @FormaPago AND Activo = 1)
            RAISERROR('La forma de pago no es valida.', 16, 1);

        DECLARE @TotalReserva DECIMAL(10,2);
        DECLARE @PagadoActual DECIMAL(10,2);
        DECLARE @NuevoPagado DECIMAL(10,2);
        DECLARE @CantidadPagos INT;
        DECLARE @PoliticaConfirmacionPago TINYINT = 0;
        DECLARE @PorcentajeAdelantoMinimo DECIMAL(5,2) = NULL;
        DECLARE @MontoMinimoAdelanto DECIMAL(10,2) = NULL;

        SELECT
            @TotalReserva = r.Total,
            @PoliticaConfirmacionPago = ISNULL(n.PoliticaConfirmacionPago, 0),
            @PorcentajeAdelantoMinimo = n.PorcentajeAdelantoMinimo
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
        WHERE r.Id = @ReservaId
          AND s.NegocioId = @NegocioId;

        IF @TotalReserva IS NULL
            RAISERROR('Reserva invalida para el negocio.', 16, 1);

        SELECT
            @PagadoActual = COALESCE(SUM(p.Monto), 0),
            @CantidadPagos = COUNT(1)
        FROM dbo.Pagos p
        WHERE p.ReservaId = @ReservaId;

        IF COALESCE(@CantidadPagos, 0) >= 2
            RAISERROR('La reserva ya tiene 2 pagos registrados. No se pueden registrar mas pagos.', 16, 1);

        IF COALESCE(@CantidadPagos, 0) = 1
        BEGIN
            DECLARE @SaldoRestante DECIMAL(10,2) = (@TotalReserva - @PagadoActual);
            IF ABS(@Monto - @SaldoRestante) > 0.009
                RAISERROR('Al registrar el segundo pago, el monto debe ser exactamente el saldo restante de la reserva.', 16, 1);
        END

        SET @NuevoPagado = @PagadoActual + @Monto;
        IF @NuevoPagado > @TotalReserva
            RAISERROR('El pago excede el total de la reserva.', 16, 1);

        IF @PoliticaConfirmacionPago = 2 AND @NuevoPagado < @TotalReserva
            RAISERROR('La configuracion del negocio exige pago total (100%) para confirmar la reserva.', 16, 1);

        IF @PoliticaConfirmacionPago = 1
        BEGIN
            SET @PorcentajeAdelantoMinimo = ISNULL(@PorcentajeAdelantoMinimo, 0);
            IF @PorcentajeAdelantoMinimo > 0
            BEGIN
                SET @MontoMinimoAdelanto = ROUND((@TotalReserva * @PorcentajeAdelantoMinimo) / 100.0, 2);
                IF @NuevoPagado < @MontoMinimoAdelanto AND @NuevoPagado < @TotalReserva
                    RAISERROR('El pago acumulado no alcanza el adelanto minimo configurado para confirmar la reserva.', 16, 1);
            END
        END

        BEGIN TRANSACTION;

        INSERT INTO dbo.Pagos
        (
            ReservaId, FechaPago, Monto, FormaPago, NumeroOperacion, Observacion,
            FechaCreacion, UsuarioCreacion
        )
        VALUES
        (
            @ReservaId, @FechaPago, @Monto, @FormaPago, @NumeroOperacion, @Observacion,
            SYSUTCDATETIME(), @Usuario
        );

        UPDATE r
        SET Adelanto = @NuevoPagado,
            Saldo = (r.Total - @NuevoPagado),
            Estado = CASE
                        WHEN (r.Total - @NuevoPagado) <= 0 THEN 4
                        WHEN @NuevoPagado > 0 THEN 2
                        ELSE r.Estado
                     END,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        FROM dbo.Reservas r
        WHERE r.Id = @ReservaId;

        DECLARE @Id INT;
        SET @Id = SCOPE_IDENTITY();
        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'PAGOS', @Accion = N'CREATE', @Entidad = N'Pago', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;

        COMMIT TRANSACTION;
        SELECT @Id;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0
            ROLLBACK TRANSACTION;

        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
