-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Reglas de negocio sprint 2 para reservas y pagos (validaciones, saldo, estado).
-- =============================================

CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_Crear
    @NegocioId INT,
    @EspacioDeportivoId INT,
    @ClienteId INT,
    @Fecha DATE,
    @HoraInicio TIME,
    @HoraFin TIME,
    @Total DECIMAL(10,2),
    @Adelanto DECIMAL(10,2),
    @Estado INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @HoraFin <= @HoraInicio
            RAISERROR('La hora fin debe ser mayor que la hora inicio.', 16, 1);

        IF @Total < 0 OR @Adelanto < 0
            RAISERROR('Los montos no pueden ser negativos.', 16, 1);

        IF @Adelanto > @Total
            RAISERROR('El adelanto no puede ser mayor que el total.', 16, 1);

        IF NOT EXISTS (
            SELECT 1
            FROM dbo.EspaciosDeportivos e
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
            WHERE e.Id = @EspacioDeportivoId
              AND s.NegocioId = @NegocioId
              AND e.Estado = 1
        )
            RAISERROR('El espacio deportivo no esta disponible para este negocio.', 16, 1);

        IF EXISTS (
            SELECT 1
            FROM dbo.Reservas r
            WHERE r.EspacioDeportivoId = @EspacioDeportivoId
              AND r.Fecha = @Fecha
              AND r.Estado NOT IN (5, 6)
              AND @HoraInicio < r.HoraFin
              AND @HoraFin > r.HoraInicio
        )
            RAISERROR('Cruce de horario detectado.', 16, 1);

        INSERT INTO dbo.Reservas
        (
            EspacioDeportivoId, ClienteId, Fecha, HoraInicio, HoraFin,
            Estado, Total, Adelanto, Saldo, FechaRegistro, UsuarioCreacion
        )
        VALUES
        (
            @EspacioDeportivoId, @ClienteId, @Fecha, @HoraInicio, @HoraFin,
            @Estado, @Total, @Adelanto, (@Total - @Adelanto), SYSUTCDATETIME(), @Usuario
        );

        DECLARE @Id INT;
        SET @Id = SCOPE_IDENTITY();
        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'RESERVAS', @Accion = N'CREATE', @Entidad = N'Reserva', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        SELECT @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_Actualizar
    @Id INT,
    @NegocioId INT,
    @EspacioDeportivoId INT,
    @ClienteId INT,
    @Fecha DATE,
    @HoraInicio TIME,
    @HoraFin TIME,
    @Total DECIMAL(10,2),
    @Adelanto DECIMAL(10,2),
    @Estado INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @HoraFin <= @HoraInicio
            RAISERROR('La hora fin debe ser mayor que la hora inicio.', 16, 1);

        IF @Total < 0 OR @Adelanto < 0
            RAISERROR('Los montos no pueden ser negativos.', 16, 1);

        IF @Adelanto > @Total
            RAISERROR('El adelanto no puede ser mayor que el total.', 16, 1);

        IF NOT EXISTS (
            SELECT 1
            FROM dbo.EspaciosDeportivos e
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
            WHERE e.Id = @EspacioDeportivoId
              AND s.NegocioId = @NegocioId
              AND e.Estado = 1
        )
            RAISERROR('El espacio deportivo no esta disponible para este negocio.', 16, 1);

        IF EXISTS (
            SELECT 1
            FROM dbo.Reservas r
            WHERE r.EspacioDeportivoId = @EspacioDeportivoId
              AND r.Fecha = @Fecha
              AND r.Estado NOT IN (5, 6)
              AND r.Id <> @Id
              AND @HoraInicio < r.HoraFin
              AND @HoraFin > r.HoraInicio
        )
            RAISERROR('Cruce de horario detectado.', 16, 1);

        UPDATE r
        SET r.EspacioDeportivoId = @EspacioDeportivoId,
            r.ClienteId = @ClienteId,
            r.Fecha = @Fecha,
            r.HoraInicio = @HoraInicio,
            r.HoraFin = @HoraFin,
            r.Total = @Total,
            r.Adelanto = @Adelanto,
            r.Saldo = (@Total - @Adelanto),
            r.Estado = @Estado,
            r.FechaActualizacion = SYSUTCDATETIME(),
            r.UsuarioActualizacion = @Usuario
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE r.Id = @Id
          AND s.NegocioId = @NegocioId;

        IF @@ROWCOUNT > 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
            EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'RESERVAS', @Accion = N'EDIT', @Entidad = N'Reserva', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        END;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Pagos_Crear
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

        DECLARE @TotalReserva DECIMAL(10,2);
        DECLARE @PagadoActual DECIMAL(10,2);
        DECLARE @NuevoPagado DECIMAL(10,2);

        SELECT @TotalReserva = r.Total
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE r.Id = @ReservaId
          AND s.NegocioId = @NegocioId;

        IF @TotalReserva IS NULL
            RAISERROR('Reserva invalida para el negocio.', 16, 1);

        SELECT @PagadoActual = COALESCE(SUM(p.Monto), 0)
        FROM dbo.Pagos p
        WHERE p.ReservaId = @ReservaId;

        SET @NuevoPagado = @PagadoActual + @Monto;
        IF @NuevoPagado > @TotalReserva
            RAISERROR('El pago excede el total de la reserva.', 16, 1);

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
            Estado = CASE WHEN (r.Total - @NuevoPagado) <= 0 AND r.Estado = 1 THEN 2 ELSE r.Estado END,
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

CREATE OR ALTER PROCEDURE dbo.Sp_Pagos_Actualizar
    @Id INT,
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

        DECLARE @ReservaAnteriorId INT;
        SELECT @ReservaAnteriorId = p.ReservaId FROM dbo.Pagos p WHERE p.Id = @Id;

        IF @ReservaAnteriorId IS NULL
            RETURN;

        IF NOT EXISTS (
            SELECT 1
            FROM dbo.Reservas r
            INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
            WHERE r.Id = @ReservaId
              AND s.NegocioId = @NegocioId
        )
            RAISERROR('Reserva invalida para el negocio.', 16, 1);

        DECLARE @TotalReserva DECIMAL(10,2);
        DECLARE @PagadoSinEste DECIMAL(10,2);
        DECLARE @NuevoPagado DECIMAL(10,2);

        SELECT @TotalReserva = r.Total FROM dbo.Reservas r WHERE r.Id = @ReservaId;

        SELECT @PagadoSinEste = COALESCE(SUM(p.Monto), 0)
        FROM dbo.Pagos p
        WHERE p.ReservaId = @ReservaId
          AND p.Id <> @Id;

        SET @NuevoPagado = @PagadoSinEste + @Monto;
        IF @NuevoPagado > @TotalReserva
            RAISERROR('El pago excede el total de la reserva.', 16, 1);

        BEGIN TRANSACTION;

        UPDATE dbo.Pagos
        SET ReservaId = @ReservaId,
            FechaPago = @FechaPago,
            Monto = @Monto,
            FormaPago = @FormaPago,
            NumeroOperacion = @NumeroOperacion,
            Observacion = @Observacion,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id;

        DECLARE @ReservaRecalculo TABLE (ReservaId INT PRIMARY KEY);
        INSERT INTO @ReservaRecalculo (ReservaId) VALUES (@ReservaId);
        IF @ReservaAnteriorId <> @ReservaId
            INSERT INTO @ReservaRecalculo (ReservaId) VALUES (@ReservaAnteriorId);

        UPDATE r
        SET Adelanto = x.Pagado,
            Saldo = (r.Total - x.Pagado),
            Estado = CASE WHEN (r.Total - x.Pagado) <= 0 AND r.Estado = 1 THEN 2 ELSE r.Estado END,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        FROM dbo.Reservas r
        INNER JOIN (
            SELECT rr.ReservaId, COALESCE(SUM(p.Monto), 0) AS Pagado
            FROM @ReservaRecalculo rr
            LEFT JOIN dbo.Pagos p ON p.ReservaId = rr.ReservaId
            GROUP BY rr.ReservaId
        ) x ON x.ReservaId = r.Id;

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'PAGOS', @Accion = N'EDIT', @Entidad = N'Pago', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;

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

CREATE OR ALTER PROCEDURE dbo.Sp_Pagos_Eliminar
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
            RETURN;

        BEGIN TRANSACTION;

        DELETE FROM dbo.Pagos WHERE Id = @Id;

        UPDATE r
        SET Adelanto = x.Pagado,
            Saldo = (r.Total - x.Pagado),
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        FROM dbo.Reservas r
        INNER JOIN (
            SELECT @ReservaId AS ReservaId, COALESCE(SUM(p.Monto), 0) AS Pagado
            FROM dbo.Pagos p
            WHERE p.ReservaId = @ReservaId
        ) x ON x.ReservaId = r.Id;

        IF @@ROWCOUNT > 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
            EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'PAGOS', @Accion = N'DELETE', @Entidad = N'Pago', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        END;

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
