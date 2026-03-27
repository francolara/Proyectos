-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Sprint 7 - Automatizacion de recordatorios por correo y marcado automatico de no-show.
-- =============================================

IF COL_LENGTH('dbo.Reservas', 'RecordatorioEnviado') IS NULL
BEGIN
    ALTER TABLE dbo.Reservas
    ADD RecordatorioEnviado BIT NOT NULL CONSTRAINT DF_Reservas_RecordatorioEnviado DEFAULT (0);
END;
GO

IF COL_LENGTH('dbo.Reservas', 'FechaRecordatorio') IS NULL
BEGIN
    ALTER TABLE dbo.Reservas
    ADD FechaRecordatorio DATETIME2 NULL;
END;
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_RecordatoriosPendientes
    @FechaHoraDesde DATETIME2,
    @FechaHoraHasta DATETIME2
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            r.Id AS ReservaId,
            s.NegocioId,
            c.NombresORazonSocial AS Cliente,
            c.Correo,
            s.Nombre AS Sede,
            e.Nombre AS Espacio,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
        WHERE r.Estado IN (1, 2)
          AND r.RecordatorioEnviado = 0
          AND c.Correo IS NOT NULL
          AND LTRIM(RTRIM(c.Correo)) <> N''
          AND DATEADD(MINUTE, DATEDIFF(MINUTE, 0, r.HoraInicio), CAST(r.Fecha AS DATETIME2)) BETWEEN @FechaHoraDesde AND @FechaHoraHasta
        ORDER BY r.Fecha, r.HoraInicio;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_MarcarRecordatorioEnviado
    @NegocioId INT,
    @ReservaId INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE r
        SET r.RecordatorioEnviado = 1,
            r.FechaRecordatorio = SYSUTCDATETIME(),
            r.FechaActualizacion = SYSUTCDATETIME(),
            r.UsuarioActualizacion = @Usuario
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE r.Id = @ReservaId
          AND s.NegocioId = @NegocioId;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_AutoNoShow
    @FechaHoraActual DATETIME2,
    @ToleranciaMinutos INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @Actualizadas TABLE
        (
            ReservaId INT NOT NULL,
            NegocioId INT NOT NULL
        );

        UPDATE r
        SET r.Estado = 6,
            r.FechaActualizacion = SYSUTCDATETIME(),
            r.UsuarioActualizacion = @Usuario
        OUTPUT inserted.Id, s.NegocioId
        INTO @Actualizadas (ReservaId, NegocioId)
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE r.Estado IN (1, 2)
          AND DATEADD(MINUTE, @ToleranciaMinutos, DATEADD(MINUTE, DATEDIFF(MINUTE, 0, r.HoraInicio), CAST(r.Fecha AS DATETIME2))) <= @FechaHoraActual;

        DECLARE @ReservaId INT, @NegocioId INT;
        DECLARE c CURSOR LOCAL FAST_FORWARD FOR
            SELECT a.ReservaId, a.NegocioId
            FROM @Actualizadas a;

        OPEN c;
        FETCH NEXT FROM c INTO @ReservaId, @NegocioId;

        WHILE @@FETCH_STATUS = 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @ReservaId);
            EXEC dbo.Sp_Auditoria_Registrar
                @NegocioId = @NegocioId,
                @Modulo = N'RESERVAS',
                @Accion = N'AUTO_NOSHOW',
                @Entidad = N'Reserva',
                @EntidadId = @EntidadIdAudit,
                @Usuario = @Usuario,
                @DetalleJson = NULL;

            FETCH NEXT FROM c INTO @ReservaId, @NegocioId;
        END;

        CLOSE c;
        DEALLOCATE c;

        SELECT COUNT(1) AS TotalActualizadas FROM @Actualizadas;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
