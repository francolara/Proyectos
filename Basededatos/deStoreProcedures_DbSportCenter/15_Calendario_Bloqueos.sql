-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Sprint 5 - Calendario avanzado de reservas con drag/drop y bloqueos operativos. Devuelve todos los estados de reserva.
-- Firma:         Codex - 28/03/2026 | Ajuste de estados en calendario (incluye canceladas/no show) y color bloqueado unificado (#64748b).
-- =============================================

IF OBJECT_ID(N'dbo.BloqueosHorario', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.BloqueosHorario
    (
        Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        EspacioDeportivoId INT NOT NULL,
        Fecha DATE NOT NULL,
        HoraInicio TIME NOT NULL,
        HoraFin TIME NOT NULL,
        Motivo NVARCHAR(250) NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_BloqueosHorario_Activo DEFAULT (1),
        FechaCreacion DATETIME2 NOT NULL CONSTRAINT DF_BloqueosHorario_FechaCreacion DEFAULT (SYSUTCDATETIME()),
        UsuarioCreacion NVARCHAR(200) NULL,
        FechaActualizacion DATETIME2 NULL,
        UsuarioActualizacion NVARCHAR(200) NULL,
        CONSTRAINT FK_BloqueosHorario_EspaciosDeportivos_EspacioDeportivoId
            FOREIGN KEY (EspacioDeportivoId) REFERENCES dbo.EspaciosDeportivos (Id)
    );
END;
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_CalendarioEventos
    @NegocioId INT,
    @FechaDesde DATE,
    @FechaHasta DATE,
    @SedeId INT = NULL,
    @EspacioDeportivoId INT = NULL,
    @Estado INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            r.Id,
            CAST(N'RESERVA' AS NVARCHAR(20)) AS TipoEvento,
            CONCAT(e.Nombre, N' - ', c.NombresORazonSocial) AS Titulo,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin,
            r.Estado,
            CAST(
                CASE r.Estado
                    WHEN 1 THEN N'#f59f00'
                    WHEN 2 THEN N'#2f9e44'
                    WHEN 3 THEN N'#1971c2'
                    WHEN 4 THEN N'#495057'
                    WHEN 5 THEN N'#c92a2a'
                    WHEN 6 THEN N'#212529'
                    ELSE N'#6c757d'
                END
                AS NVARCHAR(20)
            ) AS Color,
            e.Id AS EspacioDeportivoId,
            e.Nombre AS Espacio,
            s.Nombre AS Sede
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
        WHERE s.NegocioId = @NegocioId
          AND r.Fecha BETWEEN @FechaDesde AND @FechaHasta
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND (@EspacioDeportivoId IS NULL OR e.Id = @EspacioDeportivoId)
          AND (@Estado IS NULL OR r.Estado = @Estado)

        UNION ALL

        SELECT
            b.Id,
            CAST(N'BLOQUEO' AS NVARCHAR(20)) AS TipoEvento,
            CONCAT(N'Bloqueado: ', b.Motivo) AS Titulo,
            b.Fecha,
            b.HoraInicio,
            b.HoraFin,
            NULL AS Estado,
            CAST(N'#64748b' AS NVARCHAR(20)) AS Color,
            e.Id AS EspacioDeportivoId,
            e.Nombre AS Espacio,
            s.Nombre AS Sede
        FROM dbo.BloqueosHorario b
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = b.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND b.Activo = 1
          AND b.Fecha BETWEEN @FechaDesde AND @FechaHasta
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND (@EspacioDeportivoId IS NULL OR e.Id = @EspacioDeportivoId)
        ORDER BY Fecha, HoraInicio;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_Mover
    @NegocioId INT,
    @Id INT,
    @Fecha DATE,
    @HoraInicio TIME,
    @HoraFin TIME,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @HoraFin <= @HoraInicio
            RAISERROR('La hora fin debe ser mayor a la hora inicio.', 16, 1);

        DECLARE @EspacioDeportivoId INT;

        SELECT @EspacioDeportivoId = r.EspacioDeportivoId
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE r.Id = @Id
          AND s.NegocioId = @NegocioId
          AND r.Estado NOT IN (5, 6);

        IF @EspacioDeportivoId IS NULL
            RETURN;

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
            RAISERROR('Cruce de horario con otra reserva.', 16, 1);

        IF EXISTS (
            SELECT 1
            FROM dbo.BloqueosHorario b
            WHERE b.EspacioDeportivoId = @EspacioDeportivoId
              AND b.Fecha = @Fecha
              AND b.Activo = 1
              AND @HoraInicio < b.HoraFin
              AND @HoraFin > b.HoraInicio
        )
            RAISERROR('El horario esta bloqueado para ese espacio.', 16, 1);

        UPDATE dbo.Reservas
        SET Fecha = @Fecha,
            HoraInicio = @HoraInicio,
            HoraFin = @HoraFin,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id;

        IF @@ROWCOUNT > 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
            EXEC dbo.Sp_Auditoria_Registrar
                @NegocioId = @NegocioId,
                @Modulo = N'RESERVAS',
                @Accion = N'MOVE',
                @Entidad = N'Reserva',
                @EntidadId = @EntidadIdAudit,
                @Usuario = @Usuario,
                @DetalleJson = NULL;
        END;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Bloqueos_Listar
    @NegocioId INT,
    @FechaDesde DATE,
    @FechaHasta DATE,
    @SedeId INT = NULL,
    @EspacioDeportivoId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            b.Id,
            s.Nombre AS Sede,
            e.Nombre AS Espacio,
            b.Fecha,
            b.HoraInicio,
            b.HoraFin,
            b.Motivo,
            b.Activo
        FROM dbo.BloqueosHorario b
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = b.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND b.Activo = 1
          AND b.Fecha BETWEEN @FechaDesde AND @FechaHasta
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND (@EspacioDeportivoId IS NULL OR e.Id = @EspacioDeportivoId)
        ORDER BY b.Fecha, b.HoraInicio;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Bloqueos_Crear
    @NegocioId INT,
    @EspacioDeportivoId INT,
    @Fecha DATE,
    @HoraInicio TIME,
    @HoraFin TIME,
    @Motivo NVARCHAR(250),
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @HoraFin <= @HoraInicio
            RAISERROR('La hora fin debe ser mayor a la hora inicio.', 16, 1);

        IF NOT EXISTS (
            SELECT 1
            FROM dbo.EspaciosDeportivos e
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
            WHERE e.Id = @EspacioDeportivoId
              AND s.NegocioId = @NegocioId
        )
            RAISERROR('Espacio no valido para el negocio.', 16, 1);

        IF EXISTS (
            SELECT 1
            FROM dbo.BloqueosHorario b
            WHERE b.EspacioDeportivoId = @EspacioDeportivoId
              AND b.Fecha = @Fecha
              AND b.Activo = 1
              AND @HoraInicio < b.HoraFin
              AND @HoraFin > b.HoraInicio
        )
            RAISERROR('Ya existe un bloqueo que se cruza con ese horario.', 16, 1);

        IF EXISTS (
            SELECT 1
            FROM dbo.Reservas r
            WHERE r.EspacioDeportivoId = @EspacioDeportivoId
              AND r.Fecha = @Fecha
              AND r.Estado NOT IN (5, 6)
              AND @HoraInicio < r.HoraFin
              AND @HoraFin > r.HoraInicio
        )
            RAISERROR('Existe una reserva en ese horario, no se puede bloquear.', 16, 1);

        INSERT INTO dbo.BloqueosHorario
        (
            EspacioDeportivoId, Fecha, HoraInicio, HoraFin, Motivo,
            Activo, FechaCreacion, UsuarioCreacion
        )
        VALUES
        (
            @EspacioDeportivoId, @Fecha, @HoraInicio, @HoraFin, @Motivo,
            1, SYSUTCDATETIME(), @Usuario
        );

        DECLARE @Id INT;
        SET @Id = SCOPE_IDENTITY();
        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar
            @NegocioId = @NegocioId,
            @Modulo = N'RESERVAS',
            @Accion = N'BLOCK',
            @Entidad = N'BloqueoHorario',
            @EntidadId = @EntidadIdAudit,
            @Usuario = @Usuario,
            @DetalleJson = NULL;

        SELECT @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Bloqueos_Eliminar
    @NegocioId INT,
    @Id INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE b
        SET b.Activo = 0,
            b.FechaActualizacion = SYSUTCDATETIME(),
            b.UsuarioActualizacion = @Usuario
        FROM dbo.BloqueosHorario b
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = b.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE b.Id = @Id
          AND s.NegocioId = @NegocioId;

        IF @@ROWCOUNT > 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
            EXEC dbo.Sp_Auditoria_Registrar
                @NegocioId = @NegocioId,
                @Modulo = N'RESERVAS',
                @Accion = N'UNBLOCK',
                @Entidad = N'BloqueoHorario',
                @EntidadId = @EntidadIdAudit,
                @Usuario = @Usuario,
                @DetalleJson = NULL;
        END;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
