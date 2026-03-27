-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Gestion interna de solicitudes publicas y conversion a reserva.
-- =============================================

IF COL_LENGTH('dbo.SolicitudesReservaPublica', 'ReservaId') IS NULL
BEGIN
    ALTER TABLE dbo.SolicitudesReservaPublica ADD ReservaId INT NULL;
END;
GO

IF COL_LENGTH('dbo.SolicitudesReservaPublica', 'FechaGestion') IS NULL
BEGIN
    ALTER TABLE dbo.SolicitudesReservaPublica ADD FechaGestion DATETIME2 NULL;
END;
GO

IF COL_LENGTH('dbo.SolicitudesReservaPublica', 'UsuarioGestion') IS NULL
BEGIN
    ALTER TABLE dbo.SolicitudesReservaPublica ADD UsuarioGestion NVARCHAR(200) NULL;
END;
GO

IF COL_LENGTH('dbo.SolicitudesReservaPublica', 'ComentarioGestion') IS NULL
BEGIN
    ALTER TABLE dbo.SolicitudesReservaPublica ADD ComentarioGestion NVARCHAR(300) NULL;
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_SolicitudesReservaPublica_Reservas_ReservaId')
BEGIN
    ALTER TABLE dbo.SolicitudesReservaPublica
    ADD CONSTRAINT FK_SolicitudesReservaPublica_Reservas_ReservaId
    FOREIGN KEY (ReservaId) REFERENCES dbo.Reservas (Id);
END;
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Seguridad_SeedModulosPermisosBase
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'DASHBOARD')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'DASHBOARD', N'Dashboard', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'SEDES')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'SEDES', N'Sedes', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'CLIENTES')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'CLIENTES', N'Clientes', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'ESPACIOS')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'ESPACIOS', N'Espacios deportivos', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'RESERVAS')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'RESERVAS', N'Reservas', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'SOLICITUDES')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'SOLICITUDES', N'Solicitudes publicas', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'PAGOS')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'PAGOS', N'Pagos', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'COMPROBANTES')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'COMPROBANTES', N'Comprobantes electronicos', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'REPORTES')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'REPORTES', N'Reportes', 1);

        ;WITH Roles AS
        (
            SELECT CAST(1 AS INT) AS RolNegocio UNION ALL
            SELECT 2 UNION ALL
            SELECT 3 UNION ALL
            SELECT 4 UNION ALL
            SELECT 5
        )
        INSERT INTO dbo.RolesNegocioPermiso (RolNegocio, ModuloSistemaId, PuedeVer, PuedeCrear, PuedeEditar, PuedeEliminar)
        SELECT
            r.RolNegocio,
            m.Id,
            CAST(CASE WHEN r.RolNegocio = 1 THEN 1
                      WHEN r.RolNegocio = 2 AND m.Codigo IN (N'DASHBOARD', N'RESERVAS', N'SOLICITUDES', N'PAGOS', N'CLIENTES', N'REPORTES') THEN 1
                      WHEN r.RolNegocio = 3 AND m.Codigo IN (N'DASHBOARD', N'RESERVAS', N'SOLICITUDES', N'CLIENTES') THEN 1
                      WHEN r.RolNegocio = 4 AND m.Codigo IN (N'DASHBOARD', N'PAGOS', N'COMPROBANTES', N'REPORTES') THEN 1
                      WHEN r.RolNegocio = 5 AND m.Codigo IN (N'DASHBOARD', N'RESERVAS', N'ESPACIOS', N'REPORTES') THEN 1
                      ELSE 0 END AS BIT),
            CAST(CASE WHEN r.RolNegocio = 1 THEN 1
                      WHEN r.RolNegocio = 2 AND m.Codigo IN (N'RESERVAS', N'SOLICITUDES', N'PAGOS', N'CLIENTES') THEN 1
                      WHEN r.RolNegocio = 3 AND m.Codigo IN (N'RESERVAS', N'SOLICITUDES', N'CLIENTES') THEN 1
                      WHEN r.RolNegocio = 4 AND m.Codigo IN (N'PAGOS', N'COMPROBANTES') THEN 1
                      ELSE 0 END AS BIT),
            CAST(CASE WHEN r.RolNegocio = 1 THEN 1
                      WHEN r.RolNegocio = 2 AND m.Codigo IN (N'RESERVAS', N'SOLICITUDES', N'PAGOS', N'CLIENTES') THEN 1
                      WHEN r.RolNegocio = 3 AND m.Codigo IN (N'RESERVAS', N'SOLICITUDES', N'CLIENTES') THEN 1
                      WHEN r.RolNegocio = 4 AND m.Codigo IN (N'PAGOS', N'COMPROBANTES') THEN 1
                      WHEN r.RolNegocio = 5 AND m.Codigo IN (N'RESERVAS', N'ESPACIOS') THEN 1
                      ELSE 0 END AS BIT),
            CAST(CASE WHEN r.RolNegocio = 1 THEN 1 ELSE 0 END AS BIT)
        FROM Roles r
        CROSS JOIN dbo.ModulosSistema m
        WHERE NOT EXISTS (
            SELECT 1
            FROM dbo.RolesNegocioPermiso rp
            WHERE rp.RolNegocio = r.RolNegocio
              AND rp.ModuloSistemaId = m.Id
        );
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_SolicitudesPublicas_Listar
    @NegocioId INT,
    @FechaDesde DATE = NULL,
    @FechaHasta DATE = NULL,
    @Estado INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            s.Id,
            s.CodigoSolicitud,
            se.Nombre AS Sede,
            e.Nombre AS Espacio,
            s.Fecha,
            s.HoraInicio,
            s.HoraFin,
            s.NombreSolicitante,
            s.Telefono,
            s.Correo,
            s.Estado,
            s.ReservaId,
            s.FechaRegistro
        FROM dbo.SolicitudesReservaPublica s
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = s.EspacioDeportivoId
        INNER JOIN dbo.Sedes se ON se.Id = e.SedeId
        WHERE se.NegocioId = @NegocioId
          AND (@FechaDesde IS NULL OR s.Fecha >= @FechaDesde)
          AND (@FechaHasta IS NULL OR s.Fecha <= @FechaHasta)
          AND (@Estado IS NULL OR s.Estado = @Estado)
        ORDER BY s.FechaRegistro DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_SolicitudesPublicas_ActualizarEstado
    @NegocioId INT,
    @Id INT,
    @Estado INT,
    @ComentarioGestion NVARCHAR(300) = NULL,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @Estado NOT IN (2, 3)
            RAISERROR('Estado invalido. Solo se permite aprobar(2) o rechazar(3).', 16, 1);

        UPDATE s
        SET s.Estado = @Estado,
            s.ComentarioGestion = @ComentarioGestion,
            s.FechaGestion = SYSUTCDATETIME(),
            s.UsuarioGestion = @Usuario
        FROM dbo.SolicitudesReservaPublica s
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = s.EspacioDeportivoId
        INNER JOIN dbo.Sedes se ON se.Id = e.SedeId
        WHERE s.Id = @Id
          AND se.NegocioId = @NegocioId
          AND s.Estado IN (1, 2, 3);

        IF @@ROWCOUNT > 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
            EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'SOLICITUDES', @Accion = N'EDIT', @Entidad = N'SolicitudReservaPublica', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        END;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_SolicitudesPublicas_ConvertirAReserva
    @NegocioId INT,
    @Id INT,
    @Total DECIMAL(10,2),
    @Adelanto DECIMAL(10,2),
    @EstadoReserva INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @Total < 0 OR @Adelanto < 0 OR @Adelanto > @Total
            RAISERROR('Montos invalidos para la conversion.', 16, 1);

        DECLARE @EspacioDeportivoId INT, @Fecha DATE, @HoraInicio TIME, @HoraFin TIME, @NombreSolicitante NVARCHAR(200), @Telefono NVARCHAR(30), @Correo NVARCHAR(200);
        DECLARE @ClienteId INT, @ReservaId INT;

        SELECT
            @EspacioDeportivoId = s.EspacioDeportivoId,
            @Fecha = s.Fecha,
            @HoraInicio = s.HoraInicio,
            @HoraFin = s.HoraFin,
            @NombreSolicitante = s.NombreSolicitante,
            @Telefono = s.Telefono,
            @Correo = s.Correo
        FROM dbo.SolicitudesReservaPublica s
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = s.EspacioDeportivoId
        INNER JOIN dbo.Sedes se ON se.Id = e.SedeId
        WHERE s.Id = @Id
          AND se.NegocioId = @NegocioId
          AND s.Estado IN (1, 2);

        IF @EspacioDeportivoId IS NULL
            RAISERROR('Solicitud invalida para el negocio.', 16, 1);

        IF EXISTS (
            SELECT 1
            FROM dbo.Reservas r
            WHERE r.EspacioDeportivoId = @EspacioDeportivoId
              AND r.Fecha = @Fecha
              AND r.Estado NOT IN (5, 6)
              AND @HoraInicio < r.HoraFin
              AND @HoraFin > r.HoraInicio
        )
            RAISERROR('No se puede convertir: el horario ya fue tomado.', 16, 1);

        SELECT TOP (1) @ClienteId = c.Id
        FROM dbo.Clientes c
        INNER JOIN dbo.NegocioClientes nc ON nc.ClienteId = c.Id
        WHERE nc.NegocioId = @NegocioId
          AND nc.Activo = 1
          AND c.Activo = 1
          AND c.NombresORazonSocial = @NombreSolicitante
          AND c.Telefono = @Telefono;

        BEGIN TRANSACTION;

        IF @ClienteId IS NULL
        BEGIN
            INSERT INTO dbo.Clientes
            (
                NombresORazonSocial, TipoDocumento, NumeroDocumento, Telefono, Correo,
                Activo, FechaCreacion, UsuarioCreacion
            )
            VALUES
            (
                @NombreSolicitante, N'OTRO', CONCAT(N'SOL', @Id), @Telefono, @Correo,
                1, SYSUTCDATETIME(), @Usuario
            );

            SET @ClienteId = SCOPE_IDENTITY();

            INSERT INTO dbo.NegocioClientes (NegocioId, ClienteId, Activo, FechaRegistro, UsuarioCreacion)
            VALUES (@NegocioId, @ClienteId, 1, SYSUTCDATETIME(), @Usuario);
        END;

        INSERT INTO dbo.Reservas
        (
            EspacioDeportivoId, ClienteId, Fecha, HoraInicio, HoraFin,
            Estado, Total, Adelanto, Saldo, FechaRegistro, UsuarioCreacion
        )
        VALUES
        (
            @EspacioDeportivoId, @ClienteId, @Fecha, @HoraInicio, @HoraFin,
            @EstadoReserva, @Total, @Adelanto, (@Total - @Adelanto), SYSUTCDATETIME(), @Usuario
        );

        SET @ReservaId = SCOPE_IDENTITY();

        UPDATE dbo.SolicitudesReservaPublica
        SET Estado = 4,
            ReservaId = @ReservaId,
            FechaGestion = SYSUTCDATETIME(),
            UsuarioGestion = @Usuario,
            ComentarioGestion = N'Convertida a reserva'
        WHERE Id = @Id;

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'SOLICITUDES', @Accion = N'EDIT', @Entidad = N'SolicitudReservaPublica', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;

        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @ReservaId);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'RESERVAS', @Accion = N'CREATE', @Entidad = N'Reserva', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;

        COMMIT TRANSACTION;

        SELECT @ReservaId;
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
