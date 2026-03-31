-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/03/2026
-- Description:   CRUD de sedes, espacios y procedimientos de combos base.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Ajuste de llamadas de auditoria con parametros nombrados para evitar errores de sintaxis.
-- Firma:         Codex - 30/03/2026 | Sedes/Espacios eliminar ahora devuelve error cuando no existe registro para el negocio.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.Sp_Combos_Sedes
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT s.Id, s.Nombre
        FROM dbo.Sedes s
        WHERE s.NegocioId = @NegocioId
        ORDER BY s.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Combos_TiposDeporte
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT td.Id, td.Nombre
        FROM dbo.TiposDeporte td
        WHERE td.Activo = 1
        ORDER BY td.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Sedes_Listar
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT s.Id, s.Nombre, s.Direccion, s.Activo
        FROM dbo.Sedes s
        WHERE s.NegocioId = @NegocioId
        ORDER BY s.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Sedes_ObtenerPorId
    @NegocioId INT,
    @Id INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT s.Id, s.NegocioId, s.Nombre, s.Direccion, s.Telefono, s.Activo
        FROM dbo.Sedes s
        WHERE s.NegocioId = @NegocioId
          AND s.Id = @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Sedes_Crear
    @NegocioId INT,
    @Nombre NVARCHAR(150),
    @Direccion NVARCHAR(250),
    @Telefono NVARCHAR(20) = NULL,
    @Activo BIT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        INSERT INTO dbo.Sedes (NegocioId, Nombre, Direccion, Telefono, Activo, FechaCreacion, UsuarioCreacion)
        VALUES (@NegocioId, @Nombre, @Direccion, @Telefono, @Activo, SYSUTCDATETIME(), @Usuario);

        DECLARE @Id INT;
        SET @Id = SCOPE_IDENTITY();
        DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'SEDES', @Accion = N'CREATE', @Entidad = N'Sede', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        SELECT @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Sedes_Actualizar
    @Id INT,
    @NegocioId INT,
    @Nombre NVARCHAR(150),
    @Direccion NVARCHAR(250),
    @Telefono NVARCHAR(20) = NULL,
    @Activo BIT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE dbo.Sedes
        SET Nombre = @Nombre,
            Direccion = @Direccion,
            Telefono = @Telefono,
            Activo = @Activo,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id
          AND NegocioId = @NegocioId;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro la sede para eliminar.', 16, 1);

        IF @@ROWCOUNT > 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
            EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'SEDES', @Accion = N'EDIT', @Entidad = N'Sede', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        END
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Sedes_Eliminar
    @NegocioId INT,
    @Id INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE dbo.Sedes
        SET Activo = 0,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id
          AND NegocioId = @NegocioId;

        IF @@ROWCOUNT > 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
            EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'SEDES', @Accion = N'DELETE', @Entidad = N'Sede', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        END
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Espacios_Listar
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            e.Id,
            e.Codigo,
            e.Nombre,
            s.Nombre AS Sede,
            td.Nombre AS TipoDeporte,
            CASE e.Estado WHEN 1 THEN N'Activo' WHEN 2 THEN N'EnMantenimiento' ELSE N'Inactivo' END AS Estado
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.TiposDeporte td ON td.Id = e.TipoDeporteId
        WHERE s.NegocioId = @NegocioId
        ORDER BY s.Nombre, e.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Espacios_ObtenerPorId
    @NegocioId INT,
    @Id INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            e.Id,
            e.SedeId,
            e.TipoDeporteId,
            e.Codigo,
            e.Nombre,
            e.Capacidad,
            e.TieneIluminacion,
            e.Techada,
            e.Estado
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE e.Id = @Id
          AND s.NegocioId = @NegocioId;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Espacios_Crear
    @NegocioId INT,
    @SedeId INT,
    @TipoDeporteId INT,
    @Codigo NVARCHAR(20),
    @Nombre NVARCHAR(150),
    @Capacidad INT,
    @TieneIluminacion BIT,
    @Techada BIT,
    @Estado INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.Sedes WHERE Id = @SedeId AND NegocioId = @NegocioId)
            RAISERROR('Sede invalida para el negocio.', 16, 1);

        INSERT INTO dbo.EspaciosDeportivos
        (
            SedeId, TipoDeporteId, Codigo, Nombre, Capacidad,
            TieneIluminacion, Techada, Estado, FechaCreacion, UsuarioCreacion
        )
        VALUES
        (
            @SedeId, @TipoDeporteId, @Codigo, @Nombre, @Capacidad,
            @TieneIluminacion, @Techada, @Estado, SYSUTCDATETIME(), @Usuario
        );

        DECLARE @Id INT;
        SET @Id = SCOPE_IDENTITY();
        DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'ESPACIOS', @Accion = N'CREATE', @Entidad = N'EspacioDeportivo', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        SELECT @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Espacios_Actualizar
    @Id INT,
    @NegocioId INT,
    @SedeId INT,
    @TipoDeporteId INT,
    @Codigo NVARCHAR(20),
    @Nombre NVARCHAR(150),
    @Capacidad INT,
    @TieneIluminacion BIT,
    @Techada BIT,
    @Estado INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE e
        SET
            e.SedeId = @SedeId,
            e.TipoDeporteId = @TipoDeporteId,
            e.Codigo = @Codigo,
            e.Nombre = @Nombre,
            e.Capacidad = @Capacidad,
            e.TieneIluminacion = @TieneIluminacion,
            e.Techada = @Techada,
            e.Estado = @Estado,
            e.FechaActualizacion = SYSUTCDATETIME(),
            e.UsuarioActualizacion = @Usuario
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE e.Id = @Id
          AND s.NegocioId = @NegocioId;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el espacio deportivo para eliminar.', 16, 1);

        IF @@ROWCOUNT > 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
            EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'ESPACIOS', @Accion = N'EDIT', @Entidad = N'EspacioDeportivo', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        END
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Espacios_Eliminar
    @NegocioId INT,
    @Id INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE e
        SET
            e.Estado = 3,
            e.FechaActualizacion = SYSUTCDATETIME(),
            e.UsuarioActualizacion = @Usuario
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE e.Id = @Id
          AND s.NegocioId = @NegocioId;

        IF @@ROWCOUNT > 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
            EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'ESPACIOS', @Accion = N'DELETE', @Entidad = N'EspacioDeportivo', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        END
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
