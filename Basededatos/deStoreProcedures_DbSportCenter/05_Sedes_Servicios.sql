-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Catalogo de servicios de sede y ajuste de SP de sedes para seleccion multiple.
-- =============================================

IF OBJECT_ID(N'dbo.CatalogoServiciosSede', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CatalogoServiciosSede
    (
        Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        Nombre NVARCHAR(120) NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_CatalogoServiciosSede_Activo DEFAULT (1)
    );
END;
GO

IF OBJECT_ID(N'dbo.SedeServicios', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SedeServicios
    (
        SedeId INT NOT NULL,
        ServicioId INT NOT NULL,
        FechaRegistro DATETIME2 NOT NULL CONSTRAINT DF_SedeServicios_FechaRegistro DEFAULT (SYSUTCDATETIME()),
        UsuarioCreacion NVARCHAR(200) NULL,
        CONSTRAINT PK_SedeServicios PRIMARY KEY (SedeId, ServicioId),
        CONSTRAINT FK_SedeServicios_Sedes_SedeId FOREIGN KEY (SedeId) REFERENCES dbo.Sedes (Id),
        CONSTRAINT FK_SedeServicios_CatalogoServiciosSede_ServicioId FOREIGN KEY (ServicioId) REFERENCES dbo.CatalogoServiciosSede (Id)
    );
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.SedeServicios') AND name = N'IX_SedeServicios_ServicioId')
BEGIN
    CREATE INDEX IX_SedeServicios_ServicioId ON dbo.SedeServicios (ServicioId);
END;
GO

IF NOT EXISTS (SELECT 1 FROM dbo.CatalogoServiciosSede WHERE Nombre = N'Wi-Fi')
    INSERT INTO dbo.CatalogoServiciosSede (Nombre, Activo) VALUES (N'Wi-Fi', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.CatalogoServiciosSede WHERE Nombre = N'Vestuario')
    INSERT INTO dbo.CatalogoServiciosSede (Nombre, Activo) VALUES (N'Vestuario', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.CatalogoServiciosSede WHERE Nombre = N'Estacionamiento')
    INSERT INTO dbo.CatalogoServiciosSede (Nombre, Activo) VALUES (N'Estacionamiento', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.CatalogoServiciosSede WHERE Nombre = N'Ayuda Medica')
    INSERT INTO dbo.CatalogoServiciosSede (Nombre, Activo) VALUES (N'Ayuda Medica', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.CatalogoServiciosSede WHERE Nombre = N'Torneos')
    INSERT INTO dbo.CatalogoServiciosSede (Nombre, Activo) VALUES (N'Torneos', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.CatalogoServiciosSede WHERE Nombre = N'Cumpleanos')
    INSERT INTO dbo.CatalogoServiciosSede (Nombre, Activo) VALUES (N'Cumpleanos', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.CatalogoServiciosSede WHERE Nombre = N'Escuelita deportiva')
    INSERT INTO dbo.CatalogoServiciosSede (Nombre, Activo) VALUES (N'Escuelita deportiva', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.CatalogoServiciosSede WHERE Nombre = N'Bar / Restaurante')
    INSERT INTO dbo.CatalogoServiciosSede (Nombre, Activo) VALUES (N'Bar / Restaurante', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.CatalogoServiciosSede WHERE Nombre = N'Tienda Deportiva')
    INSERT INTO dbo.CatalogoServiciosSede (Nombre, Activo) VALUES (N'Tienda Deportiva', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.CatalogoServiciosSede WHERE Nombre = N'Beelup')
    INSERT INTO dbo.CatalogoServiciosSede (Nombre, Activo) VALUES (N'Beelup', 1);
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Combos_ServiciosSede
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT cs.Id, cs.Nombre
        FROM dbo.CatalogoServiciosSede cs
        WHERE cs.Activo = 1
        ORDER BY cs.Nombre;
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
        SELECT
            s.Id,
            s.Nombre,
            s.Direccion,
            STUFF((
                SELECT N', ' + cs.Nombre
                FROM dbo.SedeServicios ss
                INNER JOIN dbo.CatalogoServiciosSede cs ON cs.Id = ss.ServicioId
                WHERE ss.SedeId = s.Id
                  AND cs.Activo = 1
                ORDER BY cs.Nombre
                FOR XML PATH(''), TYPE
            ).value('.', 'NVARCHAR(MAX)'), 1, 2, N'') AS Servicios,
            s.Activo
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
        SELECT
            s.Id,
            s.NegocioId,
            s.Nombre,
            s.Direccion,
            s.Telefono,
            s.Activo,
            STUFF((
                SELECT N',' + CONVERT(NVARCHAR(20), ss.ServicioId)
                FROM dbo.SedeServicios ss
                WHERE ss.SedeId = s.Id
                ORDER BY ss.ServicioId
                FOR XML PATH(''), TYPE
            ).value('.', 'NVARCHAR(MAX)'), 1, 1, N'') AS ServiciosIdsCsv
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
    @ServiciosIdsCsv NVARCHAR(MAX) = NULL,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        BEGIN TRANSACTION;

        INSERT INTO dbo.Sedes (NegocioId, Nombre, Direccion, Telefono, Activo, FechaCreacion, UsuarioCreacion)
        VALUES (@NegocioId, @Nombre, @Direccion, @Telefono, @Activo, SYSUTCDATETIME(), @Usuario);

        DECLARE @Id INT;
        SET @Id = SCOPE_IDENTITY();

        IF @ServiciosIdsCsv IS NOT NULL AND LEN(LTRIM(RTRIM(@ServiciosIdsCsv))) > 0
        BEGIN
            ;WITH Servicios AS
            (
                SELECT DISTINCT TRY_CONVERT(INT, LTRIM(RTRIM(value))) AS ServicioId
                FROM STRING_SPLIT(@ServiciosIdsCsv, N',')
                WHERE TRY_CONVERT(INT, LTRIM(RTRIM(value))) IS NOT NULL
            )
            INSERT INTO dbo.SedeServicios (SedeId, ServicioId, FechaRegistro, UsuarioCreacion)
            SELECT @Id, s.ServicioId, SYSUTCDATETIME(), @Usuario
            FROM Servicios s
            INNER JOIN dbo.CatalogoServiciosSede cs ON cs.Id = s.ServicioId
            WHERE cs.Activo = 1;
        END;

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'SEDES', @Accion = N'CREATE', @Entidad = N'Sede', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;

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

CREATE OR ALTER PROCEDURE dbo.Sp_Sedes_Actualizar
    @Id INT,
    @NegocioId INT,
    @Nombre NVARCHAR(150),
    @Direccion NVARCHAR(250),
    @Telefono NVARCHAR(20) = NULL,
    @Activo BIT,
    @ServiciosIdsCsv NVARCHAR(MAX) = NULL,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        BEGIN TRANSACTION;

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
        BEGIN
            ROLLBACK TRANSACTION;
            RETURN;
        END;

        DELETE ss
        FROM dbo.SedeServicios ss
        WHERE ss.SedeId = @Id;

        IF @ServiciosIdsCsv IS NOT NULL AND LEN(LTRIM(RTRIM(@ServiciosIdsCsv))) > 0
        BEGIN
            ;WITH Servicios AS
            (
                SELECT DISTINCT TRY_CONVERT(INT, LTRIM(RTRIM(value))) AS ServicioId
                FROM STRING_SPLIT(@ServiciosIdsCsv, N',')
                WHERE TRY_CONVERT(INT, LTRIM(RTRIM(value))) IS NOT NULL
            )
            INSERT INTO dbo.SedeServicios (SedeId, ServicioId, FechaRegistro, UsuarioCreacion)
            SELECT @Id, s.ServicioId, SYSUTCDATETIME(), @Usuario
            FROM Servicios s
            INNER JOIN dbo.CatalogoServiciosSede cs ON cs.Id = s.ServicioId
            WHERE cs.Activo = 1;
        END;

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'SEDES', @Accion = N'EDIT', @Entidad = N'Sede', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;

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
