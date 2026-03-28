-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Catalogos de deportes y tipos de suelo para espacios deportivos + actualizacion de SPs.
-- Firma:         Codex - 27/03/2026
-- =============================================

IF OBJECT_ID(N'dbo.TiposDeporte', N'U') IS NOT NULL
BEGIN
    IF NOT EXISTS (SELECT 1 FROM dbo.TiposDeporte WHERE UPPER(Nombre) = N'FUTBOL')
        INSERT INTO dbo.TiposDeporte (Nombre, Activo) VALUES (N'Futbol', 1);
    IF NOT EXISTS (SELECT 1 FROM dbo.TiposDeporte WHERE UPPER(Nombre) = N'VOLEY')
        INSERT INTO dbo.TiposDeporte (Nombre, Activo) VALUES (N'Voley', 1);
    IF NOT EXISTS (SELECT 1 FROM dbo.TiposDeporte WHERE UPPER(Nombre) = N'BASKET')
        INSERT INTO dbo.TiposDeporte (Nombre, Activo) VALUES (N'Basket', 1);
    IF NOT EXISTS (SELECT 1 FROM dbo.TiposDeporte WHERE UPPER(Nombre) = N'FRONTON')
        INSERT INTO dbo.TiposDeporte (Nombre, Activo) VALUES (N'Fronton', 1);
    IF NOT EXISTS (SELECT 1 FROM dbo.TiposDeporte WHERE UPPER(Nombre) = N'PADEL')
        INSERT INTO dbo.TiposDeporte (Nombre, Activo) VALUES (N'Padel', 1);
END;
GO

IF OBJECT_ID(N'dbo.TiposSuelo', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.TiposSuelo
    (
        Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        Nombre NVARCHAR(80) NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_TiposSuelo_Activo DEFAULT (1)
    );
END;
GO

IF NOT EXISTS (SELECT 1 FROM dbo.TiposSuelo WHERE UPPER(Nombre) = N'GRASS SINTETICO')
    INSERT INTO dbo.TiposSuelo (Nombre, Activo) VALUES (N'Grass Sintetico', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.TiposSuelo WHERE UPPER(Nombre) = N'GRASS NATURAL')
    INSERT INTO dbo.TiposSuelo (Nombre, Activo) VALUES (N'Grass Natural', 1);
IF NOT EXISTS (SELECT 1 FROM dbo.TiposSuelo WHERE UPPER(Nombre) = N'LOSA')
    INSERT INTO dbo.TiposSuelo (Nombre, Activo) VALUES (N'Losa', 1);
GO

IF COL_LENGTH('dbo.EspaciosDeportivos', 'TipoSueloId') IS NULL
BEGIN
    ALTER TABLE dbo.EspaciosDeportivos ADD TipoSueloId INT NULL;
END;
GO

DECLARE @TipoSueloDefaultId INT;
SELECT TOP 1 @TipoSueloDefaultId = ts.Id
FROM dbo.TiposSuelo ts
WHERE ts.Activo = 1
ORDER BY CASE WHEN UPPER(ts.Nombre) = N'LOSA' THEN 0 ELSE 1 END, ts.Id;

IF @TipoSueloDefaultId IS NOT NULL
BEGIN
    UPDATE dbo.EspaciosDeportivos
    SET TipoSueloId = @TipoSueloDefaultId
    WHERE TipoSueloId IS NULL;
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_EspaciosDeportivos_TiposSuelo_TipoSueloId')
BEGIN
    ALTER TABLE dbo.EspaciosDeportivos
    ADD CONSTRAINT FK_EspaciosDeportivos_TiposSuelo_TipoSueloId
        FOREIGN KEY (TipoSueloId) REFERENCES dbo.TiposSuelo (Id);
END;
GO

IF COL_LENGTH('dbo.EspaciosDeportivos', 'TipoSueloId') IS NOT NULL
BEGIN
    IF EXISTS (SELECT 1 FROM dbo.EspaciosDeportivos WHERE TipoSueloId IS NULL)
    BEGIN
        RAISERROR('Existen espacios sin TipoSueloId. Revisa datos antes de continuar.', 16, 1);
        RETURN;
    END;

    DECLARE @isNullable INT;
    SELECT @isNullable = c.is_nullable
    FROM sys.columns c
    WHERE c.object_id = OBJECT_ID(N'dbo.EspaciosDeportivos')
      AND c.name = N'TipoSueloId';

    IF @isNullable = 1
        ALTER TABLE dbo.EspaciosDeportivos ALTER COLUMN TipoSueloId INT NOT NULL;
END;
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

CREATE OR ALTER PROCEDURE dbo.Sp_Combos_TiposSuelo
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT ts.Id, ts.Nombre
        FROM dbo.TiposSuelo ts
        WHERE ts.Activo = 1
        ORDER BY ts.Nombre;
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
            ts.Nombre AS TipoSuelo,
            CASE e.Estado WHEN 1 THEN N'Activo' WHEN 2 THEN N'EnMantenimiento' ELSE N'Inactivo' END AS Estado
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.TiposDeporte td ON td.Id = e.TipoDeporteId
        INNER JOIN dbo.TiposSuelo ts ON ts.Id = e.TipoSueloId
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
            e.TipoSueloId,
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
    @TipoSueloId INT,
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
        IF NOT EXISTS (SELECT 1 FROM dbo.TiposDeporte WHERE Id = @TipoDeporteId AND Activo = 1)
            RAISERROR('Tipo de deporte invalido.', 16, 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.TiposSuelo WHERE Id = @TipoSueloId AND Activo = 1)
            RAISERROR('Tipo de suelo invalido.', 16, 1);

        INSERT INTO dbo.EspaciosDeportivos
        (
            SedeId, TipoDeporteId, TipoSueloId, Codigo, Nombre, Capacidad,
            TieneIluminacion, Techada, Estado, FechaCreacion, UsuarioCreacion
        )
        VALUES
        (
            @SedeId, @TipoDeporteId, @TipoSueloId, @Codigo, @Nombre, @Capacidad,
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
    @TipoSueloId INT,
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
        IF NOT EXISTS (SELECT 1 FROM dbo.TiposDeporte WHERE Id = @TipoDeporteId AND Activo = 1)
            RAISERROR('Tipo de deporte invalido.', 16, 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.TiposSuelo WHERE Id = @TipoSueloId AND Activo = 1)
            RAISERROR('Tipo de suelo invalido.', 16, 1);

        UPDATE e
        SET
            e.SedeId = @SedeId,
            e.TipoDeporteId = @TipoDeporteId,
            e.TipoSueloId = @TipoSueloId,
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
