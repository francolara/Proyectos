-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Configuracion del club y moneda de trabajo (Soles/Dolares) para panel de Configuracion. Incluye ajuste de compatibilidad en llamada de auditoria.
-- Firma:         Codex - 27/03/2026
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   28/03/2026
-- Description:   Se agregan TipoDocumentoFiscal, NumeroDocumentoFiscal y DireccionFiscal para configuracion del club.
-- Firma:         Codex - 28/03/2026
-- =============================================

IF OBJECT_ID(N'dbo.Monedas', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Monedas
    (
        Id INT IDENTITY(1,1) NOT NULL,
        Codigo NVARCHAR(10) NOT NULL,
        Nombre NVARCHAR(80) NOT NULL,
        Simbolo NVARCHAR(10) NULL,
        Activo BIT NOT NULL CONSTRAINT DF_Monedas_Activo DEFAULT (1),
        FechaCreacion DATETIME2 NOT NULL CONSTRAINT DF_Monedas_FechaCreacion DEFAULT (SYSUTCDATETIME()),
        CONSTRAINT PK_Monedas PRIMARY KEY CLUSTERED (Id),
        CONSTRAINT UQ_Monedas_Codigo UNIQUE (Codigo)
    );
END;
GO

IF NOT EXISTS (SELECT 1 FROM dbo.Monedas WHERE Codigo = N'PEN')
BEGIN
    INSERT INTO dbo.Monedas (Codigo, Nombre, Simbolo, Activo)
    VALUES (N'PEN', N'Soles', N'S/', 1);
END;
GO

IF NOT EXISTS (SELECT 1 FROM dbo.Monedas WHERE Codigo = N'USD')
BEGIN
    INSERT INTO dbo.Monedas (Codigo, Nombre, Simbolo, Activo)
    VALUES (N'USD', N'Dolares', N'$', 1);
END;
GO

IF COL_LENGTH('dbo.Negocios', 'MonedaId') IS NULL
BEGIN
    ALTER TABLE dbo.Negocios
        ADD MonedaId INT NULL;
END;
GO

IF COL_LENGTH('dbo.Negocios', 'TipoDocumentoFiscal') IS NULL
BEGIN
    ALTER TABLE dbo.Negocios
        ADD TipoDocumentoFiscal NVARCHAR(20) NULL;
END;
GO

IF COL_LENGTH('dbo.Negocios', 'NumeroDocumentoFiscal') IS NULL
BEGIN
    ALTER TABLE dbo.Negocios
        ADD NumeroDocumentoFiscal NVARCHAR(20) NULL;
END;
GO

IF COL_LENGTH('dbo.Negocios', 'DireccionFiscal') IS NULL
BEGIN
    ALTER TABLE dbo.Negocios
        ADD DireccionFiscal NVARCHAR(250) NULL;
END;
GO

IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.Negocios') AND name = N'MonedaId')
BEGIN
    UPDATE n
    SET n.MonedaId = 1
    FROM dbo.Negocios n
    WHERE n.MonedaId IS NULL;
END;
GO

IF COL_LENGTH('dbo.Negocios', 'DocumentoFiscal') IS NOT NULL
BEGIN
    UPDATE n
    SET n.NumeroDocumentoFiscal = n.DocumentoFiscal
    FROM dbo.Negocios n
    WHERE n.NumeroDocumentoFiscal IS NULL
      AND n.DocumentoFiscal IS NOT NULL;
END;
GO

IF NOT EXISTS
(
    SELECT 1
    FROM sys.default_constraints dc
    INNER JOIN sys.columns c ON c.default_object_id = dc.object_id
    WHERE dc.parent_object_id = OBJECT_ID(N'dbo.Negocios')
      AND c.name = N'MonedaId'
)
BEGIN
    ALTER TABLE dbo.Negocios
        ADD CONSTRAINT DF_Negocios_MonedaId DEFAULT (1) FOR MonedaId;
END;
GO

IF NOT EXISTS
(
    SELECT 1
    FROM sys.foreign_keys
    WHERE name = N'FK_Negocios_Monedas_MonedaId'
      AND parent_object_id = OBJECT_ID(N'dbo.Negocios')
)
BEGIN
    ALTER TABLE dbo.Negocios WITH CHECK
        ADD CONSTRAINT FK_Negocios_Monedas_MonedaId
            FOREIGN KEY (MonedaId) REFERENCES dbo.Monedas (Id);
END;
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Combos_Monedas
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            m.Id,
            CONCAT(m.Nombre, N' (', m.Codigo, N')') AS Nombre
        FROM dbo.Monedas m
        WHERE m.Activo = 1
        ORDER BY m.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_ConfiguracionClub_Obtener
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            n.Id,
            n.NombreComercial,
            n.RazonSocial,
            COALESCE(NULLIF(n.TipoDocumentoFiscal, N''), N'DNI') AS TipoDocumentoFiscal,
            COALESCE(NULLIF(n.NumeroDocumentoFiscal, N''), n.DocumentoFiscal) AS NumeroDocumentoFiscal,
            n.DireccionFiscal,
            COALESCE(n.MonedaId, 1) AS MonedaId
        FROM dbo.Negocios n
        WHERE n.Id = @NegocioId
          AND n.Activo = 1;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_ConfiguracionClub_Actualizar
    @NegocioId INT,
    @NombreComercial NVARCHAR(200),
    @RazonSocial NVARCHAR(200) = NULL,
    @TipoDocumentoFiscal NVARCHAR(20) = NULL,
    @NumeroDocumentoFiscal NVARCHAR(20) = NULL,
    @DireccionFiscal NVARCHAR(250) = NULL,
    @MonedaId INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.Monedas WHERE Id = @MonedaId AND Activo = 1)
            RAISERROR('La moneda seleccionada no es valida.', 16, 1);

        UPDATE n
        SET
            n.NombreComercial = @NombreComercial,
            n.RazonSocial = NULLIF(@RazonSocial, N''),
            n.TipoDocumentoFiscal = NULLIF(@TipoDocumentoFiscal, N''),
            n.NumeroDocumentoFiscal = NULLIF(@NumeroDocumentoFiscal, N''),
            n.DireccionFiscal = NULLIF(@DireccionFiscal, N''),
            n.DocumentoFiscal = NULLIF(@NumeroDocumentoFiscal, N''),
            n.MonedaId = @MonedaId
        FROM dbo.Negocios n
        WHERE n.Id = @NegocioId
          AND n.Activo = 1;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el club para actualizar.', 16, 1);

        DECLARE @EntidadIdAuditoria NVARCHAR(80);
        SET @EntidadIdAuditoria = CONVERT(NVARCHAR(80), @NegocioId);

        EXEC dbo.Sp_Auditoria_Registrar
            @NegocioId = @NegocioId,
            @Modulo = N'CONFIGURACION',
            @Accion = N'EDIT',
            @Entidad = N'Negocio',
            @EntidadId = @EntidadIdAuditoria,
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
