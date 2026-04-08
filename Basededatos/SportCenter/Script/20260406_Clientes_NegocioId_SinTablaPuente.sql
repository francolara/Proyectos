USE [DbSportCenter]
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/04/2026
-- Description:   Migra clientes a modelo directo por NegocioId y elimina tabla puente NegocioClientes.
-- Firma:         Codex - 06/04/2026
-- =============================================

IF COL_LENGTH('dbo.Clientes', 'NegocioId') IS NULL
BEGIN
    ALTER TABLE dbo.Clientes ADD NegocioId INT NULL;
END;

IF OBJECT_ID('dbo.NegocioClientes', 'U') IS NOT NULL
BEGIN
    ;WITH Priorizado AS
    (
        SELECT
            nc.ClienteId,
            nc.NegocioId,
            ROW_NUMBER() OVER (
                PARTITION BY nc.ClienteId
                ORDER BY
                    CASE WHEN nc.Activo = 1 THEN 0 ELSE 1 END,
                    nc.FechaRegistro DESC,
                    nc.NegocioId ASC
            ) AS rn
        FROM dbo.NegocioClientes nc
    )
    UPDATE c
    SET c.NegocioId = p.NegocioId
    FROM dbo.Clientes c
    INNER JOIN Priorizado p ON p.ClienteId = c.Id AND p.rn = 1
    WHERE c.NegocioId IS NULL;
END;

IF EXISTS (SELECT 1 FROM dbo.Clientes WHERE NegocioId IS NULL)
BEGIN
    RAISERROR('Hay clientes sin NegocioId luego de la migracion. Revisar datos antes de continuar.', 16, 1);
    RETURN;
END;

IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = 'FK_Clientes_Negocios_NegocioId' AND parent_object_id = OBJECT_ID('dbo.Clientes'))
BEGIN
    ALTER TABLE dbo.Clientes WITH CHECK
    ADD CONSTRAINT FK_Clientes_Negocios_NegocioId
    FOREIGN KEY (NegocioId) REFERENCES dbo.Negocios(Id);
END;

IF EXISTS (
    SELECT 1
    FROM sys.columns
    WHERE object_id = OBJECT_ID('dbo.Clientes')
      AND name = 'NegocioId'
      AND is_nullable = 1
)
BEGIN
    ALTER TABLE dbo.Clientes ALTER COLUMN NegocioId INT NOT NULL;
END;

IF EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID('dbo.Clientes') AND name = 'IX_Clientes_TipoDocumento_NumeroDocumento')
BEGIN
    DROP INDEX IX_Clientes_TipoDocumento_NumeroDocumento ON dbo.Clientes;
END;

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID('dbo.Clientes') AND name = 'UX_Clientes_Negocio_Tipo_Numero_Activo')
BEGIN
    CREATE UNIQUE NONCLUSTERED INDEX UX_Clientes_Negocio_Tipo_Numero_Activo
    ON dbo.Clientes (NegocioId, TipoDocumento, NumeroDocumento)
    WHERE Activo = 1;
END;

IF OBJECT_ID('dbo.NegocioClientes', 'U') IS NOT NULL
BEGIN
    DROP TABLE dbo.NegocioClientes;
END;
