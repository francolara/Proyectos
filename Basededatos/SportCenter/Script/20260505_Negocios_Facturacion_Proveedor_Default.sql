USE [DbSportCenter]
GO
-- Firma: Codex - 05/05/2026 | Agrega campos de configuracion base para proveedor electronico por defecto en Negocios.

IF COL_LENGTH('dbo.Negocios', 'ProveedorElectronicoDefaultId') IS NULL
BEGIN
    ALTER TABLE dbo.Negocios
    ADD ProveedorElectronicoDefaultId INT NULL;
END
GO

IF COL_LENGTH('dbo.Negocios', 'ModoEmisionElectronica') IS NULL
BEGIN
    ALTER TABLE dbo.Negocios
    ADD ModoEmisionElectronica NVARCHAR(15) NOT NULL
        CONSTRAINT DF_Negocios_ModoEmisionElectronica DEFAULT (N'PRODUCCION');
END
GO

IF COL_LENGTH('dbo.Negocios', 'EnviarComprobanteAutomatico') IS NULL
BEGIN
    ALTER TABLE dbo.Negocios
    ADD EnviarComprobanteAutomatico BIT NOT NULL
        CONSTRAINT DF_Negocios_EnviarComprobanteAutomatico DEFAULT ((0));
END
GO

IF COL_LENGTH('dbo.Negocios', 'UsaContingenciaFacturacion') IS NULL
BEGIN
    ALTER TABLE dbo.Negocios
    ADD UsaContingenciaFacturacion BIT NOT NULL
        CONSTRAINT DF_Negocios_UsaContingenciaFacturacion DEFAULT ((0));
END
GO

IF COL_LENGTH('dbo.Negocios', 'FechaUltimaSyncCatalogosFacturacion') IS NULL
BEGIN
    ALTER TABLE dbo.Negocios
    ADD FechaUltimaSyncCatalogosFacturacion DATETIME2(7) NULL;
END
GO

IF NOT EXISTS (
    SELECT 1
    FROM sys.check_constraints
    WHERE name = 'CK_Negocios_ModoEmisionElectronica'
      AND parent_object_id = OBJECT_ID('dbo.Negocios')
)
BEGIN
    ALTER TABLE dbo.Negocios WITH CHECK
    ADD CONSTRAINT CK_Negocios_ModoEmisionElectronica
    CHECK (ModoEmisionElectronica IN (N'BETA', N'PRODUCCION', N'AMBOS'));
END
GO

IF NOT EXISTS (
    SELECT 1
    FROM sys.foreign_keys
    WHERE name = 'FK_Negocios_FacturacionProveedores_ProveedorElectronicoDefaultId'
      AND parent_object_id = OBJECT_ID('dbo.Negocios')
)
BEGIN
    ALTER TABLE dbo.Negocios  WITH CHECK
    ADD CONSTRAINT FK_Negocios_FacturacionProveedores_ProveedorElectronicoDefaultId
    FOREIGN KEY(ProveedorElectronicoDefaultId)
    REFERENCES dbo.FacturacionProveedores(Id);
END
GO

