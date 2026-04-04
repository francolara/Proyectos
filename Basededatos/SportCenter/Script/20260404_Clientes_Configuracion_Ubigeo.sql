-- =============================================
-- Author:        FRANCO LARA
-- Create date:   04/04/2026
-- Description:   Integra ubigeo fiscal en Clientes y Configuracion de club, con combos encadenados por departamento/provincia/distrito.
-- Firma:         Codex - 04/04/2026 | Agrega CodigoUbigeo en Clientes/Negocios, FKs y SP de ubigeo + ajustes CRUD.
-- =============================================

IF COL_LENGTH('dbo.Clientes', 'CodigoUbigeo') IS NULL
BEGIN
    ALTER TABLE dbo.Clientes ADD CodigoUbigeo CHAR(6) NULL;
END;
GO

IF COL_LENGTH('dbo.Negocios', 'CodigoUbigeo') IS NULL
BEGIN
    ALTER TABLE dbo.Negocios ADD CodigoUbigeo CHAR(6) NULL;
END;
GO

IF NOT EXISTS
(
    SELECT 1
    FROM sys.foreign_keys
    WHERE name = N'FK_Clientes_UbigeoDistritos_CodigoUbigeo'
      AND parent_object_id = OBJECT_ID(N'dbo.Clientes')
)
BEGIN
    ALTER TABLE dbo.Clientes WITH CHECK
    ADD CONSTRAINT FK_Clientes_UbigeoDistritos_CodigoUbigeo
    FOREIGN KEY (CodigoUbigeo) REFERENCES dbo.UbigeoDistritos (CodigoUbigeo);
END;
GO

IF NOT EXISTS
(
    SELECT 1
    FROM sys.foreign_keys
    WHERE name = N'FK_Negocios_UbigeoDistritos_CodigoUbigeo'
      AND parent_object_id = OBJECT_ID(N'dbo.Negocios')
)
BEGIN
    ALTER TABLE dbo.Negocios WITH CHECK
    ADD CONSTRAINT FK_Negocios_UbigeoDistritos_CodigoUbigeo
    FOREIGN KEY (CodigoUbigeo) REFERENCES dbo.UbigeoDistritos (CodigoUbigeo);
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.Clientes') AND name = N'IX_Clientes_CodigoUbigeo')
BEGIN
    CREATE NONCLUSTERED INDEX IX_Clientes_CodigoUbigeo ON dbo.Clientes (CodigoUbigeo);
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.Negocios') AND name = N'IX_Negocios_CodigoUbigeo')
BEGIN
    CREATE NONCLUSTERED INDEX IX_Negocios_CodigoUbigeo ON dbo.Negocios (CodigoUbigeo);
END;
GO

PRINT 'Estructura ubigeo fiscal aplicada (Clientes/Negocios).';
GO
