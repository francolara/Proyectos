-- SOURCE: 20260507_Monedas_Unicidad_PorNegocio.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   07/05/2026
-- Description:   Ajusta unicidad de Monedas por negocio para permitir mismo codigo en diferentes negocios.
-- =============================================

IF EXISTS (
    SELECT 1
    FROM sys.key_constraints
    WHERE [type] = 'UQ'
      AND [name] = 'UQ_Monedas_Codigo'
      AND [parent_object_id] = OBJECT_ID('dbo.Monedas')
)
BEGIN
    ALTER TABLE dbo.Monedas DROP CONSTRAINT UQ_Monedas_Codigo;
END;

IF NOT EXISTS (
    SELECT 1
    FROM sys.indexes
    WHERE [name] = 'UX_Monedas_Negocio_Codigo'
      AND [object_id] = OBJECT_ID('dbo.Monedas')
)
BEGIN
    CREATE UNIQUE NONCLUSTERED INDEX UX_Monedas_Negocio_Codigo
        ON dbo.Monedas (NegocioId, Codigo)
        WHERE NegocioId IS NOT NULL;
END;