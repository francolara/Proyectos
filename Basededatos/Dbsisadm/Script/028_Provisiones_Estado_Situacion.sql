-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Unifica el estado de compras y ventas provisionadas a PROVISIONADO y normaliza los registros existentes.
-- =============================================

IF EXISTS
(
    SELECT 1
    FROM sys.default_constraints AS dc
    INNER JOIN sys.columns AS c
        ON c.default_object_id = dc.object_id
    INNER JOIN sys.tables AS t
        ON t.object_id = c.object_id
    WHERE t.name = N'VEN_Venta'
      AND c.name = N'Estado'
      AND dc.name = N'DF_VEN_Venta_Estado'
)
BEGIN
    ALTER TABLE dbo.VEN_Venta DROP CONSTRAINT DF_VEN_Venta_Estado;
END;

ALTER TABLE dbo.VEN_Venta
    ADD CONSTRAINT DF_VEN_Venta_Estado DEFAULT (N'PROVISIONADO') FOR Estado;

UPDATE dbo.VEN_Venta
SET Estado = N'PROVISIONADO'
WHERE Estado <> N'PROVISIONADO';

UPDATE dbo.COM_Compra
SET Estado = N'PROVISIONADO'
WHERE Estado <> N'PROVISIONADO';
