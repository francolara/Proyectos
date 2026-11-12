-- Firma: FRANCO LARA | Create date: 11/11/2026 | Resuelve el mayor porcentaje de detracción de los materiales de la factura.
USE [SistemaWebHYT];
GO
CREATE OR ALTER FUNCTION dbo.fn_ComprobanteHYT_DetraccionMaterial(@Detalle NVARCHAR(MAX))
RETURNS TABLE
AS
RETURN (
    SELECT TOP (1) catalogo.Codigo,catalogo.Porcentaje
    FROM OPENJSON(@Detalle) WITH (Codigo NVARCHAR(100)) AS linea
    CROSS APPLY (VALUES (NULLIF(TRIM(linea.Codigo),N''))) AS partida(Codigo)
    LEFT JOIN dbo.ServicioMaterial AS interno ON interno.CodigoInterno=partida.Codigo
    OUTER APPLY (
        SELECT MAX(material.MaterialId) AS MaterialId FROM dbo.ServicioMaterial AS material
        WHERE interno.MaterialId IS NULL AND material.CodigoMaterialSAP=partida.Codigo
        HAVING COUNT_BIG(1)=1
    ) AS sap
    INNER JOIN dbo.ServicioMaterial AS material ON material.MaterialId=COALESCE(interno.MaterialId,sap.MaterialId)
    INNER JOIN dbo.DetraccionCatalogoHYT AS catalogo ON catalogo.Codigo=material.CodigoDetraccion
    WHERE catalogo.Activo=1 AND catalogo.Porcentaje>0
    ORDER BY catalogo.Porcentaje DESC,catalogo.Codigo
);
GO
