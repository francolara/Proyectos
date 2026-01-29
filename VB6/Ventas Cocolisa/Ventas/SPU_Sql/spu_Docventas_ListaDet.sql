
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE OR ALTER  PROCEDURE [dbo].[spu_Docventas_ListaDet]
@idEmpresa		CHAR(2),
@idSucursal		VARCHAR(8),
@idDocumento	CHAR(2),
@idDocVentas	CHAR(8),
@idSerie		CHAR(4)
AS
BEGIN

SELECT item, idProducto, CAST((CAST(GlsProducto AS VARCHAR(500)) + ' ( ' + NumLote + ' ) ') AS VARCHAR(8000)) AS GlsProducto, GlsMarca, GlsUM, 
CAST(Cantidad AS NUMERIC(12,2)) AS Cantidad,CAST(Cantidad2 AS NUMERIC(12,2)) AS Cantidad2, CAST(PVUnit AS NUMERIC(12,2)) AS PVUnit, 
PorDcto, CAST(TotalPVNeto AS NUMERIC(12,2)) AS TotalPVNeto 
FROM docventasdet WHERE idEmpresa = @idEmpresa
AND idSucursal = @idSucursal 
AND idDocumento = @idDocumento
AND idDocVentas = @idDocVentas
AND idSerie = @idSerie 
Order By Item


END;