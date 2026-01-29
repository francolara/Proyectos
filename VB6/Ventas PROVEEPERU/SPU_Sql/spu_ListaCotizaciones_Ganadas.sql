
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE OR ALTER  PROCEDURE [dbo].[spu_ListaCotizaciones_Ganadas]
@idEmpresa		CHAR(2),
@idSucursal		VARCHAR(8),
@idProducto		CHAR(8)
AS
BEGIN

SELECT 
(b.idSerie + '-' + b.idDocVentas) As Coti,convert(varchar,b.FecEmision,103) as Fecha,b.GlsMoneda,CAST(a.Cantidad AS DECIMAL(10,2)) AS Cantidad ,CAST(a.VVUnit AS DECIMAL(10,2)) AS PVUnit,CAST(a.TotalVVNeto AS DECIMAL(10,2)) AS TotalPVNeto,b.GlsCliente,
(e.idDocVentas) As Ord,convert(varchar,e.FecEmision,103) as FechaOc ,e.GlsMoneda as GlsMonedaOc,CAST(d.Cantidad AS DECIMAL(10,2)) as CantidadOc,CAST(d.VVUnit AS DECIMAL(10,2)) as PVUnitOc ,CAST(d.TotalVVNeto AS DECIMAL(10,2)) AS TotalPVNetoOC,e.GlsCliente As Proveedor
FROM docventasdet A
INNER JOIN docventas B
ON A.idEmpresa = B.idEmpresa
AND A.idSucursal = B.idSucursal
AND A.idSerie = B.idSerie
and a.idDocVentas = b.idDocVentas
and a.idDocumento = b.idDocumento
Left Join docreferencia c
on c.idEmpresa = b.idEmpresa
and c.idSucursal = b.idSucursal
and c.tipoDocReferencia = b.idDocumento
and c.numDocReferencia = b.idDocVentas
and c.serieDocReferencia = b.idSerie
and c.tipoDocOrigen = '94'
left join docventasdet d
ON d.idEmpresa = c.idEmpresa
AND d.idSucursal = c.idSucursal
AND d.idSerie = c.serieDocOrigen
and d.idDocVentas = c.numDocOrigen
and d.idDocumento = c.tipoDocOrigen
and d.idProducto = a.idProducto
and d.idDocumento = '94'

left JOIN docventas e
ON d.idEmpresa = e.idEmpresa
AND d.idSucursal = e.idSucursal
AND d.idSerie = e.idSerie
and d.idDocVentas = e.idDocVentas
and d.idDocumento = e.idDocumento

WHERE a.idEmpresa = @idEmpresa
AND a.idSucursal = @idSucursal 
AND a.idProducto = @idProducto
and B.estDocVentas <> 'ANU'
AND ISNULL(b.ind_CTGanada,0) = 1 
and a.idDocumento = '92'
Order By b.FecEmision ASC


END;