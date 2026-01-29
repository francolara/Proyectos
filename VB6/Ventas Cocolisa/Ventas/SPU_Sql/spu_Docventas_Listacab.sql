
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE OR ALTER  PROCEDURE [dbo].[spu_Docventas_Listacab]
@idEmpresa		CHAR(2),
@idSucursal		VARCHAR(8),
@idDocumento	CHAR(2),
@Anno			INT,
@MES			INT,
@Busqueda		VARCHAR(200)
AS
BEGIN

IF @Busqueda <> '' BEGIN
	SET @Busqueda = '%' + @Busqueda + '%'
END ELSE BEGIN
	SET @Busqueda = '%%'
END

SELECT (A.idDocumento + A.idDocVentas + A.idSerie) as Item , A.idDocVentas,A.idSerie,A.idPerCliente,A.GlsCliente, A.RUCCliente,
CAST(A.FecEmision AS DATE) as FecEmision,A.estDocVentas,  
CASE WHEN A.idMoneda='USD' THEN 'Dolares' ELSE 'Soles' END as Moneda,
CAST(A.TotalPrecioVenta AS NUMERIC(12,2)) AS TotalPrecioVenta,B.docReferencia,A.NumOrdenCompra 
FROM docventas A  
Left Join( 
Select IdEmpresaReferencia,IdSucursalReferencia,TipoDocReferencia,SerieDocReferencia,  NumDocReferencia,IdEmpresaOrigen,IdSucursalOrigen,TipoDocOrigen,SerieDocOrigen,NumDocOrigen, 
(d.AbreDocumento + '' + serieDocReferencia + '-' + numDocReferencia) As docReferencia  
From docreferenciaempresas dr  Inner Join documentos d  On dr.tipoDocReferencia = d.idDocumento  
Where IdEmpresaReferencia ='' And IdSucursalReferencia = ''  
Group By IdEmpresaReferencia,IdSucursalReferencia,TipoDocReferencia,SerieDocReferencia,NumDocReferencia,IdEmpresaOrigen,IdSucursalOrigen,TipoDocOrigen,
SerieDocOrigen,NumDocOrigen,d.AbreDocumento,serieDocReferencia,numDocReferencia ) B  
On A.IdEmpresa = B.IdEmpresaOrigen  And A.IdSucursal = B.IdSucursalOrigen  And A.IdDocumento = B.TipoDocOrigen  
And A.IdSerie = B.SerieDocOrigen  And A.IdDocVentas = B.NumDocOrigen  
WHERE A.idEmpresa = @idEmpresa
AND A.idSucursal = @idSucursal 
AND A.idDocumento = @idDocumento 
And IndAtribucionNC = 0 
AND year(A.FecEmision) = @Anno 
AND Month(A.FecEmision) = @MES
and (a.idDocVentas like @Busqueda or A.GlsCliente like @Busqueda)
ORDER BY A.idSerie,A.idDocVentas,A.FecEmision

END;