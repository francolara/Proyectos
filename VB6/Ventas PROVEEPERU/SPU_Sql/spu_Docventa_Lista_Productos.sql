
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE OR ALTER  PROCEDURE [dbo].[spu_Docventa_Lista_Productos]
@idEmpresa		  CHAR(2),
@idLista		  CHAR(8),
@IdAlmacen		  CHAR(8),
@idTipoProducto	  CHAR(6),
@IdNivel		  CHAR(8),
@Busqueda		  VARCHAR(300)
AS
BEGIN

IF @idTipoProducto <> '06002' BEGIN
	SELECT p.CodigoRapido,p.idProducto,p.GlsProducto,m.GlsMarca,p.idUMCompra AS idUMVenta,u.GlsUM,o.idMoneda as GlsMoneda,
	CASE WHEN p.afectoIGV = 1 THEN 'S' ELSE 'N' END Afecto, CAST(ISNULL(a.CantidadStock,0) AS NUMERIC(12,2)) as Stock, t.GlsTallaPeso, p.idfabricante 
	FROM productos p 
	INNER JOIN marcas m ON p.idEmpresa = m.idEmpresa AND p.idMarca = m.idMarca 
	INNER JOIN unidadMedida u ON p.idUMCompra = u.idUM 
	INNER JOIN monedas o ON p.idMoneda = o.idMoneda 
	LEFT JOIN tallapeso t ON p.idEmpresa = t.idEmpresa AND p.idTallaPeso = t.idTallaPeso 
	Left Join ( 
	Select P.idEmpresa,IsNull(vd.idSucursal,'') Idsucursal,P.idProducto, sum(CASE WHEN vd.tipovale = 'I' THEN Cantidad ELSE Cantidad * -1 END) as CantidadStock 
	From Valescab vc Inner Join PeriodosINV pi on vc.idEmpresa = pi.idEmpresa And vc.idPeriodoINV = pi.idPeriodoINV And vc.idSucursal = pi.idSucursal 
	Left Join ValesDet vd On vc.idEmpresa = vd.idEmpresa And vc.idSucursal = vd.idSucursal And vc.tipoVale = vd.tipoVale And vc.idValesCab = vd.idValesCab 
	left Join Productos P On P.IdEmpresa = vd.IdEmpresa And P.IdProducto = vd.IdProducto 
	Where vc.idEmpresa = @idEmpresa And vc.EstValeCab <> 'ANU' And (vc.IdAlmacen = @IdAlmacen Or '' = @IdAlmacen) 
	AND p.idTipoProducto = @idTipoProducto  AND estProducto = 'A' AND (p.IdProducto like @Busqueda OR p.GlsProducto like @Busqueda OR CodigoRapido like @Busqueda OR IdFabricante like @Busqueda) 
	And pi.estPeriodoInv = 'ACT' Group bY P.idEmpresa,vd.idSucursal,P.idProducto
	) A 
	On P.idEmpresa = A.idEmpresa And P.idProducto = A.idProducto 
	Where p.idEmpresa = @idEmpresa
	AND p.idProducto IN (
	SELECT preciosventa.idProducto 
	FROM preciosventa 
	WHERE idEmpresa = @idEmpresa AND idLista = @idLista)  
	AND (p.IdProducto like @Busqueda OR p.GlsProducto like @Busqueda OR CodigoRapido like @Busqueda OR IdFabricante like @Busqueda) 
	AND estProducto = 'A' AND (p.GlsProducto  like @Busqueda OR P.IdProducto like @Busqueda OR CodigoRapido like @Busqueda OR IdFabricante like @Busqueda)  
	AND idTipoProducto = @idTipoProducto AND estProducto = 'A'  
	AND (idNivel = @idNivel OR @idNivel = '')
	order by 2

END ELSE IF @idTipoProducto = '06002' BEGIN

	SELECT p.idProducto,p.GlsProducto,'' as GlsMarca, p.idUMVenta, u.GlsUM, o.idMoneda as GlsMoneda,
	CASE WHEN p.afectoIGV = 1 THEN 'S' ELSE 'N' END Afecto, CAST(0 AS NUMERIC(12,2)) as Stock, '' AS GlsTallaPeso ,'' AS CodigoRapido,P.IdFabricante 
	FROM productos p,monedas o, unidadMedida u WHERE p.idMoneda  = o.idMoneda AND p.idEmpresa = @idEmpresa
	AND p.idTipoProducto = @idTipoProducto AND p.idUMVenta = u.idUM 
	AND (p.GlsProducto  like @Busqueda OR P.IdProducto like @Busqueda OR CodigoRapido like @Busqueda OR IdFabricante like @Busqueda)  
	AND estProducto = 'A' 
	AND (idNivel = @idNivel OR @idNivel = '')
END

;

END;
