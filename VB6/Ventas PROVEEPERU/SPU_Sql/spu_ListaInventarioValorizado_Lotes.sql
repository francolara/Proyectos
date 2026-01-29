CREATE OR ALTER PROCEDURE dbo.spu_ListaInventarioValorizado_Lotes
    @varEmpresa CHAR(2),
    @varSucursal CHAR(8),
    @varAlmacen VARCHAR(8),
    @varMoneda CHAR(3),
    @varFecha VARCHAR(20),
    @varNiveles VARCHAR(250),
    @varGlsNiveles VARCHAR(250),
    @VarOrdena			VARCHAR(30),
    @VarCodigoRapido	VARCHAR(50),
	@varNivel01		VARCHAR(8),
	@varNivel02		VARCHAR(8)
AS
BEGIN
    DECLARE @varGlsSuc VARCHAR(180);
    DECLARE @varGlsAlm VARCHAR(180);
    DECLARE @varGlsMon VARCHAR(80);
    DECLARE @varGlsEmpresa VARCHAR(180);
    DECLARE @varGlsRuc VARCHAR(180);
    DECLARE @varGlsSistema VARCHAR(180);
    DECLARE @strSQL NVARCHAR(MAX);
    DECLARE @VarStock NVARCHAR(MAX);
    DECLARE @varCodRapido VARCHAR(2);
	DECLARE @SQLQuery NVARCHAR(MAX);

    SET @varCodRapido = ISNULL((SELECT valparametro FROM parametros WHERE idempresa = @varEmpresa AND glsparametro = 'VIZUALIZA_CODIGO_RAPIDO'), '');

    IF @VarOrdena = 'C'
    BEGIN
        SET @VarOrdena = 'vd.IdProducto';
    END
    ELSE
    BEGIN
        SET @VarOrdena = 'p.GlsProducto';
    END;

    IF @varEmpresa <> ''
    BEGIN
        SET @varGlsEmpresa = (SELECT GlsEmpresa FROM empresas WHERE idempresa = @varEmpresa);
    END;

    IF @varEmpresa <> ''
    BEGIN
        SET @varGlsRuc = (SELECT ruc FROM empresas WHERE idempresa = @varEmpresa);
    END;

    SET @varGlsSistema = 'Sistema de Inventarios';

    IF @varSucursal <> ''
    BEGIN
        SET @varGlsSuc = (SELECT GlsPersona FROM Personas WHERE idPersona = @varSucursal);
    END
    ELSE
    BEGIN
        SET @varGlsSuc = 'TODAS LAS SUCURSALES';
    END;

    IF @varAlmacen <> ''
    BEGIN
        SET @varGlsAlm = (SELECT GlsAlmacen FROM Almacenes WHERE idEmpresa = @varEmpresa AND idAlmacen = @varAlmacen);
    END
    ELSE
    BEGIN
        SET @varGlsAlm = 'TODOS LOS ALMACENES';
    END;

    SET @varFecha = CONVERT(DATE, @varFecha);

    SET @varGlsMon = (SELECT GlsMoneda FROM Monedas WHERE idMoneda = @varMoneda);

	SELECT x.idEmpresa,x.idNivel01, x.GlsNivel01,x.idNivel02, x.GlsNivel02,
		   CONVERT(VARCHAR,CAST(@varFecha AS DATE), 103) AS FechaCorte, 
		   x.idProducto, x.GlsProducto, x.abreUM,
		   SUM(STOCK) AS STOCK,
		   IIF(Round(SUM(STOCK),2) > 0,((SUM(STOCK) * SUM(VVUNITconvert)) / SUM(STOCK))
		   , 0) AS Costo,
		   @varGlsAlm AS GlsAlmacen,
		   @varGlsMon AS GlsMoneda, 
		   @varGlsSuc AS GlsSucursal, 
		   @varGlsEmpresa AS varGlsEmpresa, 
		   @varGlsRuc AS varGlsRuc, 
		   @varGlsSistema AS varGlsSistema,
		   x.IdFabricante, x.GlsMarca,
		   x.IdLote, x.NumLote
		   FROM	(
	
		SELECT vn.idEmpresa,vn.idNivel01, vn.GlsNivel01,vn.idNivel02, vn.GlsNivel02, 		   
		   vd.idProducto AS idProducto, 
		   p.GlsProducto, 
		   u.abreUM, 
		   
		   IIF((vc.tipoVale = 'I'), vd.Cantidad, (vd.Cantidad * -1)) AS STOCK, 

					P.IdFabricante, M.GlsMarca,

					((IIF(vc.tipoVale = 'I', vd.Cantidad, vd.Cantidad * -1)) * 
					CASE @varMoneda
						WHEN 'PEN' THEN IIF(vc.idMoneda = 'PEN', vd.VVUnit, vd.VVUnit * vc.TipoCambio) 
						WHEN 'USD' THEN IIF(vc.idMoneda = 'USD', vd.VVUnit, vd.VVUnit / vc.TipoCambio) 
					END) as VVUNITconvert,vd.IdLote,vd.NumLote

	FROM valescab vc 
	INNER JOIN valesdet vd ON vc.idValesCab = vd.idValesCab 
		 AND vc.idEmpresa = vd.idEmpresa 
		 AND vc.idSucursal = vd.idSucursal 
		 AND vc.tipoVale = vd.tipoVale 
	INNER JOIN productos p ON vd.idProducto = p.idProducto 
		 AND vd.idEmpresa = p.idEmpresa 
	INNER JOIN unidadmedida u ON p.idUMCompra = u.idUM 
	INNER JOIN conceptos c ON vc.idConcepto = c.idConcepto 
	INNER JOIN vw_niveles vn ON p.idNivel = vn.idNivel01 
		 AND p.idEmpresa = vn.idEmpresa 
	LEFT JOIN Marcas M ON P.IdMarca = M.IdMarca 
		 AND P.IdEmpresa = M.IdEmpresa 
	INNER JOIN Almacenes A ON vc.IdEmpresa = A.IdEmpresa 
		 AND vc.IdAlmacen = A.IdAlmacen 
	LEFT JOIN personas pe ON vc.idProvCliente = pe.IdPersona 
	INNER JOIN niveles n ON p.idnivel = n.idnivel 
		 AND p.idempresa = n.idempresa 
	WHERE vc.idEmpresa = @varEmpresa 
	  AND p.estProducto = 'A' 
	  AND vc.estValeCab <> 'ANU' 
	  AND (vc.idSucursal = @varSucursal OR @varSucursal = '') 
	  AND (vc.idPeriodoInv IN 
		   (SELECT pi.idPeriodoInv 
			FROM periodosinv pi 
			WHERE pi.idEmpresa = vc.idEmpresa 
			  AND pi.idSucursal = vc.idSucursal 
			  AND CAST(pi.FecInicio AS DATE) <= CAST(@varFecha AS DATE)
			  AND (CAST(pi.FecFin AS DATE) >= CAST(@varFecha AS DATE) OR pi.FecFin IS NULL)
		   )
		 ) 
	  AND (vc.idAlmacen = @varAlmacen OR @varAlmacen = '') 
	  AND vc.fechaEmision <= CAST(@varFecha AS DATE) 
	  AND (p.CodigoRapido = @VarCodigoRapido OR @VarCodigoRapido = '') 
	 AND (vn.idNivel02 = @varNivel01 OR @varNivel01 = '')
	 AND (vn.idNivel01 = @varNivel02 OR @varNivel02 = '')
	) X
	GROUP BY x.idEmpresa,x.idNivel01, x.GlsNivel01,x.idNivel02, x.GlsNivel02,x.idProducto, x.GlsProducto, x.abreUM,x.IdFabricante, x.GlsMarca, x.IdLote, x.NumLote
	HAVING ROUND(SUM(STOCK),2) <> 0.00
	ORDER BY x.idProducto ;

END