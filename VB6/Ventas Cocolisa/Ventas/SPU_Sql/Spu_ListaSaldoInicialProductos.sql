GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
CREATE OR ALTER PROCEDURE dbo.Spu_ListaSaldoInicialProductos
(
    @VarIdEmpresa               CHAR(2),
    @VarIdAlmacen               CHAR(8),
    @VarIdProducto              CHAR(8),
    @VarIdMoneda                CHAR(3),
    @VarFecDesde                VARCHAR(10),
    @VarFecHasta                VARCHAR(10),
    @VarNiveles                 VARCHAR(250),
    @VarTipoNivel               VARCHAR(100),
    @VarIndAgrupoCR             INT,
    @VarCodigoRapido            VARCHAR(50)
)
AS
BEGIN

    DECLARE @VarParam	VARCHAR(10);
    DECLARE @StrSql		VARCHAR(MAX);

    SET @VarParam = (SELECT ValParametro FROM Parametros WHERE IdEmpresa = @VarIdEmpresa AND GlsParametro = 'VIZUALIZA_CODIGO_RAPIDO');

    IF @VarIndAgrupoCR = 1
    BEGIN
        SET @VarParam = 'S';
    END;

    IF @VarIdProducto = ''
    BEGIN
        SET @VarIdProducto = '%%';
    END;

    IF @VarIdAlmacen = ''
    BEGIN
        SET @VarIdAlmacen = '%%';
    END;

    IF OBJECT_ID('TempDB..#TempSaldoIniPro') IS NOT NULL
        DROP TABLE #TempSaldoIniPro;

    CREATE TABLE #TempSaldoIniPro
    (
        IdEmpresa           CHAR(2) NOT NULL DEFAULT '',
        IdSucursal          CHAR(8) NOT NULL DEFAULT '',
        IdValesCab          CHAR(8) NOT NULL DEFAULT '',
        IdConcepto          CHAR(8) NOT NULL DEFAULT '',
        TipoVale            CHAR(1) NOT NULL DEFAULT '',
        FechaEmision        DATE DEFAULT NULL,
        XIdProducto         CHAR(50) NOT NULL DEFAULT '',
        IdProducto          CHAR(8) NOT NULL DEFAULT '',
        GlsProducto         VARCHAR(300) NOT NULL DEFAULT '',
        IdAlmacen           CHAR(8) NOT NULL DEFAULT '',
        IdProvCliente       CHAR(8) NOT NULL DEFAULT '',
        Ruc                 CHAR(11) NULL DEFAULT '',
        GlsPersona          VARCHAR(300) NULL DEFAULT '',
        GlsDocReferencia    VARCHAR(300) NULL DEFAULT '',
        AbreUM              CHAR(8) NOT NULL DEFAULT '',
        Almacen             CHAR(150) NULL DEFAULT '',
        GlsConcepto         VARCHAR(50) NULL DEFAULT '',
        Cantidad            DECIMAL(14,5) DEFAULT NULL,
        VVUnit              DECIMAL(24,10) DEFAULT NULL,
        SaldoInicial        DECIMAL(24,10) DEFAULT NULL,
        Stock               DECIMAL(14,5) DEFAULT NULL,
        CodigoRapido        VARCHAR(50) NOT NULL DEFAULT '',
        PRIMARY KEY(IdEmpresa, IdSucursal, TipoVale, IdValesCab, IdProducto)
    );

    INSERT INTO #TempSaldoIniPro
    (
        IdEmpresa, IdSucursal, IdValesCab, IdConcepto, TipoVale, FechaEmision, XIdProducto, IdProducto, GlsProducto, IdAlmacen,
        IdProvCliente, Ruc, GlsPersona, GlsDocReferencia, AbreUM, Almacen, GlsConcepto, Cantidad, VVUnit, SaldoInicial, Stock, CodigoRapido
    )
    SELECT 
        VC.IdEmpresa, VC.IdSucursal, '' IdValesCab, '' IdConcepto, '' TipoVale, CAST(GETDATE() AS DATE) FechaEmision,
        CASE WHEN @VarParam = 'S' THEN Pr.Codigorapido ELSE VD.IdProducto END AS XIdProducto, VD.IdProducto, Pr.GlsProducto, VC.IdAlmacen, 
        '' IdProvCliente, '' Ruc, '' GlsPersona, '' GlsDocReferencia, um.AbreUM, 
        CONCAT(VC.IdAlmacen, ' ', A.GlsAlmacen) AS Almacen, '' GlsConcepto, 0 Cantidad,
        CAST(SUM(CASE WHEN CAST(VC.FechaEmision AS DATE) < CAST(@VarFecDesde AS DATE) THEN CASE WHEN VC.TipoVale = 'I' THEN VD.Cantidad ELSE VD.Cantidad * -1 END ELSE 0 END *
            CASE @VarIdMoneda
                WHEN 'PEN' THEN CASE WHEN CAST(VC.FechaEmision AS DATE) < CAST(@VarFecDesde AS DATE) THEN CASE WHEN VC.idMoneda = 'PEN' THEN VD.VVUnit ELSE VD.VVUnit * vc.TipoCambio END ELSE 0 END
                WHEN 'USD' THEN CASE WHEN CAST(VC.FechaEmision AS DATE) < CAST(@VarFecDesde AS DATE) THEN CASE WHEN VC.idMoneda = 'USD' THEN VD.VVUnit ELSE VD.VVUnit / vc.TipoCambio END ELSE 0 END
            END) / 
            SUM(CASE WHEN VC.FechaEmision < @VarFecDesde THEN CASE WHEN VC.TipoVale = 'I' THEN VD.Cantidad ELSE VD.Cantidad * -1 END ELSE 0 END) AS DECIMAL(24,10)) AS VVUnit,
        SUM(CASE WHEN VC.FechaEmision < @VarFecDesde THEN CASE WHEN VC.TipoVale = 'I' THEN VD.Cantidad ELSE VD.Cantidad * -1 END ELSE 0 END *
            CASE @VarIdMoneda
                WHEN 'PEN' THEN CASE WHEN CAST(VC.FechaEmision AS DATE) < CAST(@VarFecDesde AS DATE) THEN CASE WHEN VC.idMoneda = 'PEN' THEN VD.VVUnit ELSE VD.VVUnit * vc.TipoCambio END ELSE 0 END
                WHEN 'USD' THEN CASE WHEN CAST(VC.FechaEmision AS DATE) < CAST(@VarFecDesde AS DATE) THEN CASE WHEN VC.idMoneda = 'USD' THEN VD.VVUnit ELSE VD.VVUnit / vc.TipoCambio END ELSE 0 END
            END) /
            SUM(CASE WHEN CAST(VC.FechaEmision AS DATE) < CAST(@VarFecDesde AS DATE) THEN CASE WHEN VC.TipoVale = 'I' THEN VD.Cantidad ELSE VD.Cantidad * -1 END ELSE 0 END) AS SaldoInicial,

        CAST(SUM(CASE WHEN CAST(VC.FechaEmision AS DATE) < CAST(@VarFecDesde AS DATE) THEN CASE WHEN VC.TipoVale = 'I' THEN VD.Cantidad ELSE VD.Cantidad * -1 END ELSE 0 END) AS DECIMAL(14,5)) AS Stock,
        Pr.CodigoRapido
    FROM 
        valescab vc
    INNER JOIN 
        valesdet vd ON VC.IdValesCab = VD.IdValesCab AND VC.IdEmpresa = VD.IdEmpresa AND VC.IdSucursal = VD.IdSucursal AND VC.TipoVale = VD.TipoVale
    LEFT JOIN 
        Conceptos Z ON VC.IdConcepto = Z.IdConcepto
    LEFT JOIN 
        personas pe ON VC.IdProvCliente = pe.IdPersona
    INNER JOIN 
        Almacenes A ON VC.IdEmpresa = A.IdEmpresa AND VC.IdAlmacen = A.IdAlmacen
    INNER JOIN 
        productos pr ON VD.idProducto = Pr.idProducto AND VD.IdEmpresa = Pr.IdEmpresa
    INNER JOIN 
        unidadmedida um ON Pr.idUMCompra = um.idUM
    INNER JOIN 
        niveles n ON Pr.idnivel = n.idnivel AND Pr.IdEmpresa = n.IdEmpresa
    WHERE 
        VC.IdEmpresa = @VarIdEmpresa AND Pr.IdEmpresa = @VarIdEmpresa AND VC.IdAlmacen LIKE @VarIdAlmacen
        AND (VC.idPeriodoInv) IN (
            SELECT pi.idPeriodoInv FROM periodosinv pi
            WHERE pi.IdEmpresa = VC.IdEmpresa AND pi.IdSucursal = VC.IdSucursal
            AND CAST(pi.FecInicio AS DATE) <= CAST(@VarFecHasta AS DATE) AND (CAST(pi.FecFin AS DATE) >= CAST(@VarFecHasta AS DATE) OR pi.FecFin IS NULL)
        )
        AND VC.estValeCab <> 'ANU'
        AND CAST(VC.FechaEmision AS DATE) < CAST(@VarFecDesde AS DATE)
        AND VD.idProducto LIKE @VarIdProducto
        AND Pr.estProducto = 'A' AND (N.Tipo = @VarTipoNivel OR @VarTipoNivel = '')
        AND (Pr.CodigoRapido = @VarCodigoRapido OR @VarCodigoRapido = '')
    GROUP BY VC.IdEmpresa,VC.IdSucursal,VC.IdAlmacen,A.GlsAlmacen, Pr.Codigorapido,Pr.IdProducto , Pr.GlsProducto, um.AbreUM,VD.IdProducto, Pr.GlsProducto
    HAVING  CAST(SUM(CASE WHEN CAST(VC.FechaEmision AS DATE) < CAST(@VarFecDesde AS DATE) THEN CASE WHEN VC.TipoVale = 'I' THEN VD.Cantidad ELSE VD.Cantidad * -1 END ELSE 0 END) AS DECIMAL(14,5)) <> 0.00
    ORDER BY VC.IdAlmacen,CASE WHEN @VarParam = 'S' THEN Pr.Codigorapido ELSE Pr.IdProducto END;



    SET @StrSql = "SELECT A.IdEmpresa, IdSucursal, IdValesCab, IdConcepto, TipoVale, FechaEmision, XIdProducto, A.IdProducto, A.GlsProducto, IdAlmacen, IdProvCliente, Ruc,
					GlsPersona, GlsDocReferencia, AbreUM, Almacen, GlsConcepto, Cantidad, VVUnit, SaldoInicial, Stock, A.CodigoRapido
					FROM #TempSaldoIniPro A
					INNER JOIN Productos P ON A.IdEmpresa = P.IdEmpresa AND A.IdProducto = P.IdProducto
					INNER JOIN vw_niveles vn ON P.idNivel = vn.idNivel01 AND P.IdEmpresa = vn.IdEmpresa
					WHERE A.IdEmpresa = '" + @VarIdEmpresa + "' " +  @VarNiveles ;

    EXEC  (@StrSql)

END