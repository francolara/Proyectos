CREATE OR ALTER PROCEDURE dbo.spu_ImpValesTrans
    @varEmpresa CHAR(2),
    @varSucursal CHAR(8),
    @varNumVale CHAR(8),
    @varTipoVale CHAR(1),
    @varNumValeTrans CHAR(8)
AS
BEGIN
    SELECT CASE WHEN c.estValeCab = 'ANU' THEN 'ANULADO' ELSE '' END AS estado,
           c.idValesCab,
           c.tipoVale,
           c.fechaEmision,
           c.idProvCliente,
           p.glsPersona,
           c.idConcepto,
           o.GlsConcepto,
           c.idAlmacen,
           a.GlsAlmacen,
           c.TipoCambio,
           c.GlsDocReferencia,
           c.obsValesCab,
           c.idMoneda,
           m.glsMoneda,
           c.idCentroCosto,
           cc.glsCentroCosto,
           d.idProducto,
           d.GlsProducto,
           d.idUM,
           u.GlsUM,
           u.abreUM,
           d.Factor,
           d.Cantidad,
           d.Cantidad2,
           d.NumLote,
           d.FecVencProd,
           d.TotalVVNeto,
           d.TotalIGVNeto,
           d.TotalPVNeto,
           (SELECT idAlmacenDestino 
            FROM valestrans 
            WHERE idValesTrans = @varNumValeTrans 
            AND idEmpresa = @varEmpresa 
            AND idSucursal = @varSucursal) AS AlmacenDestino,
           (SELECT glsalmacen 
            FROM valestrans x 
            INNER JOIN almacenes z ON x.idAlmacenDestino = z.idalmacen 
                                   AND x.idempresa = z.idempresa
            WHERE x.idValesTrans = @varNumValeTrans 
            AND x.idEmpresa = @varEmpresa 
            AND x.idSucursal = @varSucursal) AS GlsAlmacenDestino
    FROM valescab c
    INNER JOIN valesdet d ON c.idEmpresa = d.idEmpresa 
                           AND c.idSucursal = d.idSucursal 
                           AND c.idValesCab = d.idValesCab 
                           AND c.tipoVale = d.tipoVale
    LEFT JOIN personas p ON c.idProvCliente = p.idPersona
    INNER JOIN conceptos o ON c.idConcepto = o.idConcepto
    INNER JOIN almacenes a ON c.idEmpresa = a.idEmpresa 
                            AND c.idAlmacen = a.idAlmacen
    INNER JOIN unidadmedida u ON d.idUM = u.idUM
    INNER JOIN monedas m ON c.idMoneda = m.idMoneda
    LEFT JOIN centroscosto cc ON c.idCentroCosto = cc.idCentroCosto
    WHERE c.idEmpresa = @varEmpresa
      AND c.idSucursal = @varSucursal
      AND c.idValesCab = @varNumVale
      AND c.TipoVale = @varTipoVale;
END;
