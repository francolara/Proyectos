
CREATE OR ALTER PROCEDURE spu_ImpVales
    @varEmpresa CHAR(2),
    @varSucursal CHAR(8),
    @varNumVale CHAR(8),
    @varTipoVale CHAR(1)
AS
BEGIN
    DECLARE @varGlsRuc VARCHAR(180);
    DECLARE @varGlsEmpresa VARCHAR(180);

    SET @varGlsEmpresa = (SELECT glsempresa FROM empresas WHERE idempresa = @varEmpresa);
    SET @varGlsRuc = (SELECT Ruc FROM empresas WHERE idempresa = @varEmpresa);

    SELECT @varGlsRuc AS RucEmpresa,
           @varGlsEmpresa AS GlsEmpresa,
           CASE WHEN c.estValeCab = 'ANU' THEN 'ANULADO' ELSE '' END AS estado,
           c.idValesCab,
           c.tipoVale,
           c.fechaEmision,
           c.idProvCliente,
           p.glsPersona,
           c.idConcepto,
           o.GlsConcepto,
           c.idAlmacen,
           a.GlsAlmacen,
           c.TipoCambio AS TipoCambio,
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
           Z.CodigoRapido,
           d.VVUnit,
           CAST(d.IdTallaPeso AS DECIMAL(14,2)) AS IdTallaPeso
    FROM valescab c
    LEFT JOIN personas p ON c.idProvCliente = p.idPersona
    INNER JOIN valesdet d ON c.idEmpresa = d.idEmpresa 
                           AND c.idSucursal = d.idSucursal 
                           AND c.idValesCab = d.idValesCab 
                           AND c.tipoVale = d.tipoVale
    INNER JOIN Productos Z ON d.IdEmpresa = Z.IdEmpresa 
                            AND d.IdProducto = Z.IdProducto
    INNER JOIN conceptos o ON c.idConcepto = o.idConcepto
    INNER JOIN almacenes a ON c.idEmpresa = a.idEmpresa 
                            AND c.idAlmacen = a.idAlmacen
    INNER JOIN unidadmedida u ON d.idUM = u.idUM
    INNER JOIN monedas m ON c.idMoneda = m.idMoneda
    LEFT JOIN centroscosto cc ON c.idCentroCosto = cc.idCentroCosto 
                               AND c.idempresa = cc.idempresa
    LEFT JOIN TiposdeCambio tc ON c.fechaEmision = tc.fecha
    WHERE c.idEmpresa = @varEmpresa
      AND c.idSucursal = @varSucursal
      AND c.idValesCab = @varNumVale
      AND c.TipoVale = @varTipoVale;
END;
GO
