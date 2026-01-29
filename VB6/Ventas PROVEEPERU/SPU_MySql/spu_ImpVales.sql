DELIMITER $$

DROP PROCEDURE IF EXISTS  `spu_ImpVales` $$
CREATE   PROCEDURE `spu_ImpVales`(
varEmpresa     CHAR(2),
varSucursal    CHAR(8),
varNumVale     CHAR(8),
varTipoVale    CHAR(1)
)
BEGIN

DECLARE varGlsRuc     VARCHAR(180);
DECLARE varGlsEmpresa VARCHAR(180);

SET varGlsEmpresa = (SELECT glsempresa FROM empresas WHERE idempresa = varEmpresa);
SET varGlsRuc = (SELECT Ruc FROM empresas WHERE idempresa = varEmpresa);

SELECT varGlsRuc as RucEmpresa, varGlsEmpresa as GlsEmpresa,
if(c.estValeCab = 'ANU','ANULADO','') as estado,c.idValesCab,c.tipoVale,c.fechaEmision,c.idProvCliente,p.glsPersona,
c.idConcepto,o.GlsConcepto,c.idAlmacen,a.GlsAlmacen, /*tcventa*/ C.TipoCambio TipoCambio,
c.GlsDocReferencia, c.obsValesCab, c.idMoneda, m.glsMoneda, c.idCentroCosto, cc.glsCentroCosto,
d.idProducto,d.GlsProducto,d.idUM,u.GlsUM,u.abreUM, d.Factor,d.Cantidad,d.Cantidad2,
d.NumLote,d.FecVencProd, d.TotalVVNeto, d.TotalIGVNeto, d.TotalPVNeto,Z.CodigoRapido, d.VVUnit,Cast(d.IdTallaPeso As Decimal(14,2)) IdTallaPeso
FROM valescab c
left join personas p
  on c.idProvCliente = p.idPersona
inner join valesdet d
  on c.idEmpresa = d.idEmpresa and c.idSucursal = d.idSucursal and c.idValesCab = d.idValesCab and c.tipoVale = d.tipoVale
Inner Join Productos Z
  On d.IdEmpresa = Z.IdEmpresa And d.IdProducto = Z.IdProducto
inner join conceptos o
  on c.idConcepto = o.idConcepto
inner join almacenes a
  on c.idEmpresa = a.idEmpresa and c.idAlmacen = a.idAlmacen
INNER join unidadmedida u
  on d.idUM = u.idUM
inner join monedas m
  on c.idMoneda = m.idMoneda
left join centroscosto cc
  on c.idCentroCosto = cc.idCentroCosto and c.idempresa = cc.idempresa
Left Join TiposdeCambio tc
  On c.fechaEmision =  tc.fecha
WHERE c.idEmpresa = varEmpresa
AND c.idSucursal = varSucursal
AND c.idValesCab = varNumVale
AND c.TipoVale = varTipoVale;

END $$

DELIMITER ;