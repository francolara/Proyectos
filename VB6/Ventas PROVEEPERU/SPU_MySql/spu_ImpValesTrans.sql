DELIMITER $$

DROP PROCEDURE IF EXISTS `dbsistema`.`spu_ImpValesTrans` $$
CREATE DEFINER=`root`@`localhost` PROCEDURE `spu_ImpValesTrans`(
varEmpresa       CHAR(2),
varSucursal      CHAR(8),
varNumVale       CHAR(8),
varTipoVale      CHAR(1),
varNumValeTrans  CHAR(8)
)
BEGIN

SELECT if(c.estValeCab = 'ANU','ANULADO','') as estado,c.idValesCab,c.tipoVale,c.fechaEmision,c.idProvCliente,p.glsPersona,
c.idConcepto,o.GlsConcepto,c.idAlmacen,a.GlsAlmacen, c.TipoCambio,
c.GlsDocReferencia, c.obsValesCab, c.idMoneda, m.glsMoneda, c.idCentroCosto, cc.glsCentroCosto,
d.idProducto,d.GlsProducto,d.idUM,u.GlsUM,u.abreUM, d.Factor,d.Cantidad,d.Cantidad2,
d.NumLote,d.FecVencProd, d.TotalVVNeto, d.TotalIGVNeto, d.TotalPVNeto,
(
select idAlmacenDestino from valestrans where idValesTrans = varNumValeTrans and idEmpresa = varEmpresa and idSucursal = varSucursal
)   AlmacenDestino,
(
select glsalmacen from valestrans x inner join almacenes z on x.idAlmacenDestino = z.idalmacen and x.idempresa = z.idempresa
where x.idValesTrans = varNumValeTrans and x.idEmpresa = varEmpresa and x.idSucursal = varSucursal
)   GlsAlmacenDestino

FROM valescab c
inner join valesdet d
on c.idEmpresa = d.idEmpresa
and c.idSucursal = d.idSucursal
and c.idValesCab = d.idValesCab
and c.tipoVale = d.tipoVale
left join personas p on c.idProvCliente = p.idPersona
inner join conceptos o on c.idConcepto = o.idConcepto
inner join almacenes a on c.idEmpresa = a.idEmpresa and c.idAlmacen = a.idAlmacen
inner join unidadmedida u on d.idUM = u.idUM
inner join monedas m on c.idMoneda = m.idMoneda
left join centroscosto cc on c.idCentroCosto = cc.idCentroCosto
WHERE c.idEmpresa = varEmpresa
AND c.idSucursal = varSucursal
AND c.idValesCab = varNumVale
AND c.TipoVale = varTipoVale
;

END $$

DELIMITER ;