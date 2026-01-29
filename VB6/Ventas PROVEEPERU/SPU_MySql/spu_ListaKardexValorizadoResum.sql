DELIMITER $$

DROP PROCEDURE IF EXISTS `spu_ListaKardexValorizadoResum` $$
CREATE  PROCEDURE `spu_ListaKardexValorizadoResum`(
varEmpresa      CHAR(2),
varAlmacen      VARCHAR(8),
varMoneda       CHAR(3),
varFechaIni     VARCHAR(30),
varFechaFin     VARCHAR(30),
varProducto     VARCHAR(8)
)
BEGIN

DECLARE varGlsSuc          VARCHAR(180);
DECLARE varGlsAlm          VARCHAR(180);
DECLARE varGlsMon          VARCHAR(80);
DECLARE varGlsEmpresa      VARCHAR(180);
DECLARE varGlsRuc          VARCHAR(180);
DECLARE varGlsProducto     VARCHAR(180);
DECLARE varGlsSistema      VARCHAR(180);


IF varEmpresa <> '' THEN
  SET varGlsEmpresa = (SELECT GlsEmpresa FROM empresas where idempresa = varEmpresa);
END IF;

IF varEmpresa <> '' THEN
  SET varGlsRuc = (SELECT ruc FROM empresas where idempresa = varEmpresa);
END IF;

IF varProducto = '' THEN
  Set varGlsProducto = 'TODOS LOS PRODUCTOS';
else
  Set varGlsProducto = (Select GlsProducto from productos where idproducto = varProducto and idempresa = varEmpresa );
end if;

IF varAlmacen = '' THEN
  SET varGlsAlm = 'TODOS LOS ALMACENES';
ELSE
  SET varGlsAlm = (SELECT GlsAlmacen FROM Almacenes WHERE idEmpresa = varEmpresa AND idAlmacen = varAlmacen);
END IF;

IF varAlmacen = '' THEN
  SET varAlmacen = '%%';
END IF;

IF varProducto = '' THEN
  SET varProducto = '%%';
END IF;

  SET varGlsMon = (SELECT GlsMoneda FROM Monedas WHERE idMoneda = varMoneda);



  SELECT varGlsAlm,vd.Cantidad,varGlsEmpresa,varGlsRuc,DATE_FORMAT(varFechaIni,'%d/%m/%y') FecIni,DATE_FORMAT(varFechaFin,'%d/%m/%y') FecFin,varGlsProducto,varGlsMon,
  --  vn.idNivel01, vn.GlsNivel01 ,vn.idNivel02, vn.GlsNivel02,
  pr.CtaContable2,n.idnivel,n.glsNivel,um.abreUM, pe.GlsPersona,

  Sum(IF(z.idconcepto = '05' and vc.tipoVale = 'I',(vd.Cantidad),0) *
  CASE varMoneda WHEN 'PEN' THEN  IF(z.idconcepto = '05' and vc.tipoVale = 'I',IF(vc.idMoneda = 'PEN', vd.VVUnit,vd.VVUnit *  vc.TipoCambio ),0)
                 WHEN 'USD' THEN  IF(z.idconcepto = '05' and vc.tipoVale = 'I',IF(vc.idMoneda = 'USD', vd.VVUnit,vd.VVUnit /  vc.TipoCambio ),0) END ) ValTotalComprasI,

  Sum(IF(z.idconcepto = '26' and vc.tipoVale = 'I',(vd.Cantidad),0) *
  CASE varMoneda WHEN 'PEN' THEN  IF(z.idconcepto = '26' and vc.tipoVale = 'I',IF(vc.idMoneda = 'PEN', vd.VVUnit,vd.VVUnit *  vc.TipoCambio ),0)
                 WHEN 'USD' THEN  IF(z.idconcepto = '26' and vc.tipoVale = 'I',IF(vc.idMoneda = 'USD', vd.VVUnit,vd.VVUnit /  vc.TipoCambio ),0) END ) ValTotalTransferenciaI,

  Sum(IF(z.idconcepto = '13' and vc.tipoVale = 'I',(vd.Cantidad),0) *
  CASE varMoneda WHEN 'PEN' THEN  IF(z.idconcepto = '13' and vc.tipoVale = 'I',IF(vc.idMoneda = 'PEN', vd.VVUnit,vd.VVUnit *  vc.TipoCambio ),0)
                 WHEN 'USD' THEN  IF(z.idconcepto = '13' and vc.tipoVale = 'I',IF(vc.idMoneda = 'USD', vd.VVUnit,vd.VVUnit /  vc.TipoCambio ),0) END ) ValTotalConversionI,

  Sum(IF(z.idconcepto = '29' and vc.tipoVale = 'I',(vd.Cantidad),0) *
  CASE varMoneda WHEN 'PEN' THEN  IF(z.idconcepto = '29' and vc.tipoVale = 'I',IF(vc.idMoneda = 'PEN', vd.VVUnit,vd.VVUnit *  vc.TipoCambio ),0)
                 WHEN 'USD' THEN  IF(z.idconcepto = '29' and vc.tipoVale = 'I',IF(vc.idMoneda = 'USD', vd.VVUnit,vd.VVUnit /  vc.TipoCambio ),0) END ) ValTotalImportacionI,

  Sum(IF(z.idconcepto not in('05','26','13','29') and vc.tipoVale = 'I',(vd.Cantidad),0) *
  CASE varMoneda WHEN 'PEN' THEN  IF(z.idconcepto not in('05','26','13','29') and vc.tipoVale = 'I',IF(vc.idMoneda = 'PEN', vd.VVUnit,vd.VVUnit *  vc.TipoCambio ),0)
                 WHEN 'USD' THEN  IF(z.idconcepto not in('05','26','13','29') and vc.tipoVale = 'I',IF(vc.idMoneda = 'USD', vd.VVUnit,vd.VVUnit /  vc.TipoCambio ),0) END ) ValTotalOtrosI,


  Sum(IF(z.idconcepto = '22' and vc.tipoVale = 'S',(vd.Cantidad),0) *
  CASE varMoneda WHEN 'PEN' THEN  IF(z.idconcepto = '22' and vc.tipoVale = 'S',IF(vc.idMoneda = 'PEN', vd.VVUnit,vd.VVUnit *  vc.TipoCambio ),0)
                     WHEN 'USD' THEN  IF(z.idconcepto = '22' and vc.tipoVale = 'S',IF(vc.idMoneda = 'USD', vd.VVUnit,vd.VVUnit /  vc.TipoCambio ),0) END ) ValTotalVentasS,

  Sum(IF(z.idconcepto = '46' and vc.tipoVale = 'S',(vd.Cantidad),0) *
  CASE varMoneda WHEN 'PEN' THEN  IF(z.idconcepto = '46' and vc.tipoVale = 'S',IF(vc.idMoneda = 'PEN', vd.VVUnit,vd.VVUnit *  vc.TipoCambio ),0)
                     WHEN 'USD' THEN  IF(z.idconcepto = '46' and vc.tipoVale = 'S',IF(vc.idMoneda = 'USD', vd.VVUnit,vd.VVUnit /  vc.TipoCambio ),0) END ) ValTotalVentasSujetaS,

  Sum(IF(z.idconcepto = '20' and vc.tipoVale = 'S',(vd.Cantidad),0) *
  CASE varMoneda WHEN 'PEN' THEN  IF(z.idconcepto = '20' and vc.tipoVale = 'S',IF(vc.idMoneda = 'PEN', vd.VVUnit,vd.VVUnit *  vc.TipoCambio ),0)
                 WHEN 'USD' THEN  IF(z.idconcepto = '20' and vc.tipoVale = 'S',IF(vc.idMoneda = 'USD', vd.VVUnit,vd.VVUnit /  vc.TipoCambio ),0) END ) ValTotalConsumoS,

  Sum(IF(z.idconcepto = '25' and vc.tipoVale = 'S',(vd.Cantidad),0) *
  CASE varMoneda WHEN 'PEN' THEN  IF(z.idconcepto = '25' and vc.tipoVale = 'S',IF(vc.idMoneda = 'PEN', vd.VVUnit,vd.VVUnit *  vc.TipoCambio ),0)
                 WHEN 'USD' THEN  IF(z.idconcepto = '25' and vc.tipoVale = 'S',IF(vc.idMoneda = 'USD', vd.VVUnit,vd.VVUnit /  vc.TipoCambio ),0) END ) ValTotalTransferenciaS,

  Sum(IF(z.idconcepto = '16' and vc.tipoVale = 'S',(vd.Cantidad),0) *
  CASE varMoneda WHEN 'PEN' THEN  IF(z.idconcepto = '16' and vc.tipoVale = 'S',IF(vc.idMoneda = 'PEN', vd.VVUnit,vd.VVUnit *  vc.TipoCambio ),0)
                 WHEN 'USD' THEN  IF(z.idconcepto = '16' and vc.tipoVale = 'S',IF(vc.idMoneda = 'USD', vd.VVUnit,vd.VVUnit /  vc.TipoCambio ),0) END ) ValTotalCoversionS,

  Sum(IF(z.idconcepto not in('22','20','25','16','46') and vc.tipoVale = 'S',(vd.Cantidad),0) *
  CASE varMoneda WHEN 'PEN' THEN  IF(z.idconcepto not in('22','20','25','16','46') and vc.tipoVale = 'S',IF(vc.idMoneda = 'PEN', vd.VVUnit,vd.VVUnit *  vc.TipoCambio ),0)
                 WHEN 'USD' THEN  IF(z.idconcepto not in('22','20','25','16','46') and vc.tipoVale = 'S',IF(vc.idMoneda = 'USD', vd.VVUnit,vd.VVUnit /  vc.TipoCambio ),0) END ) ValTotalOtrosS,

  ConCat(Vc.IdAlmacen,' ',A.GlsAlmacen) As Almacen,Z.GlsConcepto
  From valescab vc
  Inner join valesdet vd
    on vc.idValesCab = vd.idValesCab
    AND vc.idEmpresa = vd.idEmpresa
    AND vc.idSucursal = vd.idSucursal
    AND vc.tipoVale = vd.tipoVale
  left Join Conceptos Z
    On Vc.IdConcepto = Z.IdConcepto
  Inner Join Almacenes A
    On Vc.IdEmpresa = A.IdEmpresa
    And Vc.IdAlmacen = A.IdAlmacen
  Inner join productos pr
    on vd.idProducto = pr.idProducto
    AND vd.idEmpresa = pr.idEmpresa
  left join personas pe
    on vc.idProvCliente = pe.IdPersona
  Inner join unidadmedida um
    on pr.idUMCompra = um.idUM
  /* Inner Join vw_niveles vn
    On pr.idNivel = vn.idNivel01
    And pr.idEmpresa = vn.idEmpresa
    and pr.idempresa = vn.idempresa*/
  Inner join Niveles n
    on pr.idEmpresa = n.idEmpresa
    and pr.idNivel = n.idNivel
  WHERE vc.idEmpresa = varEmpresa
  AND pr.idEmpresa = varEmpresa
  AND vc.idAlmacen like varAlmacen
  AND (vc.idPeriodoInv) IN (SELECT pi.idPeriodoInv
                            FROM periodosinv pi
                            WHERE pi.idEmpresa = vc.idEmpresa AND pi.idSucursal = vc.idSucursal and pi.FecInicio <= varFechaFin
                            and (pi.FecFin >= varFechaFin or pi.FecFin is null))
  AND vc.estValeCab <> 'ANU'
  AND vc.FechaEmision BETWEEN varFechaIni AND varFechaFin
  AND vd.idProducto like varProducto
--  Group By vn.idNivel02
  group by n.idnivel
  Order by n.GlsNivel ;

END $$

DELIMITER ;