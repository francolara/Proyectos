DELIMITER $$

DROP PROCEDURE IF EXISTS `spu_ResumenVentasDetCuenta` $$
CREATE PROCEDURE `spu_ResumenVentasDetCuenta`(
varEmpresa   CHAR(2),
varTipo      CHAR(1),
varAno       INTEGER,
varMesDesde  INTEGER,
varMesHasta  INTEGER,
varCliente   VARCHAR(10)
)
BEGIN

DECLARE varGlsEmpresa  VARCHAR(180);
DECLARE varGlsRuc      VARCHAR(180);
DECLARE xCtaContable70 VARCHAR(180);


   IF varEmpresa <> '' THEN
      SET varGlsEmpresa = (SELECT GlsEmpresa FROM empresas where idempresa = varEmpresa);
   END IF;

   IF varEmpresa <> '' THEN
      SET varGlsRuc = (SELECT ruc FROM empresas where idempresa = varEmpresa);
   END IF;

   IF varCliente <> '' THEN
      SET varCliente = varCliente;
   ELSE
      SET varCliente = '%%';
   END IF;

   SET xCtaContable70 = (Select Valparametro From ParametrosCompras where idempresa = varEmpresa and glsparametro = 'TRANS_CONTA_CTA_70');

   SELECT  if(varTipo = '1','EXPRESADO EN AMBAS MONEDA','EXPRESADO EN MONEDA ORIGINAL') as glsMoneda,varMesDesde AS MesDesde,
   varMesHasta AS MesHasta,varGlsEmpresa AS GlsEmpresa , varGlsRuc AS GlsRuc,FecEmision,GlsCliente,
   idDocumento,idSerie,idDocVentas,idPerCliente,idproducto,glsproducto,RUCCliente, idcentrocosto, glscentrocosto,
   round(SUM(IF(idDocumento = '07',Baseimp * - 1,Baseimp)),2) AS Baseimp,
   round(SUM(IF(idDocumento = '07',igv * - 1,igv)),2) AS igv,
   round(SUM(IF(idDocumento = '07',Exonerado * - 1,Exonerado)),2) AS Exonerado,
   round(SUM(IF(idDocumento = '07',Total * - 1,Total)),2) AS Total,
   round(SUM(IF(idDocumento = '07',BaseimpDol * - 1,BaseimpDol)),2) AS BaseimpDol,
   round(SUM(IF(idDocumento = '07',igvDol * - 1,igvDol)),2) AS igvDol,
   round(SUM(IF(idDocumento = '07',ExoneradoDol * - 1,ExoneradoDol)),2) AS ExoneradoDol,
   round(SUM(IF(idDocumento = '07',TotalDol * - 1,TotalDol)),2) AS TotalDol,
   if(CtaContable = '',xCtaContable70,CtaContable) CtaContable
  FROM(
    SELECT d.idSerie,d.idDocumento,d.idDocVentas,o.GlsDocumento,d.FecEmision,d.idPerCliente,d.GlsCliente,dt.idproducto,dt.glsproducto,d.RUCCliente,
    d.idcentrocosto, d.glscentrocosto,
    if(varTipo = '1',
    if(d.idmoneda = 'PEN',if(Afecto = '1',dt.TotalVVNeto,0),if(Afecto = '1',(dt.TotalVVNeto * If(D.IdDocumento <> '07',T.TcVenta,IfNull(Tc.TipoCambio,D.TipoCambio))),0)),
    if(d.idmoneda = 'PEN',if(Afecto = '1',dt.TotalVVNeto,0),0)) as Baseimp,
    if(varTipo = '1',
    if(d.idmoneda = 'PEN',dt.TotalIGVNeto,(dt.TotalIGVNeto * If(D.IdDocumento <> '07',T.TcVenta,IfNull(Tc.TipoCambio,D.TipoCambio)))),
    if(d.idmoneda = 'PEN',dt.TotalIGVNeto,0)) as igv,
    if(varTipo = '1',
    if(d.idmoneda = 'PEN',if(Afecto = '0',dt.TotalVVNeto,0),if(Afecto = '0',(dt.TotalVVNeto * If(D.IdDocumento <> '07',T.TcVenta,IfNull(Tc.TipoCambio,D.TipoCambio))),0)),
    if(d.idmoneda = 'PEN',if(Afecto = '0',dt.TotalVVNeto,0),0)) as Exonerado,
    if(varTipo = '1',
    if(d.idmoneda = 'PEN',dt.TotalPVNeto,dt.TotalPVNeto * If(D.IdDocumento <> '07',T.TcVenta,IfNull(Tc.TipoCambio,D.TipoCambio))),
    if(d.idmoneda = 'PEN',dt.TotalPVNeto,0)) as Total,
    if(varTipo = '1',
    if(d.idmoneda = 'USD',if(Afecto = '1',dt.TotalVVNeto,0),if(Afecto = '1',(dt.TotalVVNeto / If(D.IdDocumento <> '07',T.TcVenta,IfNull(Tc.TipoCambio,D.TipoCambio))),0)),
    if(d.idmoneda = 'USD',if(Afecto = '1',dt.TotalVVNeto,0),0)) as BaseimpDol,
    if(varTipo = '1',
    if(d.idmoneda = 'USD',dt.TotalIGVNeto,(dt.TotalIGVNeto / If(D.IdDocumento <> '07',T.TcVenta,IfNull(Tc.TipoCambio,D.TipoCambio)))),
    if(d.idmoneda = 'USD',dt.TotalIGVNeto,0)) as igvDol,
    if(varTipo = '1',
    if(d.idmoneda = 'USD',if(Afecto = '0',dt.TotalVVNeto,0),if(Afecto = '0',(dt.TotalVVNeto / If(D.IdDocumento <> '07',T.TcVenta,IfNull(Tc.TipoCambio,D.TipoCambio))),0)),
    if(d.idmoneda = 'USD',if(Afecto = '0',dt.TotalVVNeto,0),0)) as ExoneradoDol,
    if(varTipo = '1',
    if(d.idmoneda = 'USD',dt.TotalPVNeto,dt.TotalPVNeto / If(D.IdDocumento <> '07',T.TcVenta,IfNull(Tc.TipoCambio,D.TipoCambio))),
    if(d.idmoneda = 'USD',dt.TotalPVNeto,0)) as TotalDol,ifnull(p.CtaContable,'') CtaContable
    FROM docventas d
    inner join Documentos o
      on D.idDocumento = o.idDocumento
    Left join tiposdecambio t
      on d.FecEmision = t.fecha
    Left Join(
      Select X.TcVenta TipoCambio,R.IdEmpresa,R.IdSucursal,R.TipoDocOrigen,R.SerieDocOrigen,R.NumDocOrigen
      From DocVentas Dt
      Inner Join DocReferencia R
        On Dt.IdEmpresa = R.IdEmpresa And Dt.IdSucursal = R.IdSucursal And Dt.IdDocumento = R.TipoDocReferencia And Dt.IdSerie = R.SerieDocReferencia
        And Dt.IdDocVentas = R.NumDocReferencia
      Inner Join TiposDeCambio X
        On Dt.FecEmision = X.Fecha
      Where R.IdEmpresa = varEmpresa And R.TipoDocOrigen = '07'
      Group By R.NumDocOrigen,R.SerieDocOrigen,R.TipoDocOrigen
      Order by R.TipoDocReferencia,R.SerieDocReferencia,R.NumDocReferencia
    ) Tc
      On D.IdEmpresa = Tc.Idempresa And D.IdSucursal = Tc.IdSucursal And D.IdDocumento = Tc.TipoDocOrigen And D.IdSerie = Tc.SerieDocOrigen
      And D.IdDocVentas = Tc.NumDocOrigen
    inner join docventasdet dt
      on d.idempresa = dt.idempresa And d.iddocumento = dt.iddocumento and d.idserie = dt.idserie and d.iddocventas = dt.iddocventas
    inner join productos p
      on dt.idproducto = p.idproducto and dt.idempresa = p.idempresa
    WHERE d.idEmpresa = varEmpresa AND d.idDocumento IN ('01','03','07','08','12','90','89','56')
    AND month(d.FecEmision) between varMesDesde AND varMesHasta
    and year(d.FecEmision) = varAno and d.estDocVentas <> 'ANU' and d.idpercliente like varCliente
  ) VENTAS
  GROUP BY idPerCliente,idproducto
  ORDER BY idPerCliente,idproducto;

END $$

DELIMITER ;