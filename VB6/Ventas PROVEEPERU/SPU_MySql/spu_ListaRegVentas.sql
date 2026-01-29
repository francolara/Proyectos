DELIMITER $$

DROP PROCEDURE IF EXISTS `spu_ListaRegVentas` $$
CREATE PROCEDURE `spu_ListaRegVentas`(

-- Descripcion de version 1 : Reporte Registro de Ventas
-- Fecha de Version 1 : 21/02/2011
-- Creador de version 1 : SAV Peru - Manuel Berrospi
-- Fecha de Pase Version 1 : 28/02/2011
-- Autorizador Version 1 : Vladimir Estrada

varEmpresa         Char(2),
varSucursal        Char(8),
varTipoDoc         Char(2),
varSerie           Char(3),
varMoneda          Char(3),
varFechaIni        VarChar(10),
varFechaFin        VarChar(10),
varOficial         VarChar(250),
VarIdArea          VarChar(8))
BEGIN

DECLARE varGlsTipoDoc VARCHAR(180);
DECLARE varGlsSucursal VARCHAR(180);
DECLARE varGlsSimboloMoneda VARCHAR(180);
DECLARE varGlsMoneda VARCHAR(180);
DECLARE varGlsEmpresa VARCHAR(180);
DECLARE varGlsRuc VARCHAR(180);

SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;

IF varOficial <> '1' THEN
    SET varOficial = '%%';
END IF;

IF varEmpresa <> '' THEN
    SET varGlsEmpresa = (SELECT GlsEmpresa FROM empresas where idempresa = varEmpresa);
END IF;

IF varEmpresa <> '' THEN
    SET varGlsRuc = (SELECT ruc FROM empresas where idempresa = varEmpresa);
END IF;

SET varGlsSimboloMoneda = (SELECT simbolo FROM Monedas WHERE idMoneda = varMoneda);
SET varGlsMoneda = (SELECT GlsMoneda FROM Monedas WHERE idMoneda = varMoneda);


IF varSucursal <> '' THEN
SET varGlsSucursal = (SELECT GlsPersona FROM Personas WHERE idPersona = varSucursal);
ELSE
SET varGlsSucursal = 'TODAS LAS SUCURSALES';
END IF;

IF varTipoDoc <> '' THEN
SET varGlsTipoDoc = (SELECT GlsDocumento FROM documentos WHERE idDocumento = varTipoDoc);
ELSE
SET varGlsTipoDoc = 'TODAS LOS DOCUMENTOS';
END IF;

/*IF varCliente <> '' THEN
SET varCliente = varCliente;
ELSE
SET varCliente = '%%';
END IF;  */

SET varFechaIni = cast(varFechaIni as date);
SET varFechaFin = cast(varFechaFin as date);

SELECT varGlsSucursal AS GlsSucursal, varGlsEmpresa AS GlsEmpresa , varGlsRuc AS GlsRuc,
varGlsTipoDoc AS GlsTipoDoc,
DATE_FORMAT(varFechaIni,'%d/%m/%Y') AS FechaINI,
DATE_FORMAT(varFechaFin,'%d/%m/%Y') AS FECHAFIN,
concat(varSucursal,idDocumento,idSerie,idDocVentas) AS item,
FecEmision,
idDocumento,AbreDocumento, GlsDocumento,
idSerie,idDocVentas,idmoneda,
idPerCliente,
CASE estDocVentas WHEN 'ANU' THEN 'ANULADO' ELSE GlsCliente END AS GlsCliente,
CASE estDocVentas WHEN 'ANU' THEN '' ELSE RUCCliente END AS RUCCliente,
ROUND((IF (idDocumento = '07' Or (IdDocumento = '25' And IndAtribucionNC = 1),TotalBaseImponible * - 1,TotalBaseImponible)) * If(TotalIVAP <> 0,0,1),2) AS baseImponible,
ROUND((IF (idDocumento = '07' Or (IdDocumento = '25' And IndAtribucionNC = 1),TotalBaseImponible * - 1,TotalBaseImponible)) * If(TotalIVAP <> 0,1,0),2) AS baseImponibleIVAP,
ROUND(IF (idDocumento = '07' Or (IdDocumento = '25' And IndAtribucionNC = 1),TotalDsctoVV * - 1,TotalDsctoVV),2) AS dscto,
ROUND(IF (idDocumento = '07' Or (IdDocumento = '25' And IndAtribucionNC = 1),TotalExonerado * - 1,TotalExonerado),2) AS exonerado,
ROUND(IF (idDocumento = '07' Or (IdDocumento = '25' And IndAtribucionNC = 1),TotalIGVVenta * - 1,TotalIGVVenta),2) AS TotalIGVVenta,
ROUND(IF (idDocumento = '07' Or (IdDocumento = '25' And IndAtribucionNC = 1),TotalIVAP * - 1,TotalIVAP),2) AS TotalIVAP,
ROUND(IF (idDocumento = '07' Or (IdDocumento = '25' And IndAtribucionNC = 1),TotalPrecioVenta * - 1,TotalPrecioVenta),2) AS TotalPrecioVenta,
FORMAT(IF (idDocumento = '07' Or (IdDocumento = '25' And IndAtribucionNC = 1),TotalBaseImponible * - 1,TotalBaseImponible),2) AS strbaseImponible,
FORMAT(TotalDsctoVV,2) AS strdscto,
FORMAT(TotalExonerado,2) AS strexonerado,
FORMAT(IF (idDocumento = '07' Or (IdDocumento = '25' And IndAtribucionNC = 1),TotalIGVVenta * - 1,TotalIGVVenta),2) AS strTotalIGVVenta,
FORMAT(IF (idDocumento = '07' Or (IdDocumento = '25' And IndAtribucionNC = 1),TotalPrecioVenta * - 1,TotalPrecioVenta),2) AS strTotalPrecioVenta,
TipoCambio,
idSucursalDoc, GlsSucursalDoc, GlsMoneda as Moneda,
CASE idMoneda WHEN 'USD' THEN IF (idDocumento = '07' Or (IdDocumento = '25' And IndAtribucionNC = 1),TotalPrecioVentaOri * - 1,TotalPrecioVentaOri) ELSE 0 END AS TotalDolares,
CASE idMoneda WHEN 'USD' THEN 'US$' ELSE 'S/.' END AS Moneda,
tipoDocReferencia,serieDocReferencia, numDocReferencia,DescUnidad,GlsUPCliente,ObsRegVentas,GlsFormasPago, idCentroCosto

FROM
(SELECT d.idSucursal AS idSucursalDoc, p.GlsPersona AS GlsSucursalDoc,
d.idMoneda,m.simbolo,m.GlsMoneda,
d.idSerie,d.idDocumento,
AbreDocumento,o.GlsDocumento,d.idDocVentas,
DATE_FORMAT(d.FecEmision,'%d/%m/%Y') AS FecEmision,
idPerCliente,GlsCliente,RUCCliente,
case d.idDocumento when '07' then ifnull(tc.TipoCambio,d.Tipocambio) else t.tcVenta end as TipoCambio,
d.estDocVentas,tc.tipoDocReferencia,tc.serieDocReferencia, tc.numDocReferencia,d.ObsRegVentas,d.GlsFormasPago,
If(D.IndTransGratuita = '1' Or D.IndTransGratuitaMP = '1',0,1) *
IF(d.estDocVentas = 'ANU',0,
CASE varMoneda WHEN 'PEN' THEN IF(d.idMoneda = 'PEN', d.TotalValorVenta,d.TotalValorVenta * if(d.iddocumento <> '07', t.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio)))
WHEN 'USD' THEN IF(d.idMoneda = 'USD', d.TotalValorVenta,d.TotalValorVenta / if(d.iddocumento <> '07', t.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio))) END) AS TotalValorVenta,

If(D.IndTransGratuita = '1' Or D.IndTransGratuitaMP = '1',0,1) *
IF(d.estDocVentas = 'ANU',0,
CASE varMoneda WHEN 'PEN' THEN IF(d.idMoneda = 'PEN', d.TotalIGVVenta,d.TotalIGVVenta * if(d.iddocumento <> '07', t.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio)))
WHEN 'USD' THEN IF(d.idMoneda = 'USD', d.TotalIGVVenta,d.TotalIGVVenta / if(d.iddocumento <> '07', t.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio))) END) AS TotalIGVVenta,

If(D.IndTransGratuita = '1' Or D.IndTransGratuitaMP = '1',0,1) *
IF(d.estDocVentas = 'ANU',0,
CASE varMoneda WHEN 'PEN' THEN IF(d.idMoneda = 'PEN', d.TotalIVAP,d.TotalIVAP * if(d.iddocumento <> '07', t.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio)))
WHEN 'USD' THEN IF(d.idMoneda = 'USD', d.TotalIVAP,d.TotalIVAP / if(d.iddocumento <> '07', t.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio))) END) AS TotalIVAP,

If(D.IndTransGratuita = '1' Or D.IndTransGratuitaMP = '1',0,1) *
IF(d.estDocVentas = 'ANU',0,
CASE varMoneda WHEN 'PEN' THEN IF(d.idMoneda = 'PEN', d.TotalPrecioVenta,d.TotalPrecioVenta * if(d.iddocumento <> '07', t.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio)))
WHEN 'USD' THEN IF(d.idMoneda = 'USD', d.TotalPrecioVenta,d.TotalPrecioVenta / if(d.iddocumento <> '07', t.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio))) END) As TotalPrecioVenta,

If(D.IndTransGratuita = '1' Or D.IndTransGratuitaMP = '1',0,1) *
IF(d.estDocVentas = 'ANU',0,
CASE varMoneda WHEN 'PEN' THEN IF(d.idMoneda = 'PEN', d.TotalBaseImponible,d.TotalBaseImponible * if(d.iddocumento <> '07', t.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio)))
WHEN 'USD' THEN IF(d.idMoneda = 'USD', d.TotalBaseImponible,d.TotalBaseImponible / if(d.iddocumento <> '07', t.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio))) END) AS TotalBaseImponible,

If(D.IndTransGratuita = '1' Or D.IndTransGratuitaMP = '1',0,1) *
IF(d.estDocVentas = 'ANU',0,
CASE varMoneda WHEN 'PEN' THEN IF(d.idMoneda = 'PEN', d.TotalDsctoVV,d.TotalDsctoVV * if(d.iddocumento <> '07', t.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio)))
WHEN 'USD' THEN IF(d.idMoneda = 'USD', d.TotalDsctoVV,d.TotalDsctoVV / if(d.iddocumento <> '07', t.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio))) END) AS TotalDsctoVV,

If(D.IndTransGratuita = '1' Or D.IndTransGratuitaMP = '1',0,1) *
IF(d.estDocVentas = 'ANU',0,
CASE varMoneda WHEN 'PEN' THEN IF(d.idMoneda = 'PEN', d.TotalExonerado,d.TotalExonerado * if(d.iddocumento <> '07', t.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio)))
WHEN 'USD' THEN IF(d.idMoneda = 'USD', d.TotalExonerado,d.TotalExonerado / if(d.iddocumento <> '07', t.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio))) END) As TotalExonerado,
IF(d.estDocVentas = 'ANU',0,d.TotalPrecioVenta) As TotalPrecioVentaOri,ifnull(z.DescUnidad,'') DescUnidad,ifnull(xx.GlsUPCliente,'') GlsUPCliente,
dd.idCentroCosto,d.IndAtribucionNC


FROM docventas d
inner join Documentos o
  on D.idDocumento = o.idDocumento


left join unidadproduccion z
  on d.idempresa = z.idempresa
  and d.idupp = z.CodUnidProd

left join
(
  Select x.idempresa,x.idsucursal,x.iddocumento,x.idserie,x.iddocventas,z.GlsUPCliente
  from Docventas a
  inner join docventasdet x
    on a.idempresa = x.idempresa
    and a.idsucursal = x.idsucursal
    and a.iddocumento = x.iddocumento
    and a.idserie = x.idserie
    and a.iddocventas = x.iddocventas
  inner join unidadproduccioncliente z
    on x.idempresa = z.idempresa
    and x.IdUPCliente = z.IdUPCliente
  WHERE a.idEmpresa = varEmpresa
  AND (a.idSucursal = varSucursal OR varSucursal = '')
  AND (a.idDocumento = varTipoDoc OR varTipoDoc = '')
  AND (a.idSerie = varSerie OR varSerie = '')
  AND (a.IdUPP = VarIdArea Or VarIdArea = '')
  AND a.idDocumento IN ('01','03','07','08','12','90','89','25')
  AND a.FecEmision between varFechaIni AND varFechaFin
  Group by x.idempresa,x.iddocumento,x.idserie,x.iddocventas
) xx
  on d.idempresa = xx.idempresa
  and d.idsucursal = xx.idsucursal
  and d.iddocumento = xx.iddocumento
  and d.idserie = xx.idserie
  and d.iddocventas = xx.iddocventas

inner join Monedas m
  on varMoneda = m.idMoneda

inner join Personas p
  on d.idSucursal = p.idPersona

left join tiposdecambio t
  on d.FecEmision = t.Fecha
  /*(Day(d.FecEmision) = Day(t.fecha)
  AND Year(d.FecEmision) = Year(t.fecha)
  AND Month(d.FecEmision) = Month(t.fecha))*/

left join (select x.tcVenta as tipoCambio,r.idempresa,r.idsucursal,r.tipoDocOrigen,
r.serieDocOrigen, r.numDocOrigen, r.tipoDocReferencia,r.serieDocReferencia, r.numDocReferencia
from docventas dt
inner join docreferencia r
  on dt.idEmpresa = r.idEmpresa
  and dt.idsucursal = r.idsucursal
  and dt.iddocumento = r.tipoDocReferencia
  and dt.idSerie = r.serieDocReferencia
  and dt.idDocVentas = r.numDocReferencia
left join tiposdecambio x
  on dt.FecEmision = x.Fecha
  /*(Day(dt.FecEmision) = Day(x.fecha)
  AND Year(dt.FecEmision) = Year(x.fecha)
  AND Month(dt.FecEmision) = Month(x.fecha))*/
where r.tipoDocOrigen In('07','08')
Group By r.numDocOrigen,r.serieDocOrigen ,r.tipoDocOrigen
Order by r.tipoDocReferencia,r.serieDocReferencia, r.numDocReferencia) tc
  on d.idempresa = tc.idempresa
  and d.idsucursal = tc.idsucursal
  and d.idDocumento = tc.tipoDocOrigen
  and d.idSerie = tc.serieDocOrigen
  and d.idDocVentas = tc.numDocOrigen

Inner Join (
  Select idEmpresa, idSucursal, idDocumento, idSerie, idDocVentas, Replace(Group_Concat(Distinct(idCentroCosto)),',',' - ') As idCentroCosto
  FROM Docventasdet a
  WHERE a.idEmpresa = varEmpresa
  AND (a.idSucursal = varSucursal OR varSucursal = '')
  AND (a.idDocumento = varTipoDoc OR varTipoDoc = '')
  AND (a.idSerie = varSerie OR varSerie = '')
  AND a.idDocumento IN ('01','03','07','08','12','90','89','25','56')
  Group By idDocumento, idSerie, idDocventas, idEmpresa, idSucursal
) dd
  on d.idempresa = dd.idempresa
  And d.idsucursal = dd.idsucursal
  And d.idDocumento = dd.idDocumento
  And d.idSerie = dd.idSerie
  And d.idDocVentas = dd.idDocVentas

WHERE d.idEmpresa = varEmpresa
AND (d.idSucursal = varSucursal OR varSucursal = '')
AND (d.idDocumento = varTipoDoc OR varTipoDoc = '')
AND (d.idSerie = varSerie OR varSerie = '')
AND (d.IdUPP = VarIdArea Or VarIdArea = '')
AND d.idDocumento IN ('01','03','07','08','12','90','89','25','56')
-- AND (d.idPerCliente like varCliente)
AND (o.IndOficial LIKE varOficial)
AND d.FecEmision between varFechaIni AND varFechaFin

)
VENTAS
ORDER BY idSucursalDoc, idDocumento, idSerie, idDocVentas
;
end $$

DELIMITER ;