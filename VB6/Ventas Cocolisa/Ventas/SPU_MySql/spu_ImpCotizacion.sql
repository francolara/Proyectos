DELIMITER $$

DROP PROCEDURE IF EXISTS `spu_ImpCotizacion` $$
CREATE PROCEDURE `spu_ImpCotizacion`(
varEmpresa     CHAR(2),
varSucursal    CHAR(8),
varTipoDoc     CHAR(2),
varSerie       CHAR(4),
varDocVentas   CHAR(8)
)
BEGIN

DECLARE varGlsRuc     VARCHAR(180);
DECLARE varGlsEmpresa VARCHAR(180);

SET varGlsEmpresa = (SELECT glsempresa FROM empresas WHERE idempresa = varEmpresa);
SET varGlsRuc = (SELECT Ruc FROM empresas WHERE idempresa = varEmpresa);

SELECT  d.item,d.GlsProducto,d.GlsUM,d.Cantidad,d.VVUnit,cast(d.PorDcto as unsigned) as  PorDcto ,d.DctoVV ,d.TotalVVNeto,
c.TotalValorventa,c.TotalIgvVenta,c.TotalPrecioVenta,c.llegada,c.llegada2,c.idmoneda,
c.GlsFormaPago,c.GlsCliente,c.RUCCliente,c.ObsDocVentas,c.idDocventas, DATE_FORMAT(c.FecEmision,'%d/%m/%Y')  AS FecEmision,
d.VVUnitNeto,c.GlsVendedor,
varGlsRuc as RucEmpresa , varGlsEmpresa as GlsEmpresa,
(SELECT concat(direccion,' ',(select glsubigeo from ubigeo where iddistrito = p.iddistrito ))
FROM personas p  WHERE p.idpersona='08090004')AS Sucursal,
(SELECT concat(direccion,' ',(select glsubigeo from ubigeo where iddistrito = p.iddistrito ))
FROM personas p  WHERE p.idpersona=C.idsucursal)AS DireccionFiscal,
(SELECT  TELEFONOS FROM PERSONAS WHERE IDPERSONA=C.IDSUCURSAL)AS TelfEmpresa,
(select Telefonos from personas where idpersona=c.idcontacto) as TelefonosContacto,
(select mail from personas where idpersona=c.idcontacto) as mailContacto,
(select Glspersona from personas where idpersona=c.idcontacto) as GlsContacto,
(select GlsPersona from personas where idPersona = c.idSucursal) as DireccionComercial,
(select Telefonos from personas where idPersona = c.idPerVendedor) as TelefonoVendedor,
(select Mail from personas where idPersona = c.idPerVendedor) as MailVendedor,
(select Nextel from Vendedores where idVendedor = c.idPerVendedor And idEmpresa = varEmpresa) as NextelVendedor,
(select Rpm from Vendedores Where idVendedor = c.idPerVendedor And idEmpresa = varEmpresa ) as RpmVendedor,
(select valParametro from parametros where glsParametro = 'IGV'AND IDEMPRESA=C.IDEMPRESA) as IGV,
(select valParametro from parametros where glsParametro = 'NUMERO_CUENTAS_CORRIENTES') as CTA_CORRIENTES,
CASE c.idMoneda WHEN 'PEN' THEN 'S/.' ELSE 'US$' END AS MONEDAS
FROM Docventas c
INNER Join docventasdet d
  ON  c.idEmpresa = d.idEmpresa
  AND c.idSucursal = d.idSucursal
  AND c.idDocumento = d.idDocumento
  AND c.idSerie = d.idSerie
  AND c.idDocVentas = d.idDocVentas
INNER JOIN personas p
  ON c.idPerCliente = p.idPersona
INNER JOIN empresas e
  On c.idEmpresa=e.idEmpresa
WHERE  c.idEmpresa = varEmpresa
  AND c.idSucursal = varSucursal
  AND c.idSerie = varSerie
  AND c.idDocumento = varTipoDoc
  AND c.idDocVentas = varDocVentas
ORDER BY d.item ;

END $$

DELIMITER ;