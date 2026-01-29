DELIMITER $$

DROP PROCEDURE IF EXISTS `spu_ImpPedido` $$
CREATE PROCEDURE `spu_ImpPedido`(

varEmpresa     CHAR(2),

varSucursal    CHAR(8),

varTipoDoc     CHAR(2),

varSerie       CHAR(4),

varDocVentas   CHAR(8)

)
BEGIN



SELECT *, case c.idMoneda When 'USD' then 'US$' else 'S/.' end as Moneda

FROM docventas c, docventasdet d

WHERE c.idEmpresa = d.idEmpresa

  AND c.idSucursal = d.idSucursal

  AND c.idDocumento = d.idDocumento

  AND c.idSerie = d.idSerie

  AND c.idDocVentas = d.idDocVentas

  AND c.idEmpresa = varEmpresa

  AND c.idSucursal = varSucursal

  AND c.idDocumento = varTipoDoc

  AND c.idSerie = varSerie

  AND c.idDocVentas = varDocVentas;



END $$

DELIMITER ;