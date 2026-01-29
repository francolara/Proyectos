DELIMITER //
CREATE PROCEDURE Spu_DocVentasPDF
(
  VarIdEmpresa CHAR(2),
  VarIdDocumento CHAR(2),
  VarIdSerie CHAR(4),
  VarIdDocVentas CHAR(8)

)
BEGIN
     SELECT
      T3.GlsEmpresa,
      T3.RUC,
      T4.direccion,
      T1.GlsCliente,
      T1.RUCCliente,
      T1.dirCliente,
      T1.FecEmision,
      T1.idMoneda,
      T1.GlsMoneda,

      T2.item,
      T2.idProducto,
      T2.GlsProducto,
      T2.GlsUM,
      T2.Cantidad,
      T2.VVUnit,
      T2.DctoVV,
      T2.TotalVVNeto,

      T1.totalLetras,
      T1.TotalValorVenta,
      T1.TotalIGVVenta,
      T1.TotalPrecioVenta,
      T4.Telefonos,
      T1.idCentroCosto,
      T1.GlsCentroCosto,
      T1.GlsDocReferencia,
      T1.glsFormaPago,
      T3.IdRISunat

  FROM
      docventas T1

      left join docventasdet T2
      on T2.idDocumento = T1.idDocumento and T2.idDocVentas = T1.idDocVentas and T2.idSerie = T1.idSerie

      left join empresas T3
      on T3.idEmpresa = T1.idEmpresa

      left join personas T4
      on T4.idPersona = T3.idPersona

   WHERE
         T1.idEmpresa = VarIdEmpresa
         and T1.idDocumento = VarIdDocumento
         and T1.idSerie = VarIdSerie
         and T1.idDocVentas = VarIdDocVentas;


END //
DELIMITER ;