

CREATE OR ALTER PROCEDURE spu_ImpOC
    @varEmpresa CHAR(2),
    @varSucursal CHAR(8),
    @varTipoDoc CHAR(2),
    @varSerie CHAR(3),
    @varDocVentas CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    SELECT  c.idDocVentas,c.GlsCliente, c.RUCCliente,c.dirCliente,CONVERT(VARCHAR,C.FecEmision,103) AS FecEmision,c.TotalValorVenta,
			c.TotalIGVVenta,c.TotalPrecioVenta,c.TipoCambio,c.ObsDocVentas,c.Partida,c.llegada,c.GlsVendedor,
			c.idCentroCosto,c.GlsCentroCosto,c.glsFormaPago,c.GlsMoneda,c.GlsContacto,
	
			D.Cantidad,D.GlsProducto,D.idProducto,d.GlsUM,D.item,D.VVUnit,D.PorDcto,D.TotalVVNeto,
           CASE c.idMoneda 
               WHEN 'USD' THEN 'US$' 
               ELSE 'S/.' 
           END AS Moneda,
           (SELECT direccion + ' ' + (SELECT glsubigeo FROM ubigeo WHERE iddistrito = p.iddistrito)
            FROM personas p  
            WHERE p.idpersona = c.idsucursal) AS Direccion,
           (SELECT TELEFONOS  
            FROM PERSONAS 
            WHERE IDPERSONA = c.IDSUCURSAL) AS TELEFONOS,
           (SELECT Telefonos  
            FROM personas 
            WHERE idpersona = c.idPercliente) AS TelfContacto,
           (SELECT mail       
            FROM personas 
            WHERE idpersona = c.idPercliente) AS mailContacto,
           (SELECT Glspersona 
            FROM personas 
            WHERE idpersona = c.idcontacto) AS GlsContacto1,
           (SELECT Glspersona 
            FROM personas   
            WHERE idPersona = c.idPerVendedor) AS GlsComprador,
           (SELECT Telefonos  
            FROM personas   
            WHERE idPersona = c.idPerVendedor) AS TelefonoComprador,
           (SELECT Mail       
            FROM personas   
            WHERE idPersona = c.idPerVendedor) AS MailComprador,
           (SELECT GlsEmpresa FROM empresas where idempresa  = @varEmpresa) AS GlsEmpresa,
           (SELECT RUC FROM empresas where idempresa  = @varEmpresa) AS RUC

    FROM docventas c
    INNER JOIN docventasdet d ON c.idEmpresa = d.idEmpresa
                             AND c.idSucursal = d.idSucursal
                             AND c.idDocumento = d.idDocumento
                             AND c.idSerie = d.idSerie
                             AND c.idDocVentas = d.idDocVentas
    INNER JOIN Monedas m ON c.IdMoneda = m.idMoneda
    INNER JOIN Empresas e ON e.idEmpresa = c.idEmpresa
                         AND e.idEmpresa = d.idEmpresa
    INNER JOIN FormasPagos fp ON c.idFormaPago = fp.idFormaPago
                              AND c.idempresa = fp.idempresa
    WHERE c.idEmpresa = @varEmpresa
      AND c.idSucursal = @varSucursal
      AND c.idDocumento = @varTipoDoc
      AND c.idSerie = @varSerie
      AND c.idDocVentas = @varDocVentas;

    SET NOCOUNT OFF;
END
GO
