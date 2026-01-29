
CREATE OR ALTER PROCEDURE spu_ImpPedido
    @varEmpresa CHAR(2),
    @varSucursal CHAR(8),
    @varTipoDoc CHAR(2),
    @varSerie CHAR(4),
    @varDocVentas CHAR(8)
AS
BEGIN
    SELECT 
        c.idDocVentas,c.GlsCliente,c.RUCCliente,CONVERT(VARCHAR,c.FecEmision,103) as FecEmision,c.TotalPrecioVenta,c.TipoCambio,c.GlsDocReferencia,c.idSerie,c.ObsDocVentas,c.GlsVendedor,
		d.idProducto,CAST(d.GlsProducto AS VARCHAR(500)) AS GlsProducto,d.Cantidad,d.PVUnit,d.PorDcto,d.TotalPVNeto,
		d.GlsUM,
		CASE c.idMoneda 
            WHEN 'USD' THEN 'US$' 
            ELSE 'S/.' 
        END AS Moneda
    FROM 
        docventas c
    INNER JOIN 
        docventasdet d
    ON 
        c.idEmpresa = d.idEmpresa
        AND c.idSucursal = d.idSucursal
        AND c.idDocumento = d.idDocumento
        AND c.idSerie = d.idSerie
        AND c.idDocVentas = d.idDocVentas
    WHERE 
        c.idEmpresa = @varEmpresa
        AND c.idSucursal = @varSucursal
        AND c.idDocumento = @varTipoDoc
        AND c.idSerie = @varSerie
        AND c.idDocVentas = @varDocVentas;
END
GO
