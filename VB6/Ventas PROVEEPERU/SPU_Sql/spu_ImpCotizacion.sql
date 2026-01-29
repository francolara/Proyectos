CREATE OR ALTER PROCEDURE dbo.spu_ImpCotizacion
    @varEmpresa CHAR(2),
    @varSucursal CHAR(8),
    @varTipoDoc CHAR(2),
    @varSerie CHAR(4),
    @varDocVentas CHAR(8)
AS
BEGIN
    DECLARE @varGlsRuc VARCHAR(180);
    DECLARE @varGlsEmpresa VARCHAR(180);

    SELECT @varGlsEmpresa = glsempresa, @varGlsRuc = Ruc
    FROM empresas
    WHERE idempresa = @varEmpresa;

    SELECT
        d.item,
        d.GlsProducto,
        u.GlsUM AS  GlsUM,
        d.Cantidad,
        d.VVUnit,
        CAST(d.PorDcto AS INT) AS PorDcto,
        d.DctoVV,
        d.TotalVVNeto,
        c.TotalValorventa,
        c.TotalIgvVenta,
        c.TotalPrecioVenta,
        c.llegada,
        c.llegada2,
        c.idmoneda,
        c.GlsFormaPago,
        c.GlsCliente,
        c.RUCCliente,
        c.ObsDocVentas,
        C.idDocventas + '-' + c.idSerie AS idDocventas,
        dbo.FechaEnLetras(CAST(c.FecEmision AS DATE))  AS FecEmision,
        d.VVUnitNeto,
        c.GlsVendedor,
        @varGlsRuc AS RucEmpresa,
        @varGlsEmpresa AS GlsEmpresa,
        (SELECT CONCAT(direccion, ' ', (SELECT glsubigeo FROM ubigeo WHERE iddistrito = p.iddistrito))
         FROM personas p
         WHERE p.idpersona = '08090004') AS Sucursal,
        (SELECT CONCAT(direccion, ' ', (SELECT glsubigeo FROM ubigeo WHERE iddistrito = p.iddistrito))
         FROM personas p
         WHERE p.idpersona = C.idsucursal) AS DireccionFiscal,
        (SELECT TELEFONOS FROM PERSONAS WHERE IDPERSONA = C.IDSUCURSAL) AS TelfEmpresa,
        (SELECT Telefonos FROM personas WHERE idpersona = c.idcontacto) AS TelefonosContacto,
        (SELECT mail FROM personas WHERE idpersona = c.idcontacto) AS mailContacto,
        (SELECT Glspersona FROM personas WHERE idpersona = c.idcontacto) AS GlsContacto,
        (SELECT GlsPersona FROM personas WHERE idPersona = c.idSucursal) AS DireccionComercial,

        (SELECT Telefonos FROM personas WHERE idPersona = c.idUsuarioReg) AS TelefonoVendedor,
        (SELECT Mail FROM personas WHERE idPersona = c.idUsuarioReg) AS MailVendedor,
        (SELECT Nextel FROM Vendedores WHERE idVendedor = c.idUsuarioReg AND idEmpresa = @varEmpresa) AS NextelVendedor,
        (SELECT Rpm FROM Vendedores WHERE idVendedor = c.idUsuarioReg AND idEmpresa = @varEmpresa) AS RpmVendedor,

        (SELECT valParametro FROM parametros WHERE glsParametro = 'IGV' AND IDEMPRESA = C.IDEMPRESA) AS IGV,
        (SELECT valParametro FROM parametros WHERE glsParametro = 'NUMERO_CUENTAS_CORRIENTES') AS CTA_CORRIENTES,
        CASE c.idMoneda
            WHEN 'PEN' THEN 'S/.'
            ELSE 'US$'
        END AS MONEDAS,
		c.Partida,
		c.Modelo
    FROM Docventas c
    INNER JOIN docventasdet d ON c.idEmpresa = d.idEmpresa
                               AND c.idSucursal = d.idSucursal
                               AND c.idDocumento = d.idDocumento
                               AND c.idSerie = d.idSerie
                               AND c.idDocVentas = d.idDocVentas
    INNER JOIN personas p ON c.idPerCliente = p.idPersona
	INNER JOIN unidadmedida u ON d.idUM = u.idUM

    WHERE c.idEmpresa = @varEmpresa
      AND c.idSucursal = @varSucursal
      AND c.idSerie = @varSerie
      AND c.idDocumento = @varTipoDoc
      AND c.idDocVentas = @varDocVentas
    ORDER BY d.item;
END;
