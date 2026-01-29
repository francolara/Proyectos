
CREATE OR ALTER PROCEDURE spu_ListaComprasPorProducto
    @varEmpresa CHAR(2),
    @varSucursal CHAR(8),
    @varAlmacen CHAR(8),
    @varProveedor CHAR(8),
    @varFechaIni VARCHAR(20),
    @varFechaFin VARCHAR(20)
AS
BEGIN
    DECLARE @varGlsEmpresa VARCHAR(180);
    DECLARE @varGlsSucursal VARCHAR(180);
    DECLARE @varGlsAlmacen VARCHAR(180);
    DECLARE @varGlsProveedor VARCHAR(180);

    SET @varGlsEmpresa = (SELECT GlsEmpresa FROM Empresas WHERE idEmpresa = @varEmpresa);

    IF @varSucursal <> '' 
    BEGIN
        SET @varGlsSucursal = (SELECT GlsPersona FROM Personas WHERE idPersona = @varSucursal);
    END
    ELSE 
    BEGIN
        SET @varGlsSucursal = 'TODAS LAS SUCURSALES';
    END

    IF @varAlmacen <> '' 
    BEGIN
        SET @varGlsAlmacen = (SELECT GlsAlmacen FROM Almacenes WHERE IdEmpresa = @varEmpresa AND idAlmacen = @varAlmacen);
    END
    ELSE 
    BEGIN
        SET @varGlsAlmacen = 'TODOS LOS ALMACENES';
    END

    IF @varProveedor <> '' 
    BEGIN
        SET @varGlsProveedor = (SELECT GlsPersona FROM Personas WHERE idPersona = @varProveedor);
    END
    ELSE 
    BEGIN
        SET @varGlsProveedor = 'TODOS LOS PROVEEDORES';
    END

    SELECT 
        @varGlsEmpresa AS GlsEmpresa,
        @varGlsSucursal AS GlsSucursal,
        @varGlsAlmacen AS GlsAlmacen,
        @varGlsProveedor AS GlsProveedorEtiqueta,
        CONVERT(VARCHAR, CAST(@varFechaIni AS DATE), 103) AS FechaINI,
        CONVERT(VARCHAR, CAST(@varFechaFin AS DATE), 103) AS FECHAFIN,
        c.idValesCab,
        c.fechaemision,
        c.idProvCliente,
        p.GlsPersona AS GlsProveedor,
        p.ruc,
        c.GlsDocReferencia,
        c.idmoneda,
        CAST(c.tipocambio AS DECIMAL(14,2)) AS strtipocambio,

        SUM(d.Cantidad * CASE WHEN c.TipoVale = 'I' THEN 1 ELSE -1 END) AS Cantidad,
        po.idproducto,
        po.glsproducto,
        u.abreUM,

        CAST(SUM(
            CASE c.idmoneda
                WHEN 'PEN' THEN d.VVUnit
                ELSE d.VVUnit * c.tipocambio
            END * CASE WHEN c.TipoVale = 'I' THEN 1 ELSE -1 END
        ) AS DECIMAL(14,2)) AS strUniSoles,

        CAST(SUM(
            CASE c.idmoneda
                WHEN 'USD' THEN d.VVUnit
                ELSE d.VVUnit / c.tipocambio
            END * CASE WHEN c.TipoVale = 'I' THEN 1 ELSE -1 END
        ) AS DECIMAL(14,2)) AS strUniDolares,

        c.idSucursal,
        c.tipocambio,

        SUM(
            CASE c.idmoneda
                WHEN 'PEN' THEN (d.cantidad * d.VVUnit) * CASE WHEN c.TipoVale = 'I' THEN 1 ELSE -1 END
                ELSE ((d.cantidad * d.VVUnit) * c.tipocambio) * CASE WHEN c.TipoVale = 'I' THEN 1 ELSE -1 END
            END
        ) AS TotalSoles,

        SUM(
            CASE c.idmoneda
                WHEN 'USD' THEN (d.cantidad * d.VVUnit) * CASE WHEN c.TipoVale = 'I' THEN 1 ELSE -1 END
                ELSE ((d.cantidad * d.VVUnit) / c.tipocambio) * CASE WHEN c.TipoVale = 'I' THEN 1 ELSE -1 END
            END
        ) AS TotalDolares

    FROM ValesCab c
    INNER JOIN ValesDet d ON c.idvalescab = d.idvalescab AND c.idempresa = d.idempresa AND c.tipovale = d.tipovale AND c.idsucursal = d.idsucursal
    INNER JOIN Productos po ON d.idempresa = po.idempresa AND d.idproducto = po.idproducto
    INNER JOIN Conceptos o ON c.idConcepto = o.idConcepto
    LEFT JOIN Personas p ON c.idProvCliente = p.idPersona
    INNER JOIN UnidadMedida u ON po.idUmCompra = u.idUM
    WHERE c.idEmpresa = @varEmpresa
        AND (c.idSucursal = @varSucursal OR @varSucursal = '')
        AND (c.idAlmacen = @varAlmacen OR @varAlmacen = '')
        AND (c.idProvCliente = @varProveedor OR @varProveedor = '')
        AND o.indCosto = 'S'
        AND c.estValeCab <> 'ANU'
        AND c.fechaemision BETWEEN @varFechaIni AND @varFechaFin
    GROUP BY c.idValesCab,
        c.fechaemision,
        c.idProvCliente,
        p.GlsPersona,
        p.ruc,
        c.GlsDocReferencia,
        c.idmoneda,
		c.tipocambio,
		po.idproducto,
        po.glsproducto,
        u.abreUM,
		c.idSucursal
    ORDER BY TotalSoles;
END
GO
