
CREATE OR ALTER PROCEDURE spu_ListaComprasDet
    @varEmpresa CHAR(2),
    @varSucursal CHAR(8),
    @varAlmacen CHAR(8),
    @varProveedor CHAR(8),
    @varFechaIni VARCHAR(10),
    @varFechaFin VARCHAR(10)
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
        CONCAT(c.tipoVale, ' - ', c.idValesCab) AS idValesCab,
         CONVERT(VARCHAR, CAST(c.fechaemision AS DATE), 103) AS fechaemision,
        c.idProvCliente,
        p.GlsPersona AS GlsProveedor,
        p.ruc,
        c.GlsDocReferencia,
        c.idmoneda,
        d.idProducto,
        d.GlsProducto,
        d.Cantidad,
        u.abreUM,

        c.tipocambio,
        (CASE c.idmoneda
            WHEN 'PEN' THEN d.VVUnit
            ELSE d.VVUnit * c.tipocambio
        END) * CASE WHEN c.TipoVale = 'I' THEN 1 ELSE -1 END AS UnitSoles,

        (CASE c.idmoneda
            WHEN 'USD' THEN d.VVUnit
            ELSE d.VVUnit / c.tipocambio
        END) * CASE WHEN c.TipoVale = 'I' THEN 1 ELSE -1 END AS UnitDolares,

        (CASE c.idmoneda
            WHEN 'PEN' THEN d.TotalVVNeto
            ELSE d.TotalVVNeto * c.tipocambio
        END) * CASE WHEN c.TipoVale = 'I' THEN 1 ELSE -1 END AS TotalSoles,

        (CASE c.idmoneda
            WHEN 'USD' THEN d.TotalVVNeto
            ELSE d.TotalVVNeto / c.tipocambio
        END) * CASE WHEN c.TipoVale = 'I' THEN 1 ELSE -1 END AS TotalDolares

    FROM ValesCab c
    JOIN ValesDet d ON c.idEmpresa = d.idEmpresa AND c.idSucursal = d.idSucursal AND c.idValesCab = d.idValesCab AND c.tipoVale = d.tipoVale
    JOIN Conceptos o ON c.idConcepto = o.idConcepto
    JOIN Personas p ON c.idProvCliente = p.idPersona
    JOIN unidadmedida u ON d.idUM = u.idUM
    WHERE c.idEmpresa = @varEmpresa
        AND (c.idSucursal = @varSucursal OR @varSucursal = '')
        AND (c.idAlmacen = @varAlmacen OR @varAlmacen = '')
        AND (c.idProvCliente = @varProveedor OR @varProveedor = '')
        AND c.idPeriodoInv = (SELECT pi.idPeriodoInv FROM periodosinv pi WHERE pi.estPeriodoInv = 'ACT' AND pi.idEmpresa = c.idEmpresa AND pi.idSucursal = c.idSucursal)
        AND c.fechaemision BETWEEN CAST(@varFechaIni AS DATE) AND CAST(@varFechaFin AS DATE)
        AND o.indCosto = 'S'
        AND c.estValeCab <> 'ANU';
END
GO
