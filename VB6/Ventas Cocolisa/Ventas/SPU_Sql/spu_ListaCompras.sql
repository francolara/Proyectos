

CREATE OR ALTER PROCEDURE spu_ListaCompras
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

    SELECT @varGlsEmpresa = GlsEmpresa FROM Empresas WHERE idEmpresa = @varEmpresa;

    IF @varSucursal <> '' 
    BEGIN
        SELECT @varGlsSucursal = GlsPersona FROM Personas WHERE idPersona = @varSucursal;
    END
    ELSE 
    BEGIN
        SET @varGlsSucursal = 'TODAS LAS SUCURSALES';
    END

    IF @varAlmacen <> '' 
    BEGIN
        SELECT @varGlsAlmacen = GlsAlmacen FROM Almacenes WHERE IdEmpresa = @varEmpresa AND idAlmacen = @varAlmacen;
    END
    ELSE 
    BEGIN
        SET @varGlsAlmacen = 'TODOS LOS ALMACENES';
    END

    IF @varProveedor <> '' 
    BEGIN
        SELECT @varGlsProveedor = GlsPersona FROM Personas WHERE idPersona = @varProveedor;
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
        c.fechaemision,
        c.idProvCliente,
        p.GlsPersona AS GlsProveedor,
        p.ruc,
        c.GlsDocReferencia,
        c.idmoneda,
        CAST(c.tipocambio AS DECIMAL(14,2)) AS strtipocambio,
        CAST(
            CASE c.idmoneda
                WHEN 'PEN' THEN c.ValorTotal
                ELSE c.ValorTotal * c.tipocambio
            END * CASE WHEN c.TipoVale = 'I' THEN 1 ELSE -1 END AS DECIMAL(14,2)) AS strTotalSoles,
        CAST(
            CASE c.idmoneda
                WHEN 'USD' THEN c.ValorTotal
                ELSE c.ValorTotal / c.tipocambio
            END * CASE WHEN c.TipoVale = 'I' THEN 1 ELSE -1 END AS DECIMAL(14,2)) AS strTotalDolares,
        c.idSucursal,
        c.tipocambio,
        CASE c.idmoneda
            WHEN 'PEN' THEN c.ValorTotal
            ELSE c.ValorTotal * c.tipocambio
        END * CASE WHEN c.TipoVale = 'I' THEN 1 ELSE -1 END AS TotalSoles,
        CASE c.idmoneda
            WHEN 'USD' THEN c.ValorTotal
            ELSE c.ValorTotal / c.tipocambio
        END * CASE WHEN c.TipoVale = 'I' THEN 1 ELSE -1 END AS TotalDolares
    FROM ValesCab c
    INNER JOIN Conceptos o ON c.idConcepto = o.idConcepto
    INNER JOIN Personas p ON c.idProvCliente = p.idPersona
    WHERE c.idEmpresa = @varEmpresa
        AND (c.idSucursal = @varSucursal OR @varSucursal = '')
        AND (c.idAlmacen = @varAlmacen OR @varAlmacen = '')
        AND (c.idProvCliente = @varProveedor OR @varProveedor = '')
        AND CAST(c.fechaemision AS DATE) BETWEEN CAST(@varFechaIni AS DATE) AND CAST(@varFechaFin AS DATE)
        AND o.IndCompra = '1'
        AND c.estValeCab <> 'ANU';
END
GO
