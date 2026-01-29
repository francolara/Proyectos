
CREATE OR ALTER PROCEDURE spu_ListaComprasDetGrilla
    @varEmpresa CHAR(2),
    @varSucursal CHAR(8),
    @varIdValeCab CHAR(8),
    @varTipoValeCab CHAR(8)
AS
BEGIN
    SELECT 
        c.idValesCab,
        d.item,
        d.idProducto,
        d.GlsProducto,
        d.Cantidad,
        u.abreUM,

        CAST(
            CASE c.idmoneda
                WHEN 'PEN' THEN d.VVUnit
                ELSE d.VVUnit * c.tipocambio
            END AS DECIMAL(14,2)) AS strUnitSoles,

        CAST(
            CASE c.idmoneda
                WHEN 'USD' THEN d.VVUnit
                ELSE d.VVUnit / c.tipocambio
            END AS DECIMAL(14,2)) AS strUnitDolares,

        CAST(
            CASE c.idmoneda
                WHEN 'PEN' THEN d.TotalVVNeto
                ELSE d.TotalVVNeto * c.tipocambio
            END AS DECIMAL(14,2)) AS strTotalSoles,

        CAST(
            CASE c.idmoneda
                WHEN 'USD' THEN d.TotalVVNeto
                ELSE d.TotalVVNeto / c.tipocambio
            END AS DECIMAL(14,2)) AS strTotalDolares

    FROM ValesCab c
    INNER JOIN ValesDet d
        ON c.idEmpresa = d.idEmpresa
        AND c.idSucursal = d.idSucursal
        AND c.idValesCab = d.idValesCab
        AND c.tipoVale = d.TipoVale
    INNER JOIN unidadmedida u
        ON d.idUM = u.idUM
    WHERE c.idEmpresa = @varEmpresa
        AND c.idSucursal = @varSucursal
        AND c.idValesCab = @varIdValeCab
        AND c.tipovale = @varTipoValeCab
    ORDER BY d.item;
END
GO
