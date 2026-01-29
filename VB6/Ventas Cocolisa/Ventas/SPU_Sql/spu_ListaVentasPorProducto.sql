CREATE OR ALTER PROCEDURE spu_ListaVentasPorProducto
    @varEmpresa         CHAR(2),
    @varSucursal        CHAR(8),
    @varMoneda          CHAR(3),
    @varFechaIni        VARCHAR(20),
    @varFechaFin        VARCHAR(20),
    @varProducto        CHAR(8),
    @varOficial         VARCHAR(250),
    @varNiveles         VARCHAR(250),
    @varGrupo           VARCHAR(250),
    @varOrden           VARCHAR(250)
AS
BEGIN
    DECLARE @varGlsSucursal VARCHAR(180);
    DECLARE @strSQL NVARCHAR(MAX);
    DECLARE @VarGlsEmpresa VARCHAR(250);
    DECLARE @VarGlsRuc VARCHAR(250);

    SELECT @VarGlsEmpresa = GlsEmpresa FROM Empresas WHERE IdEmpresa = @varEmpresa;
    SELECT @VarGlsRuc = Ruc FROM Empresas WHERE IdEmpresa = @varEmpresa;

    IF @varOficial <> '1'
    BEGIN
        SET @varOficial = '%%';
    END

    IF @varSucursal <> ''
    BEGIN
        SELECT @varGlsSucursal = GlsPersona FROM Personas WHERE idPersona = @varSucursal;
    END
    ELSE
    BEGIN
        SET @varGlsSucursal = 'TODAS LAS SUCURSALES';
    END

    SELECT ROW_NUMBER() OVER(ORDER BY GlsProducto) AS Item, 
           @varGlsSucursal  AS GlsSucursal, 
           CONVERT(VARCHAR,CAST(@varFechaIni AS DATE), 103) AS FechaINI,
           CONVERT(VARCHAR,CAST(@varFechaFin AS DATE), 103) AS FECHAFIN,
           simbolo, GlsMoneda, idProducto, GlsProducto, idempresa, idNivel01, GlsNivel01,idNivel02, GlsNivel02,
           abreUM,@VarGlsRuc  AS GlsRuc, 
           @VarGlsEmpresa  AS GlsEmpresa,
           ROUND(TotalValorVenta, 2) AS TotalValorVenta, 
           ROUND(TotalIGVVenta, 2) AS TotalIGVVenta,
           ROUND(TotalPrecioVenta, 2) AS TotalPrecioVenta, 
           ROUND(TotalValorVentaSoles, 2) AS TotalValorVentaSoles, 
           ROUND(TotalValorVentaDolares, 2) AS TotalValorVentaDolares,
           ROUND(TotalIGVVentaSoles, 2) AS TotalIGVVentaSoles, 
           ROUND(TotalIGVVentaDolares, 2) AS TotalIGVVentaDolares,
           ROUND(TotalPrecioVentaSoles, 2) AS TotalPrecioVentaSoles, 
           ROUND(TotalPrecioVentaDolares, 2) AS TotalPrecioVentaDolares,
           Cantidad, TotalVVUnit, PorcentajeVentas
    FROM (
        SELECT  @varGlsSucursal  AS GlsSucursal, 
               CONVERT(VARCHAR,CAST(@varFechaIni AS DATE), 103) AS FechaINI,
               CONVERT(VARCHAR,CAST(@varFechaFin AS DATE), 103) AS FECHAFIN,
               simbolo, GlsMoneda, idProducto, GlsProducto, idempresa, idNivel01, GlsNivel01,idNivel02, GlsNivel02,
               abreUM,@VarGlsRuc  AS GlsRuc,@VarGlsEmpresa  AS GlsEmpresa,
               SUM(CASE WHEN IdDocumento IN ('07', '89') THEN TotalValorVenta * -1 ELSE TotalValorVenta END) AS TotalValorVenta,
               SUM(CASE WHEN IdDocumento IN ('07', '89') THEN TotalIGVVenta * -1 ELSE TotalIGVVenta END) AS TotalIGVVenta,
               SUM(CASE WHEN IdDocumento IN ('07', '89') THEN TotalPrecioVenta * -1 ELSE TotalPrecioVenta END) AS TotalPrecioVenta,
               SUM(CASE WHEN IdDocumento IN ('07', '89') THEN TotalValorVentaSoles * -1 ELSE TotalValorVentaSoles END) AS TotalValorVentaSoles,
               SUM(CASE WHEN IdDocumento IN ('07', '89') THEN TotalValorVentaDolares * -1 ELSE TotalValorVentaDolares END) AS TotalValorVentaDolares,
               SUM(CASE WHEN IdDocumento IN ('07', '89') THEN TotalIGVVentaSoles * -1 ELSE TotalIGVVentaSoles END) AS TotalIGVVentaSoles,
               SUM(CASE WHEN IdDocumento IN ('07', '89') THEN TotalIGVVentaDolares * -1 ELSE TotalIGVVentaDolares END) AS TotalIGVVentaDolares,
               SUM(CASE WHEN IdDocumento IN ('07', '89') THEN TotalPrecioVentaSoles * -1 ELSE TotalPrecioVentaSoles END) AS TotalPrecioVentaSoles,
               SUM(CASE WHEN IdDocumento IN ('07', '89') THEN TotalPrecioVentaDolares * -1 ELSE TotalPrecioVentaDolares END) AS TotalPrecioVentaDolares,
               SUM(CASE WHEN IdDocumento IN ('07', '89') THEN Cantidad * -1 ELSE Cantidad END) AS Cantidad,
               SUM(CASE WHEN IdDocumento IN ('07', '89') THEN TotalVVUnit * -1 ELSE TotalVVUnit END) AS TotalVVUnit,
			   0 PorcentajeVentas
               ---(ISNULL(SUM(TotalValorVenta), 0) * 100 / ISNULL(SUM(TotalPrecioVenta), 0)) AS PorcentajeVentas
        FROM (
            SELECT m.simbolo, m.GlsMoneda, t.idProducto, vn.idempresa, idNivel01, GlsNivel01,idNivel02, GlsNivel02,
                   u.abreUM, d.IdDocumento, C.GlsPersona AS GlsCliente, 
                   (t.GlsProducto + ' ( ' + ltrim(rtrim(t.NumLote)) + ' )') AS GlsProducto,
                   (o.AbreDocumento + d.idSerie + '/' + d.idDocVentas) AS Documento,
                   CONVERT(VARCHAR,d.FecEmision,103) AS FecEmision,

                   CASE @varMoneda
                       WHEN 'PEN' THEN CASE WHEN d.idMoneda = 'PEN' THEN t.TotalVVNeto ELSE t.TotalVVNeto * ISNULL(tc.TipoCambio, d.TipoCambio) END
                       WHEN 'USD' THEN CASE WHEN d.idMoneda = 'USD' THEN t.TotalVVNeto ELSE iif(isnull(ISNULL(tc.TipoCambio, d.TipoCambio),0) = 0,0, t.TotalVVNeto / ISNULL(tc.TipoCambio, d.TipoCambio)) END
                   END AS TotalValorVenta,

                   CASE @varMoneda
                       WHEN 'PEN' THEN CASE WHEN d.idMoneda = 'PEN' THEN t.TotalIGVNeto ELSE t.TotalIGVNeto * ISNULL(tc.TipoCambio, d.TipoCambio) END
                       WHEN 'USD' THEN CASE WHEN d.idMoneda = 'USD' THEN t.TotalIGVNeto ELSE iif(isnull(ISNULL(tc.TipoCambio, d.TipoCambio),0) = 0,0, t.TotalIGVNeto / ISNULL(tc.TipoCambio, d.TipoCambio)) END
                   END AS TotalIGVVenta,

                   CASE @varMoneda
                       WHEN 'PEN' THEN CASE WHEN d.idMoneda = 'PEN' THEN t.TotalPVNeto ELSE t.TotalPVNeto * ISNULL(tc.TipoCambio, d.TipoCambio) END
                       WHEN 'USD' THEN CASE WHEN d.idMoneda = 'USD' THEN t.TotalPVNeto ELSE iif(isnull(ISNULL(tc.TipoCambio, d.TipoCambio),0) = 0,0, t.TotalPVNeto / ISNULL(tc.TipoCambio, d.TipoCambio)) END
                   END AS TotalPrecioVenta,

                   CASE @varMoneda
                       WHEN 'PEN' THEN CASE WHEN d.idMoneda = 'PEN' THEN t.VVUnit ELSE t.VVUnit * ISNULL(tc.TipoCambio, d.TipoCambio) END
                       WHEN 'USD' THEN CASE WHEN d.idMoneda = 'USD' THEN t.VVUnit ELSE iif(isnull(ISNULL(tc.TipoCambio, d.TipoCambio),0) = 0,0, t.VVUnit / ISNULL(tc.TipoCambio, d.TipoCambio)) END
                   END AS TotalVVUnit,

                   CASE WHEN d.idMoneda = 'PEN' THEN t.TotalVVNeto ELSE 0 END AS TotalValorVentaSoles,
                   CASE WHEN d.idMoneda = 'USD' THEN t.TotalVVNeto ELSE 0 END AS TotalValorVentaDolares,
                   CASE WHEN d.idMoneda = 'PEN' THEN t.TotalIGVNeto ELSE 0 END AS TotalIGVVentaSoles,
                   CASE WHEN d.idMoneda = 'USD' THEN t.TotalIGVNeto ELSE 0 END AS TotalIGVVentaDolares,
                   CASE WHEN d.idMoneda = 'PEN' THEN t.TotalPVNeto ELSE 0 END AS TotalPrecioVentaSoles,
                   CASE WHEN d.idMoneda = 'USD' THEN t.TotalPVNeto ELSE 0 END AS TotalPrecioVentaDolares,
                   CASE WHEN t.idTipoProducto = '06002' THEN 0 ELSE t.Cantidad END AS Cantidad

            FROM DocVentas d
            INNER JOIN Personas C ON d.IdPerCliente = C.IdPersona
            INNER JOIN Documentos o ON d.idDocumento = o.idDocumento
            INNER JOIN Monedas m ON @varMoneda = m.idMoneda
            INNER JOIN DocVentasDet t ON d.idDocumento = t.idDocumento AND d.idDocVentas = t.idDocVentas 
                AND d.idSerie = t.idSerie AND d.idEmpresa = t.idEmpresa AND d.idSucursal = t.idSucursal
            INNER JOIN productos p ON t.idProducto = p.idProducto AND t.idEmpresa = p.idEmpresa
            INNER JOIN vw_niveles vn ON p.idNivel = vn.idNivel01 AND p.idEmpresa = vn.idEmpresa
            INNER JOIN unidadmedida u ON t.idUM = u.idUM
            LEFT JOIN tiposdecambio x ON CAST(d.FecEmision AS DATE) = CAST(x.fecha AS DATE)
            LEFT JOIN (
                SELECT x.tcVenta AS TipoCambio, r.idempresa, r.idsucursal, r.tipoDocOrigen, r.serieDocOrigen, r.numDocOrigen
                FROM docventas dt
                INNER JOIN docreferencia r ON dt.idEmpresa = r.idEmpresa AND dt.idsucursal = r.idsucursal 
                    AND dt.iddocumento = r.tipoDocReferencia AND dt.idSerie = r.serieDocReferencia AND dt.idDocVentas = r.numDocReferencia
                LEFT JOIN tiposdecambio x ON dt.FecEmision = x.fecha
                WHERE r.tipoDocOrigen = '07'
            ) tc ON d.idempresa = tc.idempresa 
                AND d.idsucursal = tc.idsucursal 
                AND d.idDocumento = tc.tipoDocOrigen 
                AND d.idSerie = tc.serieDocOrigen 
                AND d.idDocVentas = tc.numDocOrigen
            WHERE d.idEmpresa =  @varEmpresa  
                AND (d.idSucursal =  @varSucursal  OR  @varSucursal  = '')
                AND d.idDocumento IN ('01', '03', '07', '08', '12', '90', '89', '56')
                --AND o.IndOficial LIKE  @varOficial 
                AND (t.idProducto =  @varProducto  OR  @varProducto  = '')
                AND d.estDocVentas <> 'ANU'
                AND CAST(d.FecEmision AS DATE) BETWEEN CAST(@varFechaIni  AS DATE) AND CAST(@varFechaFin  AS DATE)
                AND (d.indVtaGratuita = '' OR d.indVtaGratuita IS NULL)
        ) AS VENTAS
        GROUP BY GlsProducto,simbolo, GlsMoneda, idProducto, idempresa, idNivel01, GlsNivel01,idNivel02, GlsNivel02,
                 abreUM
    ) AS VENTAS
    ORDER BY GlsProducto
END
