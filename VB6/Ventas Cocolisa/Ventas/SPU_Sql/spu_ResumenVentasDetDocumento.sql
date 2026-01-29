CREATE OR ALTER PROCEDURE spu_ResumenVentasDetDocumento
    @varEmpresa CHAR(2),
    @varTipo CHAR(1),
    @varAno INT,
    @varMesDesde INT,
    @varMesHasta INT,
    @varCliente VARCHAR(10)
AS
BEGIN
    DECLARE @varGlsEmpresa VARCHAR(180);
    DECLARE @varGlsRuc VARCHAR(180);

    IF @varEmpresa <> ''
    BEGIN
        SELECT @varGlsEmpresa = GlsEmpresa FROM empresas WHERE idempresa = @varEmpresa;
    END

    IF @varEmpresa <> ''
    BEGIN
        SELECT @varGlsRuc = ruc FROM empresas WHERE idempresa = @varEmpresa;
    END

    IF @varCliente = ''
    BEGIN
        SET @varCliente = '%%';
    END

    SELECT 
        CASE WHEN @varTipo = '1' THEN 'EXPRESADO EN AMBAS MONEDA' ELSE 'EXPRESADO EN MONEDA ORIGINAL' END AS glsMoneda,
        @varMesDesde AS MesDesde,
        @varMesHasta AS MesHasta,
        @varGlsEmpresa AS GlsEmpresa,
        @varGlsRuc AS GlsRuc,
        FecEmision,
        GlsCliente,
        abredocumento,
        glsdocumento,
        idDocumento,
        idSerie,
        idDocVentas,
        idPerCliente,
        idproducto,
        glsproducto,
        RUCCliente,
        glsdescproducto,
        CONCAT(abredocumento, idSerie, '/', idDocVentas) AS Documento,
        ROUND((CASE WHEN idDocumento = '07' THEN Baseimp * -1 ELSE Baseimp END), 2) AS Baseimp,
        ROUND((CASE WHEN idDocumento = '07' THEN igv * -1 ELSE igv END), 2) AS igv,
        ROUND((CASE WHEN idDocumento = '07' THEN Exonerado * -1 ELSE Exonerado END), 2) AS Exonerado,
        ROUND((CASE WHEN idDocumento = '07' THEN Total * -1 ELSE Total END), 2) AS Total,
        ROUND((CASE WHEN idDocumento = '07' THEN BaseimpDol * -1 ELSE BaseimpDol END), 2) AS BaseimpDol,
        ROUND((CASE WHEN idDocumento = '07' THEN igvDol * -1 ELSE igvDol END), 2) AS igvDol,
        ROUND((CASE WHEN idDocumento = '07' THEN ExoneradoDol * -1 ELSE ExoneradoDol END), 2) AS ExoneradoDol,
        ROUND((CASE WHEN idDocumento = '07' THEN TotalDol * -1 ELSE TotalDol END), 2) AS TotalDol
    FROM (
        SELECT 
            d.idSerie,
            d.idDocumento,
            d.idDocVentas,
            o.GlsDocumento,
            d.FecEmision,
            o.abredocumento,
            d.idPerCliente,
            d.GlsCliente,
            dt.idproducto,
            (dt.glsproducto + ' ( ' + ltrim(rtrim(dt.NumLote)) + ' )') AS glsproducto,
            d.RUCCliente,
            p.glsproducto AS glsdescproducto,
			iif(@varTipo = '1',
			iif(d.idmoneda = 'PEN',iif(Afecto = '1',dt.TotalVVNeto,0),iif(Afecto = '1',(dt.TotalVVNeto * iif(D.IdDocumento <> '07',T.TcVenta,IsNull(Tc.TipoCambio,D.TipoCambio))),0)),
			iif(d.idmoneda = 'PEN',iif(Afecto = '1',dt.TotalVVNeto,0),0)) as Baseimp,
			iif(@varTipo = '1',
			iif(d.idmoneda = 'PEN',dt.TotalIGVNeto,(dt.TotalIGVNeto * iif(D.IdDocumento <> '07',T.TcVenta,IsNull(Tc.TipoCambio,D.TipoCambio)))),
			iif(d.idmoneda = 'PEN',dt.TotalIGVNeto,0)) as igv,
			iif(@varTipo = '1',
			iif(d.idmoneda = 'PEN',iif(Afecto = '0',dt.TotalVVNeto,0),iif(Afecto = '0',(dt.TotalVVNeto * iif(D.IdDocumento <> '07',T.TcVenta,IsNull(Tc.TipoCambio,D.TipoCambio))),0)),
			iif(d.idmoneda = 'PEN',iif(Afecto = '0',dt.TotalVVNeto,0),0)) as Exonerado,
			iif(@varTipo = '1',
			iif(d.idmoneda = 'PEN',dt.TotalPVNeto,dt.TotalPVNeto * iif(D.IdDocumento <> '07',T.TcVenta,IsNull(Tc.TipoCambio,D.TipoCambio))),
			iif(d.idmoneda = 'PEN',dt.TotalPVNeto,0)) as Total,
			iif(@varTipo = '1',
			iif(d.idmoneda = 'USD',iif(Afecto = '1',dt.TotalVVNeto,0),iif(Afecto = '1',(dt.TotalVVNeto / iif(D.IdDocumento <> '07',T.TcVenta,IsNull(Tc.TipoCambio,D.TipoCambio))),0)),
			iif(d.idmoneda = 'USD',iif(Afecto = '1',dt.TotalVVNeto,0),0)) as BaseimpDol,
			iif(@varTipo = '1',
			iif(d.idmoneda = 'USD',dt.TotalIGVNeto,(dt.TotalIGVNeto / iif(D.IdDocumento <> '07',T.TcVenta,IsNull(Tc.TipoCambio,D.TipoCambio)))),
			iif(d.idmoneda = 'USD',dt.TotalIGVNeto,0)) as igvDol,
			iif(@varTipo = '1',
			iif(d.idmoneda = 'USD',iif(Afecto = '0',dt.TotalVVNeto,0),iif(Afecto = '0',(dt.TotalVVNeto / iif(D.IdDocumento <> '07',T.TcVenta,IsNull(Tc.TipoCambio,D.TipoCambio))),0)),
			iif(d.idmoneda = 'USD',iif(Afecto = '0',dt.TotalVVNeto,0),0)) as ExoneradoDol,
			iif(@varTipo = '1',
			iif(d.idmoneda = 'USD',dt.TotalPVNeto,dt.TotalPVNeto / iif(D.IdDocumento <> '07',T.TcVenta,IsNull(Tc.TipoCambio,D.TipoCambio))),
			iif(d.idmoneda = 'USD',dt.TotalPVNeto,0)) as TotalDol

    FROM 
        docventas d
    INNER JOIN 
        Documentos o ON d.idDocumento = o.idDocumento
    LEFT JOIN 
        tiposdecambio t ON d.FecEmision = t.fecha
    LEFT JOIN 
        (SELECT X.TcVenta AS TipoCambio,
                R.IdEmpresa,
                R.IdSucursal,
                R.TipoDocOrigen,
                R.SerieDocOrigen,
                R.NumDocOrigen
         FROM 
                DocVentas Dt
         INNER JOIN 
                DocReferencia R ON Dt.IdEmpresa = R.IdEmpresa
                               AND Dt.IdSucursal = R.IdSucursal
                               AND Dt.IdDocumento = R.TipoDocReferencia
                               AND Dt.IdSerie = R.SerieDocReferencia
                               AND Dt.IdDocVentas = R.NumDocReferencia
         INNER JOIN 
                TiposDeCambio X ON Dt.FecEmision = X.Fecha
         WHERE 
                R.IdEmpresa = @varEmpresa
                AND R.TipoDocOrigen = '07'
         GROUP BY X.TcVenta,R.IdEmpresa,
                R.IdSucursal,
                R.NumDocOrigen,
                R.SerieDocOrigen,
                R.TipoDocOrigen) Tc ON D.IdEmpresa = Tc.Idempresa
                                        AND D.IdSucursal = Tc.IdSucursal
                                        AND D.IdDocumento = Tc.TipoDocOrigen
                                        AND D.IdSerie = Tc.SerieDocOrigen
                                        AND D.IdDocVentas = Tc.NumDocOrigen
    INNER JOIN 
        docventasdet dt ON d.idempresa = dt.idempresa
                         AND d.iddocumento = dt.iddocumento
                         AND d.idserie = dt.idserie
                         AND d.iddocventas = dt.iddocventas
    INNER JOIN 
        productos p ON p.idempresa = dt.idempresa
                    AND p.idproducto = dt.idproducto
    WHERE 
        d.idEmpresa = @varEmpresa
        AND d.idDocumento IN ('01', '03', '07', '08', '12', '90', '89', '56')
        AND MONTH(d.FecEmision) BETWEEN @varMesDesde AND @varMesHasta
        AND YEAR(d.FecEmision) = @varAno
        AND d.estDocVentas <> 'ANU'
        AND (d.idpercliente LIKE @varCliente OR @varCliente = '%%')
) VENTAS
ORDER BY 
    iddocumento,
    idserie,
    iddocventas;

END
