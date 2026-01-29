CREATE OR ALTER PROCEDURE spu_ResumenVentas
    @varEmpresa CHAR(2),
    @varTipo CHAR(1),
    @varAno INTEGER,
    @varMesDesde INTEGER,
    @varMesHasta INTEGER,
    @varCliente VARCHAR(10)
AS
BEGIN
    DECLARE @varGlsEmpresa VARCHAR(180);
    DECLARE @varGlsRuc VARCHAR(180);

    IF @varEmpresa <> '' 
    BEGIN
        SELECT @varGlsEmpresa = GlsEmpresa 
        FROM empresas 
        WHERE idempresa = @varEmpresa;
    END

    IF @varEmpresa <> '' 
    BEGIN
        SELECT @varGlsRuc = ruc 
        FROM empresas 
        WHERE idempresa = @varEmpresa;
    END

    IF @varCliente = '' 
    BEGIN
        SET @varCliente = '%%';
    END

    SELECT  
        @varGlsEmpresa AS GlsEmpresa, 
        @varGlsRuc AS GlsRuc,
        --FecEmision,
        GlsCliente,
        --idDocumento,
        --idSerie,
        --idDocVentas,
        idPerCliente,
        --idproducto,
        --glsproducto,
        RUCCliente,
        ROUND(SUM(CASE WHEN idDocumento = '07' THEN Baseimp * -1 ELSE Baseimp END), 2) AS Baseimp,
        ROUND(SUM(CASE WHEN idDocumento = '07' THEN igv * -1 ELSE igv END), 2) AS igv,
        ROUND(SUM(CASE WHEN idDocumento = '07' THEN Exonerado * -1 ELSE Exonerado END), 2) AS Exonerado,
        ROUND(SUM(CASE WHEN idDocumento = '07' THEN Total * -1 ELSE Total END), 2) AS Total,
        ROUND(SUM(CASE WHEN idDocumento = '07' THEN BaseimpDol * -1 ELSE BaseimpDol END), 2) AS BaseimpDol,
        ROUND(SUM(CASE WHEN idDocumento = '07' THEN igvDol * -1 ELSE igvDol END), 2) AS igvDol,
        ROUND(SUM(CASE WHEN idDocumento = '07' THEN ExoneradoDol * -1 ELSE ExoneradoDol END), 2) AS ExoneradoDol,
        ROUND(SUM(CASE WHEN idDocumento = '07' THEN TotalDol * -1 ELSE TotalDol END), 2) AS TotalDol
    FROM (
        SELECT
            d.idSerie,
            d.idDocumento,
            d.idDocVentas,
            o.GlsDocumento,
            d.FecEmision,
            d.idPerCliente,
            d.GlsCliente,
            dt.idproducto,
            dt.glsproducto,
            d.RUCCliente,
            CASE 
                WHEN @varTipo = '1' THEN
                    CASE 
                        WHEN d.idmoneda = 'PEN' THEN 
                            CASE WHEN Afecto = '1' THEN dt.TotalVVNeto ELSE 0 END
                        ELSE 
                            CASE WHEN Afecto = '1' THEN dt.TotalVVNeto * ISNULL(T.TcVenta, ISNULL(Tc.TipoCambio, d.TipoCambio)) ELSE 0 END
                    END
                ELSE 
                    CASE 
                        WHEN d.idmoneda = 'PEN' THEN 
                            CASE WHEN Afecto = '1' THEN dt.TotalVVNeto ELSE 0 END
                        ELSE 0
                    END
            END AS Baseimp,
            CASE 
                WHEN @varTipo = '1' THEN
                    CASE 
                        WHEN d.idmoneda = 'PEN' THEN dt.TotalIGVNeto
                        ELSE dt.TotalIGVNeto * ISNULL(T.TcVenta, ISNULL(Tc.TipoCambio, d.TipoCambio))
                    END
                ELSE 
                    CASE 
                        WHEN d.idmoneda = 'PEN' THEN dt.TotalIGVNeto
                        ELSE 0
                    END
            END AS igv,
            CASE 
                WHEN @varTipo = '1' THEN
                    CASE 
                        WHEN d.idmoneda = 'PEN' THEN 
                            CASE WHEN Afecto = '0' THEN dt.TotalVVNeto ELSE 0 END
                        ELSE 
                            CASE WHEN Afecto = '0' THEN dt.TotalVVNeto * ISNULL(T.TcVenta, ISNULL(Tc.TipoCambio, d.TipoCambio)) ELSE 0 END
                    END
                ELSE 
                    CASE 
                        WHEN d.idmoneda = 'PEN' THEN 
                            CASE WHEN Afecto = '0' THEN dt.TotalVVNeto ELSE 0 END
                        ELSE 0
                    END
            END AS Exonerado,
            CASE 
                WHEN @varTipo = '1' THEN
                    CASE 
                        WHEN d.idmoneda = 'PEN' THEN dt.TotalPVNeto
                        ELSE dt.TotalPVNeto * ISNULL(T.TcVenta, ISNULL(Tc.TipoCambio, d.TipoCambio))
                    END
                ELSE 
                    CASE 
                        WHEN d.idmoneda = 'PEN' THEN dt.TotalPVNeto
                        ELSE 0
                    END
            END AS Total,
            CASE 
                WHEN @varTipo = '1' THEN
                    CASE 
                        WHEN d.idmoneda = 'USD' THEN 
                            CASE WHEN Afecto = '1' THEN dt.TotalVVNeto ELSE 0 END
                        ELSE 
                            CASE WHEN Afecto = '1' THEN dt.TotalVVNeto / ISNULL(T.TcVenta, ISNULL(Tc.TipoCambio, d.TipoCambio)) ELSE 0 END
                    END
                ELSE 
                    CASE 
                        WHEN d.idmoneda = 'USD' THEN 
                            CASE WHEN Afecto = '1' THEN dt.TotalVVNeto ELSE 0 END
                        ELSE 0
                    END
            END AS BaseimpDol,
            CASE 
                WHEN @varTipo = '1' THEN
                    CASE 
                        WHEN d.idmoneda = 'USD' THEN dt.TotalIGVNeto
                        ELSE dt.TotalIGVNeto / ISNULL(T.TcVenta, ISNULL(Tc.TipoCambio, d.TipoCambio))
                    END
                ELSE 
                    CASE 
                        WHEN d.idmoneda = 'USD' THEN dt.TotalIGVNeto
                        ELSE 0
                    END
            END AS igvDol,
            CASE 
                WHEN @varTipo = '1' THEN
                    CASE 
                        WHEN d.idmoneda = 'USD' THEN 
                            CASE WHEN Afecto = '0' THEN dt.TotalVVNeto ELSE 0 END
                        ELSE 
                            CASE WHEN Afecto = '0' THEN dt.TotalVVNeto / ISNULL(T.TcVenta, ISNULL(Tc.TipoCambio, d.TipoCambio)) ELSE 0 END
                    END
                ELSE 
                    CASE 
                        WHEN d.idmoneda = 'USD' THEN 
                            CASE WHEN Afecto = '0' THEN dt.TotalVVNeto ELSE 0 END
                        ELSE 0
                    END
            END AS ExoneradoDol,
            CASE 
                WHEN @varTipo = '1' THEN
                    CASE 
                        WHEN d.idmoneda = 'USD' THEN dt.TotalPVNeto
                        ELSE dt.TotalPVNeto / ISNULL(T.TcVenta, ISNULL(Tc.TipoCambio, d.TipoCambio))
                    END
                ELSE 
                    CASE 
                        WHEN d.idmoneda = 'USD' THEN dt.TotalPVNeto
                        ELSE 0
                    END
            END AS TotalDol
        FROM docventas d
        INNER JOIN Documentos o ON D.idDocumento = o.idDocumento
        LEFT JOIN tiposdecambio t ON CAST(d.FecEmision AS DATE) = CAST(t.fecha AS DATE)
        LEFT JOIN (
            SELECT 
                X.TcVenta AS TipoCambio,
                R.IdEmpresa,
                R.IdSucursal,
                R.TipoDocOrigen,
                R.SerieDocOrigen,
                R.NumDocOrigen
            FROM DocVentas Dt
            INNER JOIN DocReferencia R ON Dt.IdEmpresa = R.IdEmpresa 
                AND Dt.IdSucursal = R.IdSucursal 
                AND Dt.IdDocumento = R.TipoDocReferencia 
                AND Dt.IdSerie = R.SerieDocReferencia 
                AND Dt.IdDocVentas = R.NumDocReferencia
            INNER JOIN TiposDeCambio X ON Dt.FecEmision = X.Fecha
            WHERE R.IdEmpresa = @varEmpresa 
                AND R.TipoDocOrigen = '07'
            GROUP BY X.TcVenta,
				R.IdEmpresa,
                R.IdSucursal,
                R.NumDocOrigen,
                R.SerieDocOrigen,
                R.TipoDocOrigen

        ) Tc ON D.IdEmpresa = Tc.Idempresa 
            AND D.IdSucursal = Tc.IdSucursal 
            AND D.IdDocumento = Tc.TipoDocOrigen 
            AND D.IdSerie = Tc.SerieDocOrigen 
            AND D.IdDocVentas = Tc.NumDocOrigen
        INNER JOIN docventasdet dt 
			ON d.idempresa = dt.idempresa 
            AND d.IdSucursal = dt.IdSucursal 
            AND d.iddocumento = dt.iddocumento 
            AND d.idserie = dt.idserie 
            AND d.iddocventas = dt.iddocventas
        WHERE d.idEmpresa = @varEmpresa
            AND d.idDocumento IN ('01','03','07','08','12','90','89','56')
            AND MONTH(d.FecEmision) BETWEEN @varMesDesde AND @varMesHasta
            AND YEAR(d.FecEmision) = @varAno
            AND d.estDocVentas <> 'ANU'
            AND (d.idpercliente LIKE @varCliente OR @varCliente = '%%')
    ) AS VENTAS
    GROUP BY 
        idPerCliente,
        --FecEmision,
        GlsCliente,
        --idDocumento,
        --idSerie,
        --idDocVentas,
        --idproducto,
        --glsproducto,
        RUCCliente;
END;
