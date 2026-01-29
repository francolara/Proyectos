GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO

CREATE OR ALTER PROCEDURE spu_ListaVentasPorCliente
    @varEmpresa		CHAR(2),
    @varSucursal	CHAR(8),
    @varMoneda		CHAR(3),
    @varFecDesde	VARCHAR(10),
    @varFecHasta	VARCHAR(10),
    @varCliente		CHAR(8),
    @varOficial		VARCHAR(250),
    @varOrden	    VARCHAR(200)
AS
BEGIN
    DECLARE @varGlsEmpresa	VARCHAR(200);
    DECLARE @varGlsRuc		VARCHAR(180);
    DECLARE @varGlsSistema  VARCHAR(180);
    DECLARE @varGlsSucursal VARCHAR(180);
    DECLARE @strSQL			NVARCHAR(MAX);

    IF @varEmpresa <> ''
    BEGIN
        SELECT @varGlsEmpresa = glsEmpresa FROM Empresas WHERE idEmpresa = @varEmpresa;
    END;

    IF @varEmpresa <> ''
    BEGIN
        SELECT @varGlsRuc = ruc FROM Empresas WHERE idEmpresa = @varEmpresa;
    END;

    SET @varGlsSistema = 'Sistema de Ventas';

    IF @varOficial <> '1'
    BEGIN
        SET @varOficial = '%%';
    END;

    IF @varSucursal <> ''
    BEGIN
        SELECT @varGlsSucursal = GlsPersona FROM Personas WHERE idPersona = @varSucursal;
    END;
    ELSE
    BEGIN
        SET @varGlsSucursal = 'TODAS LAS SUCURSALES';
    END;

    --IF @varCliente = ''
    --BEGIN
    --    SET @varCliente = '%%';
    --END;

    SELECT @varGlsSucursal AS GlsSucursal,
               CONVERT(VARCHAR, CAST(@varFecDesde AS DATE), 103) AS FechaINI,
               CONVERT(VARCHAR, CAST(@varFecHasta AS DATE), 103) AS FECHAFIN,
               m.simbolo,
               m.GlsMoneda,
               d.idPerCliente,
               pr.Ruc AS RUCCliente,
               pr.GlsPersona AS GlsCliente,
               AbreDocumento,
               d.idSerie,
               d.idDocVentas,
               CONVERT(VARCHAR, d.FecEmision, 103) AS FecEmision,
               d.GlsFecVectos,
               z.GlsZona,
               z.idZona,
                @varGlsRuc 		AS Ruc,
                @varGlsSistema  AS GlsSistema,
                @varGlsEmpresa  AS Glsempresa,
               CASE WHEN d.idDocumento IN ('07', '89') THEN d.TotalValorVenta * -1 ELSE d.TotalValorVenta END AS TotalValorVenta,
               CASE WHEN d.idDocumento IN ('07', '89') THEN d.TotalIGVVenta * -1 ELSE d.TotalIGVVenta END AS TotalIGVVenta,
               CASE WHEN d.idDocumento IN ('07', '89') THEN d.TotalPrecioVenta * -1 ELSE d.TotalPrecioVenta END AS TotalPrecioVenta,
               CASE WHEN d.idMoneda = 'PEN' THEN d.TotalValorVenta ELSE 0 END AS TotalValorVentaSoles,
               CASE WHEN d.idMoneda = 'USD' THEN d.TotalValorVenta ELSE 0 END AS TotalValorVentaDolar,
               CASE WHEN d.idMoneda = 'PEN' THEN d.TotalIGVVenta ELSE 0 END AS TotalIGVVentaSoles,
               CASE WHEN d.idMoneda = 'USD' THEN d.TotalIGVVenta ELSE 0 END AS TotalIGVVentaDolar,
               CASE WHEN d.idMoneda = 'PEN' THEN d.TotalPrecioVenta ELSE 0 END AS TotalPrecioVentaSoles,
               CASE WHEN d.idMoneda = 'USD' THEN d.TotalPrecioVenta ELSE 0 END AS TotalPrecioVentaDolar
        FROM docventas d
        INNER JOIN Documentos o ON d.idDocumento = o.idDocumento
        INNER JOIN Monedas m On d.idMoneda=m.idMoneda
        INNER JOIN Personas pr ON pr.IdPersona = d.IdPerCliente
        LEFT JOIN Ubigeo ub ON pr.IdPais = ub.IdPais AND ub.IdDistrito = pr.IdDistrito
        LEFT JOIN Zonas z ON ub.idZona = z.idZona
        LEFT JOIN tiposdecambio t ON CAST(d.FecEmision AS DATE) = CAST(t.fecha AS DATE)
        LEFT JOIN (
            SELECT x.tcVenta AS tipoCambio,
                   r.idempresa,
                   r.idsucursal,
                   r.tipoDocOrigen,
                   r.serieDocOrigen,
                   r.numDocOrigen
            FROM docventas dt
            LEFT JOIN docreferencia r ON dt.idEmpresa = r.idEmpresa
                                       AND dt.idsucursal = r.idsucursal
                                       AND dt.iddocumento = r.tipoDocReferencia
                                       AND dt.idSerie = r.serieDocReferencia
                                       AND dt.idDocVentas = r.numDocReferencia
            LEFT JOIN tiposdecambio x ON CAST(dt.FecEmision AS DATE) = CAST(x.fecha AS DATE)
            WHERE r.tipoDocOrigen = '07'
        ) tc ON d.idempresa = tc.idempresa
               AND d.idsucursal = tc.idsucursal
               AND d.idDocumento = tc.tipoDocOrigen
               AND d.idSerie = tc.serieDocOrigen
               AND d.idDocVentas = tc.numDocOrigen
        WHERE d.estDocVentas <> 'ANU'
          AND m.idMoneda = @varMoneda 
          AND d.idEmpresa = @varEmpresa 
          AND (d.idSucursal = @varSucursal OR @varSucursal = '')
          AND d.idDocumento IN ('01', '03', '07', '08', '12', '90', '89', '56')
         
          AND d.FecEmision BETWEEN CAST(@varFecDesde AS DATE) AND CAST(@varFecHasta AS DATE)
          AND (d.idPerCliente = @varCliente OR @varCliente = '')
        ORDER BY GlsCliente,Cast(FecEmision AS Date),AbreDocumento,idSerie,idDocVentas  ;
	 --AND o.IndOficial LIKE " + @varOficial + "
	-- PRINT @strSQL

END;
GO
