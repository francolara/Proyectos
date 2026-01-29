
-- Crear el procedimiento almacenado en SQL Server
CREATE OR ALTER PROCEDURE spu_ListaRegVentas
(
    @varEmpresa CHAR(2),
    @varSucursal CHAR(8),
    @varTipoDoc CHAR(2),
    @varSerie CHAR(3),
    @varMoneda CHAR(3),
    @varFechaIni VARCHAR(10),
    @varFechaFin VARCHAR(10),
    @varOficial VARCHAR(250),
    @VarIdArea VARCHAR(8)
)
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @varGlsTipoDoc VARCHAR(180);
    DECLARE @varGlsSucursal VARCHAR(180);
    DECLARE @varGlsSimboloMoneda VARCHAR(180);
    DECLARE @varGlsMoneda VARCHAR(180);
    DECLARE @varGlsEmpresa VARCHAR(180);
    DECLARE @varGlsRuc VARCHAR(180);

    SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;

    IF @varOficial <> '1' 
    BEGIN
        SET @varOficial = '%%';
    END

    IF @varEmpresa <> '' 
    BEGIN
        SELECT @varGlsEmpresa = GlsEmpresa FROM empresas WHERE idempresa = @varEmpresa;
        SELECT @varGlsRuc = ruc FROM empresas WHERE idempresa = @varEmpresa;
    END

    SELECT @varGlsSimboloMoneda = simbolo FROM Monedas WHERE idMoneda = @varMoneda;
    SELECT @varGlsMoneda = GlsMoneda FROM Monedas WHERE idMoneda = @varMoneda;

    IF @varSucursal <> '' 
    BEGIN
        SELECT @varGlsSucursal = GlsPersona FROM Personas WHERE idPersona = @varSucursal;
    END
    ELSE 
    BEGIN
        SET @varGlsSucursal = 'TODAS LAS SUCURSALES';
    END

    IF @varTipoDoc <> '' 
    BEGIN
        SELECT @varGlsTipoDoc = GlsDocumento FROM documentos WHERE idDocumento = @varTipoDoc;
    END
    ELSE 
    BEGIN
        SET @varGlsTipoDoc = 'TODAS LOS DOCUMENTOS';
    END

    -- Convertir fechas
    SET @varFechaIni = CONVERT(VARCHAR, CAST(@varFechaIni AS DATE), 103);
    SET @varFechaFin = CONVERT(VARCHAR, CAST(@varFechaFin AS DATE), 103);

    -- Selección final
    SELECT 
        @varGlsSucursal AS GlsSucursal, 
        @varGlsEmpresa AS GlsEmpresa, 
        @varGlsRuc AS GlsRuc,
        @varGlsTipoDoc AS GlsTipoDoc,
        CONVERT(VARCHAR,@varFechaIni,103) AS FechaINI,
        CONVERT(VARCHAR,@varFechaFin,103) AS FECHAFIN,
        (@varSucursal + idDocumento+idSerie + idDocVentas) AS item,
        FecEmision,
        idDocumento, 
        AbreDocumento, 
        GlsDocumento,
        idSerie, 
        idDocVentas, 
        idmoneda,
        idPerCliente,
        CASE estDocVentas WHEN 'ANU' THEN 'ANULADO' ELSE GlsCliente END AS GlsCliente,
        CASE estDocVentas WHEN 'ANU' THEN '' ELSE RUCCliente END AS RUCCliente,
        ROUND((CASE WHEN idDocumento = '07' OR (IdDocumento = '25' AND IndAtribucionNC = 1) THEN TotalBaseImponible * -1 ELSE TotalBaseImponible END) * CASE WHEN TotalIVAP <> 0 THEN 0 ELSE 1 END, 2) AS baseImponible,
        ROUND((CASE WHEN idDocumento = '07' OR (IdDocumento = '25' AND IndAtribucionNC = 1) THEN TotalBaseImponible * -1 ELSE TotalBaseImponible END) * CASE WHEN TotalIVAP <> 0 THEN 1 ELSE 0 END, 2) AS baseImponibleIVAP,
        ROUND(CASE WHEN idDocumento = '07' OR (IdDocumento = '25' AND IndAtribucionNC = 1) THEN TotalDsctoVV * -1 ELSE TotalDsctoVV END, 2) AS dscto,
        ROUND(CASE WHEN idDocumento = '07' OR (IdDocumento = '25' AND IndAtribucionNC = 1) THEN TotalExonerado * -1 ELSE TotalExonerado END, 2) AS exonerado,
        ROUND(CASE WHEN idDocumento = '07' OR (IdDocumento = '25' AND IndAtribucionNC = 1) THEN TotalIGVVenta * -1 ELSE TotalIGVVenta END, 2) AS TotalIGVVenta,
        ROUND(CASE WHEN idDocumento = '07' OR (IdDocumento = '25' AND IndAtribucionNC = 1) THEN TotalIVAP * -1 ELSE TotalIVAP END, 2) AS TotalIVAP,
        ROUND(CASE WHEN idDocumento = '07' OR (IdDocumento = '25' AND IndAtribucionNC = 1) THEN TotalPrecioVenta * -1 ELSE TotalPrecioVenta END, 2) AS TotalPrecioVenta,
        FORMAT(CASE WHEN idDocumento = '07' OR (IdDocumento = '25' AND IndAtribucionNC = 1) THEN TotalBaseImponible * -1 ELSE TotalBaseImponible END, 'N2') AS strbaseImponible,
        FORMAT(TotalDsctoVV, 'N2') AS strdscto,
        FORMAT(TotalExonerado, 'N2') AS strexonerado,
        FORMAT(CASE WHEN idDocumento = '07' OR (IdDocumento = '25' AND IndAtribucionNC = 1) THEN TotalIGVVenta * -1 ELSE TotalIGVVenta END, 'N2') AS strTotalIGVVenta,
        FORMAT(CASE WHEN idDocumento = '07' OR (IdDocumento = '25' AND IndAtribucionNC = 1) THEN TotalPrecioVenta * -1 ELSE TotalPrecioVenta END, 'N2') AS strTotalPrecioVenta,
        TipoCambio,
        idSucursalDoc, 
        GlsSucursalDoc, 
        GlsMoneda AS Moneda,
        CASE idMoneda WHEN 'USD' THEN CASE WHEN idDocumento = '07' OR (IdDocumento = '25' AND IndAtribucionNC = 1) THEN TotalPrecioVentaOri * -1 ELSE TotalPrecioVentaOri END ELSE 0 END AS TotalDolares,
        CASE idMoneda WHEN 'USD' THEN 'US$' ELSE 'S/.' END AS Moneda,
        tipoDocReferencia, 
        serieDocReferencia, 
        numDocReferencia,
        DescUnidad,
        GlsUPCliente,
        ObsRegVentas,
        GlsFormasPago, 
        idCentroCosto
    FROM
    (
			SELECT d.idSucursal AS idSucursalDoc, p.GlsPersona AS GlsSucursalDoc,
			d.idMoneda,m.simbolo,m.GlsMoneda,
			d.idSerie,d.idDocumento,
			AbreDocumento,o.GlsDocumento,d.idDocVentas,
			CONVERT(VARCHAR,d.FecEmision,103) AS FecEmision,
			idPerCliente,GlsCliente,RUCCliente,
			case d.idDocumento when '07' then isnull(tc.TipoCambio,d.Tipocambio) else t.tcVenta end as TipoCambio,
			d.estDocVentas,tc.tipoDocReferencia,tc.serieDocReferencia, tc.numDocReferencia,d.ObsRegVentas,d.GlsFormasPago,
			
			CASE WHEN D.IndTransGratuita = '1' OR D.IndTransGratuitaMP = '1' THEN 0 ELSE 1 END *
			IIF(d.estDocVentas = 'ANU',0,
			CASE @varMoneda WHEN 'PEN' THEN IIF(d.idMoneda = 'PEN', d.TotalValorVenta,d.TotalValorVenta * iIf(d.iddocumento <> '07', t.tcVenta, isnull(tc.TipoCambio,d.TipoCambio)))
			WHEN 'USD' THEN IIF(d.idMoneda = 'USD', d.TotalValorVenta,d.TotalValorVenta / iIf(d.iddocumento <> '07', t.tcVenta, isnull(tc.TipoCambio,d.TipoCambio))) END) AS TotalValorVenta,

			CASE WHEN D.IndTransGratuita = '1' OR D.IndTransGratuitaMP = '1' THEN 0 ELSE 1 END *
			IIF(d.estDocVentas = 'ANU',0,
			CASE @varMoneda WHEN 'PEN' THEN IIF(d.idMoneda = 'PEN', d.TotalIGVVenta,d.TotalIGVVenta * iIf(d.iddocumento <> '07', t.tcVenta, isnull(tc.TipoCambio,d.TipoCambio)))
			WHEN 'USD' THEN IIF(d.idMoneda = 'USD', d.TotalIGVVenta,d.TotalIGVVenta / iIf(d.iddocumento <> '07', t.tcVenta, isnull(tc.TipoCambio,d.TipoCambio))) END) AS TotalIGVVenta,

			CASE WHEN D.IndTransGratuita = '1' OR D.IndTransGratuitaMP = '1' THEN 0 ELSE 1 END *
			IIF(d.estDocVentas = 'ANU',0,
			CASE @varMoneda WHEN 'PEN' THEN IIF(d.idMoneda = 'PEN', d.TotalIVAP,d.TotalIVAP * IIF(d.iddocumento <> '07', t.tcVenta, isnull(tc.TipoCambio,d.TipoCambio)))
			WHEN 'USD' THEN IIF(d.idMoneda = 'USD', d.TotalIVAP,d.TotalIVAP / IIF(d.iddocumento <> '07', t.tcVenta, isnull(tc.TipoCambio,d.TipoCambio))) END) AS TotalIVAP,

			CASE WHEN D.IndTransGratuita = '1' OR D.IndTransGratuitaMP = '1' THEN 0 ELSE 1 END *
			IIF(d.estDocVentas = 'ANU',0,
			CASE @varMoneda WHEN 'PEN' THEN IIF(d.idMoneda = 'PEN', d.TotalPrecioVenta,d.TotalPrecioVenta * iIf(d.iddocumento <> '07', t.tcVenta, isnull(tc.TipoCambio,d.TipoCambio)))
			WHEN 'USD' THEN IIF(d.idMoneda = 'USD', d.TotalPrecioVenta,d.TotalPrecioVenta / iIf(d.iddocumento <> '07', t.tcVenta, isnull(tc.TipoCambio,d.TipoCambio))) END) As TotalPrecioVenta,

			CASE WHEN D.IndTransGratuita = '1' OR D.IndTransGratuitaMP = '1' THEN 0 ELSE 1 END *
			IIF(d.estDocVentas = 'ANU',0,
			CASE @varMoneda WHEN 'PEN' THEN IIF(d.idMoneda = 'PEN', d.TotalBaseImponible,d.TotalBaseImponible * iIf(d.iddocumento <> '07', t.tcVenta, isnull(tc.TipoCambio,d.TipoCambio)))
			WHEN 'USD' THEN IIF(d.idMoneda = 'USD', d.TotalBaseImponible,d.TotalBaseImponible / iIf(d.iddocumento <> '07', t.tcVenta, isnull(tc.TipoCambio,d.TipoCambio))) END) AS TotalBaseImponible,

			CASE WHEN D.IndTransGratuita = '1' OR D.IndTransGratuitaMP = '1' THEN 0 ELSE 1 END *
			IIF(d.estDocVentas = 'ANU',0,
			CASE @varMoneda WHEN 'PEN' THEN IIF(d.idMoneda = 'PEN', d.TotalDsctoVV,d.TotalDsctoVV * iIf(d.iddocumento <> '07', t.tcVenta, isnull(tc.TipoCambio,d.TipoCambio)))
			WHEN 'USD' THEN IIF(d.idMoneda = 'USD', d.TotalDsctoVV,d.TotalDsctoVV / IIF(d.iddocumento <> '07', t.tcVenta, isnull(tc.TipoCambio,d.TipoCambio))) END) AS TotalDsctoVV,

			CASE WHEN D.IndTransGratuita = '1' OR D.IndTransGratuitaMP = '1' THEN 0 ELSE 1 END *
			IIF(d.estDocVentas = 'ANU',0,
			CASE @varMoneda WHEN 'PEN' THEN IIF(d.idMoneda = 'PEN', d.TotalExonerado,d.TotalExonerado * iIf(d.iddocumento <> '07', t.tcVenta, isnull(tc.TipoCambio,d.TipoCambio)))
			WHEN 'USD' THEN IIF(d.idMoneda = 'USD', d.TotalExonerado,d.TotalExonerado / IIF(d.iddocumento <> '07', t.tcVenta, isnull(tc.TipoCambio,d.TipoCambio))) END) As TotalExonerado,
			IIF(d.estDocVentas = 'ANU',0,d.TotalPrecioVenta) As TotalPrecioVentaOri,isnull(z.DescUnidad,'') DescUnidad,isnull(xx.GlsUPCliente,'') GlsUPCliente,
			dd.idCentroCosto,d.IndAtribucionNC

			FROM docventas d
			inner join Documentos o
			  on D.idDocumento = o.idDocumento

			left join unidadproduccion z
			  on d.idempresa = z.idempresa
			  and d.idupp = z.CodUnidProd

			left join
			(
			  Select x.idempresa,x.idsucursal,x.iddocumento,x.idserie,x.iddocventas,z.GlsUPCliente
			  from Docventas a
			  inner join docventasdet x
				on a.idempresa = x.idempresa
				and a.idsucursal = x.idsucursal
				and a.iddocumento = x.iddocumento
				and a.idserie = x.idserie
				and a.iddocventas = x.iddocventas
			  inner join unidadproduccioncliente z
				on x.idempresa = z.idempresa
				and x.IdUPCliente = z.IdUPCliente
			  WHERE a.idEmpresa = @varEmpresa
			  AND (a.idSucursal = @varSucursal OR @varSucursal = '')
			  AND (a.idDocumento = @varTipoDoc OR @varTipoDoc = '')
			  AND (a.idSerie = @varSerie OR @varSerie = '')
			  AND (a.IdUPP = @VarIdArea Or @VarIdArea = '')
			  AND a.idDocumento IN ('01','03','07','08','12','90','89','25')
			  AND a.FecEmision between @varFechaIni AND @varFechaFin
			  Group by x.idempresa,x.idsucursal,x.iddocumento,x.idserie,x.iddocventas,z.GlsUPCliente
			) xx
			  on d.idempresa = xx.idempresa
			  and d.idsucursal = xx.idsucursal
			  and d.iddocumento = xx.iddocumento
			  and d.idserie = xx.idserie
			  and d.iddocventas = xx.iddocventas

			inner join Monedas m
			  on m.idMoneda = @varMoneda

			inner join Personas p
			  on d.idSucursal = p.idPersona

			left join tiposdecambio t
			  on CAST(d.FecEmision AS DATE) = CAST(t.Fecha AS DATE)
			  /*(Day(d.FecEmision) = Day(t.fecha)
			  AND Year(d.FecEmision) = Year(t.fecha)
			  AND Month(d.FecEmision) = Month(t.fecha))*/

			left join (select x.tcVenta as tipoCambio,r.idempresa,r.idsucursal,r.tipoDocOrigen,
			r.serieDocOrigen, r.numDocOrigen, r.tipoDocReferencia,r.serieDocReferencia, r.numDocReferencia
			from docventas dt
			inner join docreferencia r
			  on dt.idEmpresa = r.idEmpresa
			  and dt.idsucursal = r.idsucursal
			  and dt.iddocumento = r.tipoDocReferencia
			  and dt.idSerie = r.serieDocReferencia
			  and dt.idDocVentas = r.numDocReferencia
			left join tiposdecambio x
			  on CAST(dt.FecEmision AS DATE) = CAST(x.Fecha AS DATE)
			where r.tipoDocOrigen In('07','08')
			Group By x.tcVenta,r.idempresa,r.idsucursal,r.numDocOrigen,r.serieDocOrigen ,r.tipoDocOrigen,r.tipoDocReferencia,r.serieDocReferencia, r.numDocReferencia
			) tc
			  on d.idempresa = tc.idempresa
			  and d.idsucursal = tc.idsucursal
			  and d.idDocumento = tc.tipoDocOrigen
			  and d.idSerie = tc.serieDocOrigen
			  and d.idDocVentas = tc.numDocOrigen

			Inner Join (
			  Select idEmpresa, idSucursal, idDocumento, idSerie, idDocVentas, Replace(STRING_AGG(idCentroCosto,','),',',' - ') As idCentroCosto
			  FROM Docventasdet a
			  WHERE a.idEmpresa = @varEmpresa
			  AND (a.idSucursal = @varSucursal OR @varSucursal = '')
			  AND (a.idDocumento = @varTipoDoc OR @varTipoDoc = '')
			  AND (a.idSerie = @varSerie OR @varSerie = '')
			  AND a.idDocumento IN ('01','03','07','08','12','90','89','25','56')
			  Group By idDocumento, idSerie, idDocventas, idEmpresa, idSucursal
			) dd
			  on d.idempresa = dd.idempresa
			  And d.idsucursal = dd.idsucursal
			  And d.idDocumento = dd.idDocumento
			  And d.idSerie = dd.idSerie
			  And d.idDocVentas = dd.idDocVentas

			WHERE d.idEmpresa = @varEmpresa
			AND (d.idSucursal = @varSucursal OR @varSucursal = '')
			AND (d.idDocumento = @varTipoDoc OR @varTipoDoc = '')
			AND (d.idSerie = @varSerie OR @varSerie = '')
			AND (d.IdUPP = @VarIdArea Or @VarIdArea = '')
			AND d.idDocumento IN ('01','03','07','08','12','90','89','25','56')
			-- AND (d.idPerCliente like varCliente)
			-- AND (o.IndOficial LIKE @varOficial)
			AND CAST(d.FecEmision AS DATE) between CAST(@varFechaIni AS DATE) AND CAST(@varFechaFin AS DATE)

	) VENTAS
	ORDER BY idSucursalDoc, idDocumento, idSerie, idDocVentas




END
GO
