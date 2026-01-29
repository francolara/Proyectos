CREATE OR ALTER PROCEDURE dbo.spu_ListaVentasPorCliente_Resum
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
    DECLARE @varGlsEmpresa    VARCHAR(200);
    DECLARE @varGlsRuc        VARCHAR(180);
    DECLARE @varGlsSistema    VARCHAR(180) = 'Sistema de Ventas';
    DECLARE @varGlsSucursal   VARCHAR(180);
    DECLARE @strSQL           NVARCHAR(MAX);
    DECLARE @VarSQL           NVARCHAR(MAX);

    IF @varEmpresa <> ''
    BEGIN
        SELECT @varGlsEmpresa = glsEmpresa FROM Empresas WHERE idEmpresa = @varEmpresa;
        SELECT @varGlsRuc = ruc FROM Empresas WHERE idEmpresa = @varEmpresa;
    END

    IF @varOficial <> '1'
        SET @varOficial = '%%';

    IF @varSucursal <> ''
        SELECT @varGlsSucursal = GlsPersona FROM Personas WHERE idPersona = @varSucursal;
    ELSE
        SET @varGlsSucursal = 'TODAS LAS SUCURSALES';

    --IF @varCliente = ''
    --    SET @varCliente = '%%';

				 SELECT ROW_NUMBER() OVER(ORDER BY GlsCliente) Item, 
				  @varGlsSucursal  AS GlsSucursal,
				  CONVERT(VARCHAR,CAST(@varFecDesde AS DATE) ,103) AS FechaINI,
                  CONVERT(VARCHAR,CAST(@varFecHasta AS DATE),103) AS FECHAFIN,
				  --simbolo,GlsMoneda,
				  idPerCliente,RUCCliente,GlsCliente,
				  --AbreDocumento,idSerie,idDocVentas,FecEmision,GlsFecVectos,
				  GlsZona,
				 -- idZona,
                  @varGlsRuc  as Ruc,  
				  @varGlsSistema  as GlsSistema, 
				  @varGlsEmpresa  as Glsempresa,
                  Round(TotalValorVenta,2) as TotalValorVenta,
				  Round(TotalIGVVenta,2) as TotalIGVVenta,
				  Round(TotalPrecioVenta,2) as TotalPrecioVenta,
                  Round(TotalValorVentaSoles,2) as TotalValorVentaSoles,
				  Round(TotalValorVentaDolar,2) TotalValorVentaDolar,
                  Round(TotalIGVVentaSoles,2) as TotalIGVVentaSoles,
				  Round(TotalIGVVentaDolar,2) as TotalIGVVentaDolar,
                  Round(TotalPrecioVentaSoles,2) as TotalPrecioVentaSoles,
				  Round(TotalPrecioVentaDolar,2) as TotalPrecioVentaDolar
                  From (

						  Select @varGlsSucursal  AS GlsSucursal,
						  CONVERT(VARCHAR,CAST(@varFecDesde AS DATE),103) AS FechaINI,
						  CONVERT(VARCHAR,CAST(@varFecHasta AS DATE),103) AS FECHAFIN,
						  --simbolo,GlsMoneda,
						  idPerCliente,RUCCliente,GlsCliente,
						  --AbreDocumento,idSerie,
						  --idDocVentas,FecEmision,GlsFecVectos,
						  GlsZona,
						  --idZona,
						  @varGlsRuc  as Ruc,  
						  @varGlsSistema  as GlsSistema, 
						  @varGlsEmpresa  as Glsempresa,
						  Sum(CASE WHEN idDocumento IN('07','89') THEN TotalValorVenta * -1 ELSE TotalValorVenta END) AS TotalValorVenta,
						  Sum(CASE WHEN idDocumento IN('07','89') THEN TotalIGVVenta * -1 ELSE TotalIGVVenta END) AS TotalIGVVenta,
						  Sum(CASE WHEN idDocumento IN('07','89') THEN TotalPrecioVenta * -1 ELSE TotalPrecioVenta END) AS TotalPrecioVenta,
						  Sum(CASE WHEN idDocumento IN('07','89') THEN TotalValorVentaSoles * -1 ELSE TotalValorVentaSoles END) AS TotalValorVentaSoles,
						  Sum(CASE WHEN idDocumento IN('07','89') THEN TotalValorVentaDolar * -1 ELSE TotalValorVentaDolar END) AS TotalValorVentaDolar,
						  Sum(CASE WHEN idDocumento IN('07','89') THEN TotalIGVVentaSoles * -1 ELSE TotalIGVVentaSoles END) AS TotalIGVVentaSoles,
						  Sum(CASE WHEN idDocumento IN('07','89') THEN TotalIGVVentaDolar * -1 ELSE TotalIGVVentaDolar END) AS TotalIGVVentaDolar,
						  Sum(CASE WHEN idDocumento IN('07','89') THEN TotalPrecioVentaSoles * -1 ELSE TotalPrecioVentaSoles END) AS TotalPrecioVentaSoles,
						  Sum(CASE WHEN idDocumento IN('07','89') THEN TotalPrecioVentaDolar * -1 ELSE TotalPrecioVentaDolar END) AS TotalPrecioVentaDolar
						  From(

								  SELECT m.simbolo,m.GlsMoneda,d.idSerie,d.idDocumento,d.idPerCliente,pr.Ruc RUCCliente,pr.GlsPersona GlsCliente,AbreDocumento,
								  d.idDocVentas,
								  CONVERT(VARCHAR,d.FecEmision,103) AS FecEmision,d.GlsFecVectos,z.GlsZona,z.idZona,

								  CASE  @varMoneda  
								  WHEN 'PEN' THEN  CASE WHEN d.idMoneda = 'PEN' THEN d.TotalValorVenta ELSE d.TotalValorVenta * CASE WHEN d.iddocumento <> '07' THEN t.tcVenta ELSE ISNULL(tc.TipoCambio,d.TipoCambio) END END 
								  WHEN 'USD' THEN  CASE WHEN d.idMoneda = 'USD' THEN d.TotalValorVenta ELSE d.TotalValorVenta / CASE WHEN d.iddocumento <> '07' THEN t.tcVenta ELSE ISNULL(tc.TipoCambio,d.TipoCambio) END END END AS TotalValorVenta,

								  CASE  @varMoneda  
								  WHEN 'PEN' THEN  CASE WHEN d.idMoneda = 'PEN' THEN d.TotalIGVVenta ELSE d.TotalIGVVenta * CASE WHEN d.iddocumento <> '07' THEN t.tcVenta ELSE ISNULL(tc.TipoCambio,d.TipoCambio) END END 
								  WHEN 'USD' THEN  CASE WHEN d.idMoneda = 'USD' THEN d.TotalIGVVenta ELSE d.TotalIGVVenta / CASE WHEN d.iddocumento <> '07' THEN t.tcVenta ELSE ISNULL(tc.TipoCambio,d.TipoCambio) END END END AS TotalIGVVenta,

								  CASE  @varMoneda  
								  WHEN 'PEN' THEN  CASE WHEN d.idMoneda = 'PEN' THEN d.TotalPrecioVenta ELSE d.TotalPrecioVenta * CASE WHEN d.iddocumento <> '07' THEN t.tcVenta ELSE ISNULL(tc.TipoCambio,d.TipoCambio) END END 
								  WHEN 'USD' THEN  CASE WHEN d.idMoneda = 'USD' THEN d.TotalPrecioVenta ELSE d.TotalPrecioVenta / CASE WHEN d.iddocumento <> '07' THEN t.tcVenta ELSE ISNULL(tc.TipoCambio,d.TipoCambio) END END END AS TotalPrecioVenta,

								  CASE WHEN d.idMoneda = 'PEN' THEN d.TotalValorVenta ELSE 0 END AS TotalValorVentaSoles,
								  CASE WHEN d.idMoneda = 'USD' THEN d.TotalValorVenta ELSE 0 END AS TotalValorVentaDolar,

								  CASE WHEN d.idMoneda = 'PEN' THEN d.TotalIGVVenta ELSE 0 END AS TotalIGVVentaSoles,
								  CASE WHEN d.idMoneda = 'USD' THEN d.TotalIGVVenta ELSE 0 END AS TotalIGVVentaDolar,

								  CASE WHEN d.idMoneda = 'PEN' THEN d.TotalPrecioVenta ELSE 0 END AS TotalPrecioVentaSoles,
								  CASE WHEN d.idMoneda = 'USD' THEN d.TotalPrecioVenta ELSE 0 END AS TotalPrecioVentaDolar 

								  FROM docventas d 
								  INNER JOIN Documentos o ON D.idDocumento = o.idDocumento 
								  INNER JOIN Monedas m ON d.idMoneda=m.idMoneda 
								  INNER JOIN Personas pr ON pr.IdPersona=d.IdPerCliente 
								  LEFT JOIN Ubigeo ub ON pr.IdPais = ub.IdPais AND ub.IdDistrito=pr.IdDistrito 
								  LEFT JOIN  Zonas z ON ub.idZona=z.idZona 
								  LEFT JOIN tiposdecambio t ON CAST(d.FecEmision AS DATE)  = CAST(t.fecha  AS DATE) 
								  LEFT JOIN (SELECT x.tcVenta as tipoCambio,r.idempresa,r.idsucursal,r.tipoDocOrigen,r.serieDocOrigen, r.numDocOrigen 
								  FROM docventas dt 
								  LEFT JOIN docreferencia r ON dt.idEmpresa = r.idEmpresa AND dt.idsucursal = r.idsucursal 
								  AND dt.iddocumento = r.tipoDocReferencia AND dt.idSerie = r.serieDocReferencia AND dt.idDocVentas = r.numDocReferencia 
								  LEFT JOIN tiposdecambio x ON dt.FecEmision = x.fecha 
								  WHERE r.tipoDocOrigen = '07') tc 
								  ON d.idempresa = tc.idempresa AND d.idsucursal = tc.idsucursal 
								  AND d.idDocumento = tc.tipoDocOrigen AND d.idSerie = tc.serieDocOrigen 
								  AND d.idDocVentas = tc.numDocOrigen 

								  WHERE d.estDocVentas <> 'ANU' 
								  AND  @varMoneda  = m.idMoneda 
								  AND d.idEmpresa =  @varEmpresa  
								  AND (d.idSucursal =  @varSucursal  OR  @varSucursal  = '') 
								  AND d.idDocumento IN ('01','03','07','08','12','90','89','56') 
								  --AND o.IndOficial LIKE  @varOficial  
								  AND CAST(d.FecEmision AS DATE) BETWEEN CAST(@varFecDesde  AS DATE) AND CAST(@varFecHasta  AS DATE) 
								  AND (d.idPerCliente = @varCliente OR @varCliente = '' )
								  AND (d.indVtaGratuita = '' OR d.indVtaGratuita IS NULL)

						  ) VENTAS 
						  GROUP BY idPerCliente,RUCCliente,GlsCliente,Glszona 
						  --ORDER BY  GlsCliente
                  )VENTASX 
                  ORDER BY  GlsCliente
END
GO
