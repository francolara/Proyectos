DELIMITER $$

-- Descripcion de version 1  : Reporte de Lista de Ventas por Cliente
-- Fecha de Version 1		     : 17/09/2012
-- Creador de version 1		   : SAV Peru
-- Fecha de Pase Version 1	 : 17/09/2012
-- Autorizador Version 1		 : SAV
-- Descripción               : Reporte de Lista de Ventas por Cliente
-- Llamado                   : Sistema de Ventas

DROP PROCEDURE IF EXISTS  `spu_ListaVentasPorCliente` $$
CREATE  PROCEDURE `spu_ListaVentasPorCliente`(
varEmpresa     CHAR(2),
varSucursal    CHAR(8),
varMoneda      CHAR(3),
varFechaIni    VARCHAR(10),
varFechaFin    VARCHAR(10),
varCliente     CHAR(8),
varOficial     VARCHAR(250),
varOrden       varchar(200))
BEGIN

  DECLARE varGlsEmpresa    VARCHAR(200);
  DECLARE varGlsRuc        VARCHAR(180);
  DECLARE varGlsSistema    VARCHAR(180);


  DECLARE varGlsSucursal VARCHAR(180);
  DECLARE strSQL         VARCHAR(250);

  IF varEmpresa <> '' THEN
    SET varGlsEmpresa = (SELECT glsEmpresa FROM Empresas WHERE idEmpresa = varEmpresa);
  END IF;

  IF varEmpresa <> '' THEN
      SET varGlsRuc = (SELECT ruc FROM Empresas where idEmpresa = varEmpresa);
  END IF;

  SET varGlsSistema = 'Sistema de Ventas';

  IF varOficial <> '1' THEN
    SET varOficial = '%%';
  END IF;

  IF varSucursal <> '' THEN
    SET varGlsSucursal = (SELECT GlsPersona FROM Personas WHERE idPersona = varSucursal);
  ELSE
    SET varGlsSucursal = 'TODAS LAS SUCURSALES';
  END IF;

  IF varCliente <> '' THEN
    SET varCliente = varCliente;
  ELSE
    SET varCliente = '%%';
  END IF;



  Set @VarSQl = ConCat('Select ''',varGlsSucursal,''' GlsSucursal,Date_Format(''',varFechaIni,''',''%d/%m/%Y'') FechaINI,',
  'Date_Format(''',varFechaFin,''',''%d/%m/%Y'') FECHAFIN,simbolo,GlsMoneda,idPerCliente,RUCCliente,GlsCliente,AbreDocumento,idSerie,',
  'idDocVentas,FecEmision,GlsFecVectos,GlsZona,idZona, ',
  '''',varGlsRuc,''' as Ruc, ''',varGlsSistema,''' as GlsSistema,''',varGlsEmpresa,''' as Glsempresa, ',
  'IF(idDocumento In(''07'',''89''),TotalValorVenta * - 1,TotalValorVenta) AS TotalValorVenta,',
  'IF(idDocumento In(''07'',''89''),TotalIGVVenta * - 1,TotalIGVVenta) AS TotalIGVVenta,',
  'IF(idDocumento In(''07'',''89''),TotalPrecioVenta * - 1,TotalPrecioVenta) AS TotalPrecioVenta,',

  'IF(idDocumento In(''07'',''89''),TotalValorVentaSoles * - 1,TotalValorVentaSoles) AS TotalValorVentaSoles,',
  'IF(idDocumento In(''07'',''89''),TotalValorVentaDolar * - 1,TotalValorVentaDolar) AS TotalValorVentaDolar,',

  'IF(idDocumento In(''07'',''89''),TotalIGVVentaSoles * - 1,TotalIGVVentaSoles) AS TotalIGVVentaSoles,',
  'IF(idDocumento In(''07'',''89''),TotalIGVVentaDolar * - 1,TotalIGVVentaDolar) AS TotalIGVVentaDolar,',

  'IF(idDocumento In(''07'',''89''),TotalPrecioVentaSoles * - 1,TotalPrecioVentaSoles) AS TotalPrecioVentaSoles,',
  'IF(idDocumento In(''07'',''89''),TotalPrecioVentaDolar * - 1,TotalPrecioVentaDolar) AS TotalPrecioVentaDolar ',

  'FROM(',
    'SELECT m.simbolo,m.GlsMoneda,d.idSerie,d.idDocumento,d.idPerCliente,pr.Ruc RUCCliente,pr.GlsPersona GlsCliente,AbreDocumento,',
    'd.idDocVentas,DATE_FORMAT(d.FecEmision,''%d/%m/%Y'') AS FecEmision,d.GlsFecVectos,z.GlsZona,z.idZona, ',

    'CASE ''',varMoneda,''' ',
    'WHEN ''PEN'' THEN  IF(d.idMoneda = ''PEN'', d.TotalValorVenta,d.TotalValorVenta * if(d.iddocumento <> ''07'', t.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio))) ',
    'WHEN ''USD'' THEN  IF(d.idMoneda = ''USD'', d.TotalValorVenta,d.TotalValorVenta / if(d.iddocumento <> ''07'', t.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio))) END AS TotalValorVenta,',

    'CASE ''',varMoneda,''' ',
    'WHEN ''PEN'' THEN  IF(d.idMoneda = ''PEN'', d.TotalIGVVenta,d.TotalIGVVenta * if(d.iddocumento <> ''07'', t.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio))) ',
    'WHEN ''USD'' THEN  IF(d.idMoneda = ''USD'', d.TotalIGVVenta,d.TotalIGVVenta / if(d.iddocumento <> ''07'', t.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio))) END AS TotalIGVVenta,',

    'CASE ''',varMoneda,''' ',
    'WHEN ''PEN'' THEN  IF(d.idMoneda = ''PEN'', d.TotalPrecioVenta,d.TotalPrecioVenta * if(d.iddocumento <> ''07'', t.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio))) ',
    'WHEN ''USD'' THEN  IF(d.idMoneda = ''USD'', d.TotalPrecioVenta,d.TotalPrecioVenta / if(d.iddocumento <> ''07'', t.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio))) END As TotalPrecioVenta,',

    'IF(d.idMoneda = ''PEN'', d.TotalValorVenta,0) AS TotalValorVentaSoles,',
    'IF(d.idMoneda = ''USD'', d.TotalValorVenta,0) AS TotalValorVentaDolar,',

    'IF(d.idMoneda = ''PEN'', d.TotalIGVVenta,0) AS TotalIGVVentaSoles,',
    'IF(d.idMoneda = ''USD'', d.TotalIGVVenta,0) AS TotalIGVVentaDolar,',

    'IF(d.idMoneda = ''PEN'', d.TotalPrecioVenta,0) As TotalPrecioVentaSoles,',
    'IF(d.idMoneda = ''USD'', d.TotalPrecioVenta,0) As TotalPrecioVentaDolar ',

    'FROM docventas d ',
    'Inner Join Documentos o ',
      'On D.idDocumento = o.idDocumento ',
    'Inner Join Monedas m ',
      -- 'On d.idMoneda=m.idMoneda ',
    'Inner Join Personas pr ',
      'On pr.IdPersona = d.IdPerCliente ',

    'Left Join Ubigeo ub ',
      'On pr.IdPais = ub.IdPais And ub.IdDistrito=pr.IdDistrito ',

    'Left Join  Zonas z ',
      'On ub.idZona=z.idZona ',

    'left join tiposdecambio t ',
      'on d.FecEmision = t.fecha ',

    'left join (select x.tcVenta as tipoCambio,r.idempresa,r.idsucursal,r.tipoDocOrigen,',
    'r.serieDocOrigen, r.numDocOrigen ',
    'from docventas dt ',
    'left join docreferencia r ',
    'on dt.idEmpresa = r.idEmpresa ',
    'and dt.idsucursal = r.idsucursal ',
    'and dt.iddocumento = r.tipoDocReferencia ',
    'and  dt.idSerie = r.serieDocReferencia ',
    'and dt.idDocVentas = r.numDocReferencia ',
    'left join tiposdecambio x ',
    'on dt.FecEmision = x.fecha ',
    'where r.tipoDocOrigen = ''07'') tc ',
    'on d.idempresa = tc.idempresa ',
    'and d.idsucursal = tc.idsucursal ',
    'and d.idDocumento = tc.tipoDocOrigen ',
    'and d.idSerie = tc.serieDocOrigen ',
    'and d.idDocVentas = tc.numDocOrigen ',

    'WHERE d.estDocVentas <> ''ANU'' ',
    'AND ''',varMoneda,''' = m.idMoneda ',
    'AND d.idEmpresa = ''',varEmpresa,''' ',
    'AND (d.idSucursal = ''',varSucursal,''' OR ''',varSucursal,''' = '''') ',
    'AND d.idDocumento IN (''01'',''03'',''07'',''08'',''12'',''90'',''89'',''56'') ',
    'AND o.IndOficial LIKE ''',varOficial,''' ',
    'AND d.FecEmision between CAST(''',varFechaIni,''' AS DATE) AND CAST(''',varFechaFin,''' AS DATE) ',
    'AND d.idPerCliente like ''',varCliente,''' )',

  'VENTAS ',
  'Order By ',varOrden,'');

   -- select @VarSQl;
  PREPARE strSQL FROM @VarSQl;
  EXECUTE strSQL;
  DEALLOCATE PREPARE strSQL;

END $$

DELIMITER ;