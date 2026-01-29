DELIMITER $$


DROP PROCEDURE IF EXISTS `spu_ListaVentasPorProducto` $$
CREATE  PROCEDURE `spu_ListaVentasPorProducto`(
varEmpresa     CHAR(2),
varSucursal    CHAR(8),
varMoneda      CHAR(3),
varFechaIni    VARCHAR(20),
varFechaFin    VARCHAR(20),
varProducto    CHAR(8),
varOficial     VARCHAR(250),
varNiveles     VarChar(250),
varGrupo       Varchar(250),
varOrden       Varchar(250))
BEGIN
  DECLARE varGlsSucursal     VARCHAR(180);
  DECLARE strSQL             Text;

  DECLARE VarGlsEmpresa      VARCHAR(250);
  DECLARE VarGlsRuc          VARCHAR(250);

  Set VarGlsEmpresa = (Select GlsEmpresa From Empresas Where IdEmpresa = VarEmpresa);
  Set VarGlsRuc = (Select Ruc From Empresas Where IdEmpresa = VarEmpresa);


  IF varOficial <> '1' THEN
    SET varOficial = '%%';
  END IF;

  IF varSucursal <> '' THEN
    SET varGlsSucursal = (SELECT GlsPersona FROM Personas WHERE idPersona = varSucursal);
  ELSE
    SET varGlsSucursal = 'TODAS LAS SUCURSALES';
  END IF;

  Set @VarSQl = ConCat('Select (@i:=@i + 1) Item,''',varGlsSucursal,''' As GlsSucursal,Date_Format(''',varFechaIni,''',''%d/%m/%Y'') AS FechaINI,',
  'Date_Format(''',varFechaFin,''',''%d/%m/%Y'') AS FECHAFIN,simbolo,GlsMoneda,idProducto,GlsProducto,idempresa,',varNiveles,'',
  'abreUM,Documento,GlsCliente,FecEmision,''',VarGlsruc,''' As GlsRuc,''',VarGlsEmpresa,''' As GlsEmpresa, ',
  'round(TotalValorVenta,2) as TotalValorVenta,round(TotalIGVVenta,2) as TotalIGVVenta,round(TotalPrecioVenta,2) as TotalPrecioVenta, ',
  'round(TotalValorVentaSoles) as TotalValorVentaSoles,round(TotalValorVentaDolares,2) as TotalValorVentaDolares, ',
  'round(TotalIGVVentaSoles,2) as TotalIGVVentaSoles,round(TotalIGVVentaDolares) as TotalIGVVentaDolares,round(TotalPrecioVentaSoles,2) as TotalPrecioVentaSoles, round(TotalPrecioVentaDolares,2) as TotalPrecioVentaDolares ,Cantidad,TotalVVUnit,PorcentajeVentas ',
  'From (Select @i:=0) foo,( ',
  'Select (@i:=@i) Item, ''',varGlsSucursal,''' As GlsSucursal,Date_Format(''',varFechaIni,''',''%d/%m/%Y'') AS FechaINI,',
  'Date_Format(''',varFechaFin,''',''%d/%m/%Y'') AS FECHAFIN,simbolo,GlsMoneda,idProducto,GlsProducto,idempresa,',varNiveles,'',
  'abreUM,Documento,GlsCliente,FecEmision,''',VarGlsruc,''' As GlsRuc,''',VarGlsEmpresa,''' As GlsEmpresa, ',
  'Sum(If(IdDocumento In(''07'',''89''),TotalValorVenta * - 1,TotalValorVenta)) AS TotalValorVenta,',
  'Sum(If(IdDocumento In(''07'',''89''),TotalIGVVenta * - 1,TotalIGVVenta)) AS TotalIGVVenta,',
  'Sum(If(IdDocumento In(''07'',''89''),TotalPrecioVenta * - 1,TotalPrecioVenta)) AS TotalPrecioVenta,',
  'Sum(If(IdDocumento In(''07'',''89''),TotalValorVentaSoles * - 1,TotalValorVentaSoles)) AS TotalValorVentaSoles,',
  'Sum(If(IdDocumento In(''07'',''89''),TotalValorVentaDolares * - 1,TotalValorVentaDolares)) AS TotalValorVentaDolares,',
  'Sum(If(IdDocumento In(''07'',''89''),TotalIGVVentaSoles * - 1,TotalIGVVentaSoles)) AS TotalIGVVentaSoles,',
  'Sum(If(IdDocumento In(''07'',''89''),TotalIGVVentaDolares * - 1,TotalIGVVentaDolares)) AS TotalIGVVentaDolares,',
  'Sum(If(IdDocumento In(''07'',''89''),TotalPrecioVentaSoles * - 1,TotalPrecioVentaSoles)) AS TotalPrecioVentaSoles,',
  'Sum(If(IdDocumento In(''07'',''89''),TotalPrecioVentaDolares * - 1,TotalPrecioVentaDolares)) AS TotalPrecioVentaDolares, ',
  'Sum(If(IdDocumento In(''07'',''89''),Cantidad * - 1,Cantidad)) AS Cantidad, ',
  'Sum(If(IdDocumento In(''07'',''89''),TotalVVUnit * - 1,TotalVVUnit)) AS TotalVVUnit, ',
  '((ifNull(TotalValorVenta,0) * 100)  / IfNull( TotalPrecioVenta,0)) As  PorcentajeVentas ',

  'From(',
    'Select m.simbolo,m.GlsMoneda,t.idProducto,vn.idempresa,',varNiveles,'',
    'u.abreUM,d.IdDocumento,C.GlsPersona GlsCliente,CONCAT(t.GlsProducto,'' ( '',t.NumLote,'' ) '') As  GlsProducto , ',
    'ConCat(o.AbreDocumento,d.idSerie,''/'',d.idDocVentas) As Documento,Date_Format(d.FecEmision,''%d/%m/%Y'') AS FecEmision,',

    /*'Case ''',varMoneda,''' ',
    'When ''PEN'' Then If(d.IdMoneda = ''PEN'', t.TotalVVNeto,t.TotalVVNeto * TipoCambio) ',
    'When ''USD'' Then If(d.IdMoneda = ''USD'', t.TotalVVNeto,t.TotalVVNeto / TipoCambio) END AS TotalValorVenta,',

    'Case ''',varMoneda,''' ',
    'When ''PEN'' Then If(d.IdMoneda = ''PEN'', t.TotalIGVNeto,t.TotalIGVNeto * TipoCambio) ',
    'When ''USD'' Then If(d.IdMoneda = ''USD'', t.TotalIGVNeto,t.TotalIGVNeto / TipoCambio) END AS TotalIGVVenta,',

    'Case ''',varMoneda,''' ',
    'When ''PEN'' Then If(d.IdMoneda = ''PEN'', t.TotalPVNeto,t.TotalPVNeto * TipoCambio) ',
    'When ''USD'' Then If(d.IdMoneda = ''USD'', t.TotalPVNeto,t.TotalPVNeto / TipoCambio) END AS TotalPrecioVenta,',

    'Case ''',varMoneda,''' ',
    'When ''PEN'' Then If(d.IdMoneda = ''PEN'', t.VVUnit,t.VVUnit * TipoCambio) ',
    'When ''USD'' Then If(d.IdMoneda = ''USD'', t.VVUnit,t.VVUnit / TipoCambio) END AS TotalVVUnit, ',*/

    -- --------------------------------------------------------------------------------------------------------------

    'CASE ''',varMoneda,''' ',
    'WHEN ''PEN'' THEN  IF(d.idMoneda = ''PEN'', t.TotalVVNeto,t.TotalVVNeto * if(d.iddocumento <> ''07'', x.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio))) ',
    'WHEN ''USD'' THEN  IF(d.idMoneda = ''USD'', t.TotalVVNeto,t.TotalVVNeto / if(d.iddocumento <> ''07'', x.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio))) END AS TotalValorVenta, ',

    'CASE ''',varMoneda,''' ',
    'WHEN ''PEN'' THEN  IF(d.idMoneda = ''PEN'', t.TotalIGVNeto,t.TotalIGVNeto * if(d.iddocumento <> ''07'', x.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio))) ',
    'WHEN ''USD'' THEN  IF(d.idMoneda = ''USD'', t.TotalIGVNeto,t.TotalIGVNeto / if(d.iddocumento <> ''07'', x.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio))) END AS TotalIGVVenta,',

    'CASE ''',varMoneda,''' ',
    'WHEN ''PEN'' THEN  IF(d.idMoneda = ''PEN'', t.TotalPVNeto,t.TotalPVNeto * if(d.iddocumento <> ''07'', x.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio))) ',
    'WHEN ''USD'' THEN  IF(d.idMoneda = ''USD'', t.TotalPVNeto,t.TotalPVNeto / if(d.iddocumento <> ''07'', x.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio))) END As TotalPrecioVenta,',

    'CASE ''',varMoneda,''' ',
    'WHEN ''PEN'' THEN  IF(d.idMoneda = ''PEN'',  t.VVUnit, t.VVUnit * if(d.iddocumento <> ''07'', x.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio))) ',
    'WHEN ''USD'' THEN  IF(d.idMoneda = ''USD'',  t.VVUnit, t.VVUnit / if(d.iddocumento <> ''07'', x.tcVenta, ifnull(tc.TipoCambio,d.TipoCambio))) END As TotalVVUnit,',


    -- --------------------------------------------------------------------------------------------------------------


    'If(d.IdMoneda = ''PEN'', t.TotalVVNeto,0) AS TotalValorVentaSoles,',
    'If(d.IdMoneda = ''USD'', t.TotalVVNeto,0) AS TotalValorVentaDolares,',

    'If(d.IdMoneda = ''PEN'', t.TotalIGVNeto,0) AS TotalIGVVentaSoles,',
    'If(d.IdMoneda = ''USD'', t.TotalIGVNeto,0) AS TotalIGVVentaDolares,',

    'If(d.IdMoneda = ''PEN'', t.TotalPVNeto,0) AS TotalPrecioVentaSoles,',
    'If(d.IdMoneda = ''USD'', t.TotalPVNeto,0) AS TotalPrecioVentaDolares, ',
    'If(t.idTipoProducto =''06002'',0,t.Cantidad) as Cantidad ',

    'From DocVentas d ',
    'Inner Join Personas C ',
      'On d.IdPerCliente = C.IdPersona ',
    'Inner Join Documentos o ',
      'On D.idDocumento = o.idDocumento ',
    'Inner Join Monedas m ',
      'On ''',varMoneda,''' = m.idMoneda ',
    'Inner Join DocVentasDet t ',
      'On d.idDocumento = t.idDocumento AND d.idDocVentas = t.idDocVentas AND d.idSerie = t.idSerie And d.idEmpresa = t.idEmpresa ',
      'And d.idSucursal = t.idSucursal ',
     'Inner Join productos p ',
       'On t.idProducto = p.idProducto And t.idEmpresa = p.idEmpresa ',
     'Inner Join vw_niveles vn ',
       'On p.idNivel = vn.idNivel01 And p.idEmpresa = vn.idEmpresa ',
     'Inner Join unidadmedida u ',
       'On t.idUM = u.idUM ',

     'left join tiposdecambio x ',
      'on d.FecEmision = x.fecha ',

    'left join (select x.tcVenta as tipoCambio,r.idempresa,r.idsucursal,r.tipoDocOrigen,',
    'r.serieDocOrigen, r.numDocOrigen ',
    'from docventas dt ',
    'inner join docreferencia r ',
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

    'Where d.idEmpresa = ''',varEmpresa,''' And (d.idSucursal = ''',varSucursal,''' Or ''',varSucursal,''' = '''') ',
    'And d.idDocumento In (''01'',''03'',''07'',''08'',''12'',''90'',''89'',''56'') And o.IndOficial Like ''',varOficial,''' ',
    'And (t.idProducto = ''',varProducto,''' Or ''',varProducto,''' = '''') And d.estDocVentas <> ''ANU'' ',
    'And d.FecEmision BetWeen CAST(''',varFechaIni,''' As Date) And CAST(''',varFechaFin,''' As Date) And (d.indVtaGratuita = '''' Or d.indVtaGratuita is Null)) ',

  'VENTAS ',
  '',varGrupo,' ',
  'Order By ',varOrden,' ',
  ')VENTAS ',
  'Order By item ');

  -- select @VarSQl;
  PREPARE strSQL FROM @VarSQl;
  EXECUTE strSQL;
  DEALLOCATE PREPARE strSQL;
END $$

DELIMITER ;