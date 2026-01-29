DELIMITER $$

DROP PROCEDURE IF EXISTS `spu_ListaInventarioValorizado_Lotes` $$
CREATE PROCEDURE `spu_ListaInventarioValorizado_Lotes`(
varEmpresa CHAR(2),
varSucursal CHAR(8),
varAlmacen VARCHAR(8),
varMoneda CHAR(3),
varFecha VARCHAR(20),
varNiveles VARCHAR(250),
varGlsNiveles VARCHAR(250),
VarOrdena       VarChar(30),
VarCodigoRapido        VarChar(50)
)
BEGIN

DECLARE varGlsSuc VARCHAR(180);
DECLARE varGlsAlm VARCHAR(180);
DECLARE varGlsMon VARCHAR(80);
DECLARE varGlsEmpresa VARCHAR(180);
DECLARE varGlsRuc VARCHAR(180);
DECLARE varGlsSistema VARCHAR(180);
DECLARE strSQL TEXT;
Declare VarStock            Text;
DECLARE varCodRapido VARCHAR(2);

If IfNull((SELECT valparametro FROM parametros WHERE idempresa = varEmpresa AND glsparametro = 'MUESTRA_STOCK_0'),'') <> 'S' Then

  Set VarStock = 'HAVING Round(STOCK,2) <> 0.00 ';

Else

  Set VarStock = '';

End If;

SET varCodRapido = IfNull((SELECT valparametro FROM parametros WHERE idempresa = varEmpresa AND glsparametro = 'VIZUALIZA_CODIGO_RAPIDO'),'');

If VarOrdena = 'C' Then

  Set VarOrdena = 'vd.IdProducto';

Else

  Set VarOrdena = 'p.GlsProducto';

End If;

IF varEmpresa <> '' THEN
SET varGlsEmpresa = (SELECT GlsEmpresa FROM empresas where idempresa = varEmpresa);
END IF;

IF varEmpresa <> '' THEN
SET varGlsRuc = (SELECT ruc FROM empresas where idempresa = varEmpresa);
END IF;

SET varGlsSistema = 'Sistema de Inventarios';


IF varSucursal <> '' THEN
SET varGlsSuc = (SELECT GlsPersona FROM Personas WHERE idPersona = varSucursal);
ELSE
SET varGlsSuc = 'TODAS LAS SUCURSALES';
END IF;


IF varAlmacen <> '' THEN
SET varGlsAlm = (SELECT GlsAlmacen FROM Almacenes WHERE idEmpresa = varEmpresa AND idAlmacen = varAlmacen);
ELSE
SET varGlsAlm = 'TODOS LOS ALMACENES';
END IF;

SET varFecha = cast(varFecha as date);

SET varGlsMon = (SELECT GlsMoneda FROM Monedas WHERE idMoneda = varMoneda);


Set @VarSQl = ConCat('SELECT vn.*, ',
'DATE_FORMAT(''',varFecha,''',''%d/%m/%Y'') AS FechaCorte, ',
'if(''',varCodRapido,''' = ''S'',Codigorapido,vd.idProducto) idProducto, ',
'p.GlsProducto, ',
'u.abreUM, ',
' ''',varGlsAlm,''' AS GlsAlmacen, ',
'SUM( ',
'IF((vc.tipoVale = ''I''),(vd.Cantidad),((vd.Cantidad) * -(1))) ',
') AS STOCK, ',

'If(Round(SUM( ',
'IF((vc.tipoVale = ''I''),(vd.Cantidad),((vd.Cantidad) * -(1))) ',
'),2) > 0,((SUM(IF((vc.tipoVale = ''I''),(vd.Cantidad),((vd.Cantidad) * -(1))) * ',
'CASE ''',varMoneda,''' ',
'WHEN ''PEN'' THEN IF(vc.idMoneda = ''PEN'', vd.VVUnit,vd.VVUnit * vc.TipoCambio) ',
'WHEN ''USD'' THEN IF(vc.idMoneda = ''USD'', vd.VVUnit,vd.VVUnit / vc.TipoCambio) End) / ',
'SUM(IF((vc.tipoVale = ''I''),(vd.Cantidad),((vd.Cantidad) * -(1)))))),0) AS Costo, ',

' ''',varGlsMon,''' AS GlsMoneda, ',
' ''',varGlsSuc,''' AS GlsSucursal, ',
' ''',varGlsEmpresa,''' AS varGlsEmpresa, ',
' ''',varGlsRuc,''' AS varGlsRuc, ',
' ''',varGlsSistema,''' AS varGlsSistema,P.IdFabricante,M.GlsMarca,vd.IdLote,vd.NumLote ',
'FROM valescab vc ',
'Inner Join valesdet vd ',
'On Vc.idValesCab = Vd.idValesCab And Vc.idEmpresa = Vd.idEmpresa And Vc.idSucursal = Vd.idSucursal And Vc.tipoVale = Vd.tipoVale ',
'Inner Join productos p ',
'On vd.idProducto = p.idProducto AND vd.idEmpresa = p.idEmpresa ',
'Inner Join unidadmedida u ',
'On p.idUMCompra = u.idUM ',
'Inner Join conceptos c ',
'On vc.idConcepto = c.idConcepto ',
'Inner Join vw_niveles vn ',
'On p.idNivel = vn.idNivel01 AND p.idEmpresa = vn.idEmpresa ',
'Left Join Marcas M ',
'On P.IdMarca = M.IdMarca And P.IdEmpresa = M.IdEmpresa ',
'Inner Join Almacenes A ',
'On Vc.IdEmpresa = A.IdEmpresa And Vc.IdAlmacen = A.IdAlmacen ',
'left join personas pe On vc.idProvCliente = pe.IdPersona ',
'Inner Join niveles n On p.idnivel = n.idnivel And p.idempresa = n.idempresa ',
'WHERE vc.idEmpresa = ''',varEmpresa,''' And p.estProducto = ''A'' AND vc.estValeCab <> ''ANU'' ',
'AND (vc.idSucursal = ''',varSucursal,''' OR ''',varSucursal,''' = '''') ',
'AND (vc.idPeriodoInv IN ',
'( ',
'SELECT pi.idPeriodoInv ',
'FROM periodosinv pi ',
'WHERE pi.idEmpresa = vc.idEmpresa AND pi.idSucursal = vc.idSucursal and pi.FecInicio <= ''',varFecha,''' ',
'AND (pi.FecFin >= ''',varFecha,''' or pi.FecFin is null) ',
') ',
') ',
'AND (vc.idAlmacen = ''',varAlmacen,''' OR ''',varAlmacen,''' = '''') ',
'AND vc.fechaEmision <= ''',varFecha,''' And (p.CodigoRapido = ''',VarCodigoRapido,''' Or '''' = ''',VarCodigoRapido,''') ',
varNiveles, '',
'GROUP BY vd.idProducto,p.GlsProducto,u.abreUM,vd.IdLote,vd.NumLote ',
VarStock,
'Order By ',VarOrdena,'');

-- select @VarSQl;
PREPARE strSQL FROM @VarSQl;
EXECUTE strSQL;
DEALLOCATE PREPARE strSQL;

END $$

DELIMITER ;