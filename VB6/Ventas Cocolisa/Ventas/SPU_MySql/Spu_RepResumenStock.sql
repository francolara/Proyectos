DELIMITER $$

-- Descripcion de version 1  : Reporte Resumen Stock
-- Fecha de Version 1		 : 18/10/2012
-- Creador de version 1		 : SAV Peru
-- Fecha de Pase Version 1	 : 18/10/2012
-- Autorizador Version 1	 : SAV Peru
-- Descripción               : Reporte Resumen Stock
-- Llamado                   : Sistema Inventarios


DROP PROCEDURE IF EXISTS `Spu_RepResumenStock` $$
CREATE  PROCEDURE `Spu_RepResumenStock`(
varEmpresa         CHAR(2),
varAlmacen         CHAR(8),
varProducto        CHAR(8),
varFecDesde        VARCHAR(10),
varFecHasta        VARCHAR(10)
)
BEGIN


DECLARE varGlsEmpresa      VARCHAR(180);
DECLARE varGlsRuc          VARCHAR(180);
DECLARE varGlsMoneda       VARCHAR(50);
DECLARE varGlsAlm          VARCHAR(180);

IF varEmpresa <> '' THEN
  SET varGlsEmpresa = (SELECT GlsEmpresa FROM empresas where idempresa = varEmpresa);
END IF;

IF varEmpresa <> '' THEN
  SET varGlsRuc = (SELECT ruc FROM empresas where idempresa = varEmpresa);
END IF;

IF varAlmacen = '' THEN
  SET varGlsAlm = 'TODOS LOS ALMACENES';
ELSE
  SET varGlsAlm = (SELECT GlsAlmacen FROM Almacenes WHERE idEmpresa = varEmpresa AND idAlmacen = varAlmacen);
END IF;

IF varAlmacen = '' THEN
  SET varAlmacen = '%%';
END IF;

IF varProducto = '' THEN
  SET varProducto = '%%';
END IF;

Select  p.idProducto,p.GlsProducto,um.abreUM,Ingresos,Salidas,SaldoInicial,(SaldoInicial + (Ingresos - Salidas))  As Saldo,varGlsAlm,varGlsEmpresa,varGlsRuc,
DATE_FORMAT(varFecDesde,'%d/%m/%Y') AS FechaINI,DATE_FORMAT(varFecHasta,'%d/%m/%Y') AS FechaFIN
From Productos p
Inner Join (
Select vc.IdAlmacen,vc.IdEmpresa,vd.idProducto,
Sum(If(vc.FechaEmision >= varFecDesde,If(vc.tipoVale = 'I',vd.Cantidad,0),0)) As Ingresos,
Sum(If(vc.FechaEmision >= varFecDesde,If(vc.tipoVale = 'S',vd.Cantidad,0),0)) As Salidas,
Sum(If(vc.FechaEmision < varFecDesde,If(vc.tipoVale = 'I',vd.Cantidad,(vd.Cantidad * -1)),0)) As SaldoInicial
From ValesCab vc
Inner Join ValesDet vd
  On  vc.idEmpresa = vd.idEmpresa
  And vc.idSucursal = vd.idSucursal
  And vc.tipoVale = vd.tipoVale
  And vc.idValesCab = vd.idValesCab
Where vc.idEmpresa = varEmpresa
And (vc.idPeriodoInv) IN (SELECT pi.idPeriodoInv
                            FROM periodosinv pi
                            WHERE pi.idEmpresa = vc.idEmpresa AND pi.idSucursal = vc.idSucursal and pi.FecInicio <= varFecHasta
                            and (pi.FecFin >= varFecHasta or pi.FecFin is null))
And vc.estValeCab <> 'ANU'
And vc.idAlmacen  Like varAlmacen
And vc.FechaEmision <= varFecHasta
Group By vd.idProducto) S
  On  S.idProducto = p.idProducto And S.idEmpresa = p.idEmpresa
Inner Join unidadMedida um
  On p.idUMCompra = um.idUM
Where  p.idProducto Like varProducto
Having Saldo > 0
Order by p.GlsProducto;

END $$

DELIMITER ;