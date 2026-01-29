DROP PROCEDURE IF EXISTS `Spu_RegeneraSaldos_Proc`;
DELIMITER $$

CREATE PROCEDURE `Spu_RegeneraSaldos_Proc`(
VarIdEmpresa                 Char(2),
VarFecha                     VarChar(50)
)
BEGIN

DECLARE PeriodoAnt        VARCHAR(50);
SET PeriodoAnt =   (SELECT DATE_FORMAT(DATE_add(VarFecha, INTERVAL -1 month), '%Y%m'));

-- ELIMINAMOS EL LA DATA DEL PERIODO ACTUAL
DELETE FROM tbsaldo_costo_kardex
WHERE Sc_Periodo = DATE_FORMAT(VarFecha, '%Y%m')
AND (idEmpresa = VarIdEmpresa OR VarIdEmpresa  = '');

-- INSERTAMOS AL PERIODO ACTUAL TODOS LOS SALDOS DEL PERIODO ANTERIOR
INSERT INTO  tbsaldo_costo_kardex (Sc_periodo,sc_Codalm,Sc_Codart,Sc_Stock,idempresa)
SELECT DATE_FORMAT(VarFecha, '%Y%m'),sc_Codalm,Sc_Codart,Sc_Stock,idempresa FROM tbsaldo_costo_kardex
WHERE Sc_Periodo = PeriodoAnt    AND (idEmpresa = VarIdEmpresa OR VarIdEmpresa  = '');

 -- SACAMOS LOS SALDOS DEL MES
  CREATE TEMPORARY TABLE tmp_SaldoPeriodoAct AS    (    SELECT DATE_FORMAT(VarFecha, '%Y%m') Periodo,a.idalmacen,C.IdProducto,
  Cast(IfNull(Sum(B.Cantidad * If(A.TipoVale = 'I',1,-1)),0) As Decimal(14,5)) Stock,    a.idempresa
  FROM ValesCab A
  INNER JOIN ValesDet B    ON A.IdEmpresa = B.IdEmpresa AND A.IdSucursal = B.IdSucursal AND A.TipoVale = B.TipoVale AND A.IdValesCab = B.IdValesCab
  INNER JOIN Productos C    ON B.IdEmpresa = C.IdEmpresa AND B.IdProducto = C.IdProducto    WHERE (A.IdEmpresa = VarIdEmpresa OR VarIdEmpresa  = '')
  AND A.EstValeCab <> 'ANU'    AND  DATE_FORMAT(A.fechaemision, '%Y%m')  = DATE_FORMAT(VarFecha, '%Y%m')
  -- AND  A.fechaemision  < VarFecha
  GROUP BY DATE_FORMAT(VarFecha, '%Y%m'),a.idalmacen,C.IdProducto,a.idempresa    ) ;
   -- ACTUALIZAMOS LOS REGISTROS QUE ESTEN EN LA TABLA DE SALDOS
   UPDATE tbsaldo_costo_kardex A
   INNER JOIN tmp_SaldoPeriodoAct B    ON  A.Sc_periodo =  B.Periodo and A.sc_Codalm =  B.idalmacen
   and A.Sc_Codart = B.IdProducto and A.idempresa = B.idempresa    SET A.Sc_Stock = A.Sc_Stock + B.Stock
   WHERE A.Sc_periodo = DATE_FORMAT(VarFecha, '%Y%m') ;
    -- INSERTAMOS LOS NUEVOS PRODUCTOS QUE TENGAN STOCK
   INSERT INTO  tbsaldo_costo_kardex (Sc_periodo,sc_Codalm,Sc_Codart,Sc_Stock,idempresa)
   SELECT DATE_FORMAT(VarFecha, '%Y%m'),a.idalmacen,a.IdProducto,a.Stock,a.idempresa
   FROM tmp_SaldoPeriodoAct A    LEFT JOIN tbsaldo_costo_kardex B    ON  b.Sc_periodo =  a.Periodo and b.sc_Codalm =  a.idalmacen
   AND b.Sc_Codart = a.IdProducto and b.idempresa = a.idempresa    WHERE b.Sc_periodo IS NULL;
   -- INSERTAMOS LOGICO DE EJECUCION
   INSERT INTO tbsaldo_costo_kardex_Log (Fecha,IdEmpresaParam,FechaParam)
   VALUES(SYSDATE(),VarIdEmpresa,VarFecha);
   -- ELIMINAMOS TEMPORAL
   DROP TABLE tmp_SaldoPeriodoAct;

END $$

DELIMITER ;
