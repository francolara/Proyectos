
CREATE OR ALTER  PROCEDURE Spu_RegeneraSaldos_Proc
    @VarIdEmpresa CHAR(2),
    @VarFecha VARCHAR(50)
AS
BEGIN

    DECLARE @PeriodoAnt VARCHAR(50);
    SET @PeriodoAnt = CONVERT(VARCHAR(6), DATEADD(MONTH, -1, CONVERT(DATE, @VarFecha)), 112);

    -- ELIMINAMOS EL LA DATA DEL PERIODO ACTUAL
    DELETE FROM tbsaldo_costo_kardex
    WHERE Sc_Periodo = CONVERT(VARCHAR(6), CONVERT(DATE, @VarFecha), 112)
      AND (idEmpresa = @VarIdEmpresa OR @VarIdEmpresa = '');

    -- INSERTAMOS AL PERIODO ACTUAL TODOS LOS SALDOS DEL PERIODO ANTERIOR
    INSERT INTO tbsaldo_costo_kardex (Sc_periodo, sc_Codalm, Sc_Codart, Sc_Stock, idempresa)
    SELECT CONVERT(VARCHAR(6), CONVERT(DATE, @VarFecha), 112), sc_Codalm, Sc_Codart, Sc_Stock, idempresa
    FROM tbsaldo_costo_kardex
    WHERE Sc_Periodo = @PeriodoAnt
      AND (idEmpresa = @VarIdEmpresa OR @VarIdEmpresa = '');

    -- SACAMOS LOS SALDOS DEL MES
    CREATE TABLE #tmp_SaldoPeriodoAct (
        Periodo VARCHAR(6),
        idalmacen INT,
        IdProducto INT,
        Stock DECIMAL(14, 5),
        idempresa CHAR(2)
    );

    INSERT INTO #tmp_SaldoPeriodoAct (Periodo, idalmacen, IdProducto, Stock, idempresa)
    SELECT CONVERT(VARCHAR(6), CONVERT(DATE, @VarFecha), 112), a.idalmacen, C.IdProducto,
           CAST(ISNULL(SUM(B.Cantidad * CASE WHEN A.TipoVale = 'I' THEN 1 ELSE -1 END), 0) AS DECIMAL(14, 5)) AS Stock,
           a.idempresa
    FROM ValesCab A
    INNER JOIN ValesDet B ON A.IdEmpresa = B.IdEmpresa AND A.IdSucursal = B.IdSucursal AND A.TipoVale = B.TipoVale AND A.IdValesCab = B.IdValesCab
    INNER JOIN Productos C ON B.IdEmpresa = C.IdEmpresa AND B.IdProducto = C.IdProducto
    WHERE (A.IdEmpresa = @VarIdEmpresa OR @VarIdEmpresa = '')
      AND A.EstValeCab <> 'ANU'
      AND CONVERT(VARCHAR(6), A.fechaemision, 112) = CONVERT(VARCHAR(6), CONVERT(DATE, @VarFecha), 112)
    -- AND A.fechaemision < @VarFecha  -- Esta línea ha sido comentada también en el original
    GROUP BY a.idalmacen, C.IdProducto, a.idempresa;

    -- ACTUALIZAMOS LOS REGISTROS QUE ESTEN EN LA TABLA DE SALDOS
    UPDATE A
    SET A.Sc_Stock = A.Sc_Stock + B.Stock
    FROM tbsaldo_costo_kardex A
    INNER JOIN #tmp_SaldoPeriodoAct B ON A.Sc_periodo = B.Periodo AND A.sc_Codalm = B.idalmacen
                                      AND A.Sc_Codart = B.IdProducto AND A.idempresa = B.idempresa
    WHERE A.Sc_periodo = CONVERT(VARCHAR(6), CONVERT(DATE, @VarFecha), 112);

    -- INSERTAMOS LOS NUEVOS PRODUCTOS QUE TENGAN STOCK
    INSERT INTO tbsaldo_costo_kardex (Sc_periodo, sc_Codalm, Sc_Codart, Sc_Stock, idempresa)
    SELECT CONVERT(VARCHAR(6), CONVERT(DATE, @VarFecha), 112), a.idalmacen, a.IdProducto, a.Stock, a.idempresa
    FROM #tmp_SaldoPeriodoAct A
    LEFT JOIN tbsaldo_costo_kardex B ON B.Sc_periodo = A.Periodo AND B.sc_Codalm = A.idalmacen
                                     AND B.Sc_Codart = A.IdProducto AND B.idempresa = A.idempresa
    WHERE B.Sc_periodo IS NULL;

    -- INSERTAMOS LOGICO DE EJECUCION
    INSERT INTO tbsaldo_costo_kardex_Log (Fecha, IdEmpresaParam, FechaParam)
    VALUES (GETDATE(), @VarIdEmpresa, @VarFecha);

    -- ELIMINAMOS TEMPORAL
    DROP TABLE #tmp_SaldoPeriodoAct;
END;
GO
