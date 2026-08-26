-- =============================================
-- Author:        FRANCO LARA
-- Create date:   13/08/2026
-- Description:   Adapta el proceso de cierre para generar un unico asiento compuesto entre periodos contables seleccionados.
-- =============================================
-- Firma: FRANCO LARA - 13/08/2026 | Agrega periodo de corte/generacion, vinculo al asiento unico y datos monetarios por linea, conservando compatibilidad con cierres ya registrados.
-- Firma: FRANCO LARA - 22/08/2026 | Consolida el calendario contable 00-14, fija el cierre de Inventario en 14 y limita el corte del cierre a 00-13.

IF COL_LENGTH(N'dbo.CON_CierreProceso', N'MesSaldoHasta') IS NULL
BEGIN
    ALTER TABLE dbo.CON_CierreProceso
        ADD MesSaldoHasta TINYINT NULL;
END;

IF COL_LENGTH(N'dbo.CON_CierreProceso', N'PeriodoSaldoHasta') IS NULL
BEGIN
    ALTER TABLE dbo.CON_CierreProceso
        ADD PeriodoSaldoHasta CHAR(6) NULL;
END;

IF COL_LENGTH(N'dbo.CON_CierreProceso', N'MesGeneracion') IS NULL
BEGIN
    ALTER TABLE dbo.CON_CierreProceso
        ADD MesGeneracion TINYINT NULL;
END;

IF COL_LENGTH(N'dbo.CON_CierreProceso', N'PeriodoGeneracion') IS NULL
BEGIN
    ALTER TABLE dbo.CON_CierreProceso
        ADD PeriodoGeneracion CHAR(6) NULL;
END;

IF COL_LENGTH(N'dbo.CON_CierreProceso', N'IdAsiento') IS NULL
BEGIN
    ALTER TABLE dbo.CON_CierreProceso
        ADD IdAsiento INT NULL;
END;

IF COL_LENGTH(N'dbo.CON_CierreProceso', N'NumeroAsiento') IS NULL
BEGIN
    ALTER TABLE dbo.CON_CierreProceso
        ADD NumeroAsiento INT NULL;
END;

IF COL_LENGTH(N'dbo.CON_CierreProceso', N'TotalLineas') IS NULL
BEGIN
    ALTER TABLE dbo.CON_CierreProceso
        ADD TotalLineas INT NOT NULL CONSTRAINT DF_CON_CierreProceso_TotalLineas DEFAULT (0);
END;

UPDATE p
SET MesSaldoHasta = ISNULL(p.MesSaldoHasta, 13),
    PeriodoSaldoHasta = ISNULL(p.PeriodoSaldoHasta, CONCAT(p.Anio, '13')),
    MesGeneracion = ISNULL(p.MesGeneracion, 14),
    PeriodoGeneracion = ISNULL(
        p.PeriodoGeneracion,
        CONCAT(p.Anio, '14')
    )
FROM dbo.CON_CierreProceso AS p;

ALTER TABLE dbo.CON_CierreProceso ALTER COLUMN MesSaldoHasta TINYINT NOT NULL;
ALTER TABLE dbo.CON_CierreProceso ALTER COLUMN PeriodoSaldoHasta CHAR(6) NOT NULL;
ALTER TABLE dbo.CON_CierreProceso ALTER COLUMN MesGeneracion TINYINT NOT NULL;
ALTER TABLE dbo.CON_CierreProceso ALTER COLUMN PeriodoGeneracion CHAR(6) NOT NULL;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.default_constraints
    WHERE parent_object_id = OBJECT_ID(N'dbo.CON_CierreProceso')
      AND name = N'DF_CON_CierreProceso_MesSaldoHasta'
)
BEGIN
    ALTER TABLE dbo.CON_CierreProceso
        ADD CONSTRAINT DF_CON_CierreProceso_MesSaldoHasta DEFAULT (13) FOR MesSaldoHasta;
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.default_constraints
    WHERE parent_object_id = OBJECT_ID(N'dbo.CON_CierreProceso')
      AND name = N'DF_CON_CierreProceso_MesGeneracion'
)
BEGIN
    ALTER TABLE dbo.CON_CierreProceso
        ADD CONSTRAINT DF_CON_CierreProceso_MesGeneracion DEFAULT (14) FOR MesGeneracion;
END;

IF OBJECT_ID(N'dbo.CK_CON_CierreProceso_Meses', N'C') IS NOT NULL
BEGIN
    ALTER TABLE dbo.CON_CierreProceso
        DROP CONSTRAINT CK_CON_CierreProceso_Meses;
END;

ALTER TABLE dbo.CON_CierreProceso
    ADD CONSTRAINT CK_CON_CierreProceso_Meses
        CHECK (MesSaldoHasta BETWEEN 0 AND 13 AND MesGeneracion = 14);

IF OBJECT_ID(N'dbo.CK_CON_CierreProceso_Periodos', N'C') IS NULL
BEGIN
    ALTER TABLE dbo.CON_CierreProceso
        ADD CONSTRAINT CK_CON_CierreProceso_Periodos
            CHECK (
                PeriodoSaldoHasta = CONVERT(CHAR(4), Anio) + RIGHT('0' + CONVERT(VARCHAR(2), MesSaldoHasta), 2)
                AND PeriodoGeneracion = CONVERT(CHAR(4), Anio) + RIGHT('0' + CONVERT(VARCHAR(2), MesGeneracion), 2)
            );
END;

IF OBJECT_ID(N'dbo.FK_CON_CierreProceso_CON_Asiento', N'F') IS NULL
BEGIN
    ALTER TABLE dbo.CON_CierreProceso
        ADD CONSTRAINT FK_CON_CierreProceso_CON_Asiento
            FOREIGN KEY (IdAsiento) REFERENCES dbo.CON_Asiento (IdAsiento);
END;

IF COL_LENGTH(N'dbo.CON_CierreProcesoDetalle', N'Item') IS NULL
BEGIN
    ALTER TABLE dbo.CON_CierreProcesoDetalle
        ADD Item SMALLINT NULL;

    ;WITH LineasNumeradas AS
    (
        SELECT
            d.IdCierreProcesoDetalle,
            ROW_NUMBER() OVER (PARTITION BY d.IdCierreProceso ORDER BY d.TipoCierre, d.IdPlanCuenta, d.IdCierreProcesoDetalle) AS ItemCalculado
        FROM dbo.CON_CierreProcesoDetalle AS d
    )
    UPDATE d
    SET Item = n.ItemCalculado
    FROM dbo.CON_CierreProcesoDetalle AS d
    INNER JOIN LineasNumeradas AS n
        ON n.IdCierreProcesoDetalle = d.IdCierreProcesoDetalle;

    ALTER TABLE dbo.CON_CierreProcesoDetalle ALTER COLUMN Item SMALLINT NOT NULL;
END;

IF COL_LENGTH(N'dbo.CON_CierreProcesoDetalle', N'DH') IS NULL
BEGIN
    ALTER TABLE dbo.CON_CierreProcesoDetalle
        ADD DH CHAR(1) NULL;

    UPDATE d
    SET DH = CASE WHEN d.TotalDebe > 0 THEN 'D' ELSE 'H' END
    FROM dbo.CON_CierreProcesoDetalle AS d;

    ALTER TABLE dbo.CON_CierreProcesoDetalle ALTER COLUMN DH CHAR(1) NOT NULL;
END;

IF COL_LENGTH(N'dbo.CON_CierreProcesoDetalle', N'TotalImporteS') IS NULL
BEGIN
    ALTER TABLE dbo.CON_CierreProcesoDetalle
        ADD TotalImporteS DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_CierreProcesoDetalle_TotalImporteS DEFAULT (0);

    UPDATE d
    SET TotalImporteS = d.TotalDebe + d.TotalHaber
    FROM dbo.CON_CierreProcesoDetalle AS d;
END;

IF COL_LENGTH(N'dbo.CON_CierreProcesoDetalle', N'TotalImporteD') IS NULL
BEGIN
    ALTER TABLE dbo.CON_CierreProcesoDetalle
        ADD TotalImporteD DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_CierreProcesoDetalle_TotalImporteD DEFAULT (0);
END;

IF OBJECT_ID(N'dbo.CK_CON_CierreProcesoDetalle_TipoCierre', N'C') IS NOT NULL
BEGIN
    ALTER TABLE dbo.CON_CierreProcesoDetalle
        DROP CONSTRAINT CK_CON_CierreProcesoDetalle_TipoCierre;
END;

ALTER TABLE dbo.CON_CierreProcesoDetalle
    ADD CONSTRAINT CK_CON_CierreProcesoDetalle_TipoCierre
        CHECK (TipoCierre = '14');

IF OBJECT_ID(N'dbo.CK_CON_CierreProcesoDetalle_Item', N'C') IS NULL
BEGIN
    ALTER TABLE dbo.CON_CierreProcesoDetalle
        ADD CONSTRAINT CK_CON_CierreProcesoDetalle_Item CHECK (Item >= 1);
END;

IF OBJECT_ID(N'dbo.CK_CON_CierreProcesoDetalle_DH', N'C') IS NULL
BEGIN
    ALTER TABLE dbo.CON_CierreProcesoDetalle
        ADD CONSTRAINT CK_CON_CierreProcesoDetalle_DH CHECK (DH IN ('D', 'H'));
END;

UPDATE p
SET TotalLineas = ISNULL(resumen.TotalLineas, 0),
    IdAsiento = CASE WHEN resumen.TotalAsientosVinculados = 1 THEN resumen.IdAsiento ELSE NULL END,
    NumeroAsiento = CASE WHEN resumen.TotalAsientosVinculados = 1 THEN a.NumeroAsiento ELSE NULL END
FROM dbo.CON_CierreProceso AS p
OUTER APPLY
(
    SELECT
        COUNT(d.IdCierreProcesoDetalle) AS TotalLineas,
        COUNT(DISTINCT d.IdAsiento) AS TotalAsientosVinculados,
        MAX(d.IdAsiento) AS IdAsiento
    FROM dbo.CON_CierreProcesoDetalle AS d
    WHERE d.IdCierreProceso = p.IdCierreProceso
) AS resumen
LEFT JOIN dbo.CON_Asiento AS a
    ON a.IdAsiento = resumen.IdAsiento;

IF OBJECT_ID(N'dbo.CK_CON_Asiento_Mes', N'C') IS NOT NULL
BEGIN
    ALTER TABLE dbo.CON_Asiento
        DROP CONSTRAINT CK_CON_Asiento_Mes;
END;

ALTER TABLE dbo.CON_Asiento
    ADD CONSTRAINT CK_CON_Asiento_Mes
        CHECK (Mes BETWEEN 0 AND 14);

IF OBJECT_ID(N'dbo.CK_CON_CorrelativoAsiento_Periodo', N'C') IS NOT NULL
BEGIN
    ALTER TABLE dbo.CON_CorrelativoAsiento
        DROP CONSTRAINT CK_CON_CorrelativoAsiento_Periodo;
END;

ALTER TABLE dbo.CON_CorrelativoAsiento
    ADD CONSTRAINT CK_CON_CorrelativoAsiento_Periodo
        CHECK (
            Periodo LIKE '[1-2][0-9][0-9][0-9][0-1][0-9]'
            AND RIGHT(Periodo, 2) BETWEEN '00' AND '14'
        );

IF OBJECT_ID(N'dbo.CK_CON_AperturaProceso_MesSaldo', N'C') IS NOT NULL
BEGIN
    ALTER TABLE dbo.CON_AperturaProceso
        DROP CONSTRAINT CK_CON_AperturaProceso_MesSaldo;
END;

ALTER TABLE dbo.CON_AperturaProceso
    ADD CONSTRAINT CK_CON_AperturaProceso_MesSaldo
        CHECK (MesSaldoHasta BETWEEN 0 AND 14);

IF OBJECT_ID(N'dbo.CK_CON_AperturaProceso_PeriodoSaldo', N'C') IS NOT NULL
BEGIN
    ALTER TABLE dbo.CON_AperturaProceso
        DROP CONSTRAINT CK_CON_AperturaProceso_PeriodoSaldo;
END;

ALTER TABLE dbo.CON_AperturaProceso
    ADD CONSTRAINT CK_CON_AperturaProceso_PeriodoSaldo
        CHECK (
            PeriodoSaldoHasta LIKE '[1-2][0-9][0-9][0-9][0-1][0-9]'
            AND RIGHT(PeriodoSaldoHasta, 2) BETWEEN '00' AND '14'
        );
