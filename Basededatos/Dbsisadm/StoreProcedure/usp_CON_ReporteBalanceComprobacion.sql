-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   07/07/2026
-- Description:   Replica el balance de comprobacion legacy en HTML usando rango de periodos, grado y jerarquia contable sobre CON_Asiento y CON_AsientoDetalle.
-- =============================================
-- Firma: FRANCO LARA - 22/08/2026 | Limita el rango del reporte al calendario contable vigente 00-14.

CREATE OR ALTER PROCEDURE dbo.usp_CON_ReporteBalanceComprobacion
    @IdEmpresa INT,
    @Anio SMALLINT,
    @PeriodoDesde TINYINT,
    @PeriodoHasta TINYINT,
    @Moneda VARCHAR(3) = 'PEN',
    @Grado TINYINT = 1,
    @TodasLasCuentas BIT = 1,
    @CuentaDesde VARCHAR(20) = NULL,
    @CuentaHasta VARCHAR(20) = NULL,
    @FiltrarGrado BIT = 1
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @MonedaTrabajo VARCHAR(3) = CASE WHEN UPPER(LTRIM(RTRIM(ISNULL(@Moneda, 'PEN')))) = 'USD' THEN 'USD' ELSE 'PEN' END;
        DECLARE @PeriodoDesdeTrabajo TINYINT = CASE WHEN @PeriodoDesde <= 14 THEN @PeriodoDesde ELSE 0 END;
        DECLARE @PeriodoHastaTrabajo TINYINT = CASE WHEN @PeriodoHasta <= 14 THEN @PeriodoHasta ELSE @PeriodoDesdeTrabajo END;
        DECLARE @GradoTrabajo TINYINT = CASE WHEN @Grado >= 1 THEN @Grado ELSE 1 END;
        DECLARE @AnioTrabajo CHAR(4) = RIGHT('0000' + CONVERT(VARCHAR(4), @Anio), 4);
        DECLARE @CuentaDesdeTrabajo VARCHAR(20);
        DECLARE @CuentaHastaTrabajo VARCHAR(20);

        IF @PeriodoHastaTrabajo < @PeriodoDesdeTrabajo
        BEGIN
            DECLARE @PeriodoSwap TINYINT = @PeriodoDesdeTrabajo;
            SET @PeriodoDesdeTrabajo = @PeriodoHastaTrabajo;
            SET @PeriodoHastaTrabajo = @PeriodoSwap;
        END;

        SELECT
            @CuentaDesdeTrabajo = COALESCE(NULLIF(LTRIM(RTRIM(@CuentaDesde)), ''), MIN(p.CodigoCuenta)),
            @CuentaHastaTrabajo = COALESCE(NULLIF(LTRIM(RTRIM(@CuentaHasta)), ''), MAX(p.CodigoCuenta))
        FROM dbo.CON_PlanCuenta AS p
        WHERE p.IdEmpresa = @IdEmpresa
          AND p.Estado = 1;

        IF @CuentaDesdeTrabajo IS NULL OR @CuentaHastaTrabajo IS NULL
        BEGIN
            SELECT
                CAST('' AS VARCHAR(20)) AS CodigoCuenta,
                CAST(N'' AS NVARCHAR(200)) AS NombreCuenta,
                CAST('' AS VARCHAR(1)) AS ColBalance,
                CAST(0 AS TINYINT) AS GradoCuenta,
                CAST(0 AS DECIMAL(18, 2)) AS DebAnt,
                CAST(0 AS DECIMAL(18, 2)) AS HabAnt,
                CAST(0 AS DECIMAL(18, 2)) AS DebMes,
                CAST(0 AS DECIMAL(18, 2)) AS HabMes,
                CAST(0 AS DECIMAL(18, 2)) AS Debe,
                CAST(0 AS DECIMAL(18, 2)) AS Haber
            WHERE 1 = 0;

            RETURN;
        END;

        CREATE TABLE #Movimientos
        (
            CodigoCuentaMovimiento VARCHAR(20) NOT NULL PRIMARY KEY,
            DebAnt DECIMAL(18, 2) NOT NULL,
            HabAnt DECIMAL(18, 2) NOT NULL,
            DebMes DECIMAL(18, 2) NOT NULL,
            HabMes DECIMAL(18, 2) NOT NULL,
            Debe DECIMAL(18, 2) NOT NULL,
            Haber DECIMAL(18, 2) NOT NULL
        );

        INSERT INTO #Movimientos
        (
            CodigoCuentaMovimiento,
            DebAnt,
            HabAnt,
            DebMes,
            HabMes,
            Debe,
            Haber
        )
        SELECT
            p.CodigoCuenta,
            SUM(CASE WHEN d.DH = 'D' AND TRY_CONVERT(TINYINT, RIGHT(a.Periodo, 2)) < @PeriodoHastaTrabajo THEN CASE WHEN @MonedaTrabajo = 'USD' THEN d.TotalImporteD ELSE d.TotalImporteS END ELSE 0 END),
            SUM(CASE WHEN d.DH = 'H' AND TRY_CONVERT(TINYINT, RIGHT(a.Periodo, 2)) < @PeriodoHastaTrabajo THEN CASE WHEN @MonedaTrabajo = 'USD' THEN d.TotalImporteD ELSE d.TotalImporteS END ELSE 0 END),
            SUM(CASE WHEN d.DH = 'D' AND TRY_CONVERT(TINYINT, RIGHT(a.Periodo, 2)) = @PeriodoHastaTrabajo THEN CASE WHEN @MonedaTrabajo = 'USD' THEN d.TotalImporteD ELSE d.TotalImporteS END ELSE 0 END),
            SUM(CASE WHEN d.DH = 'H' AND TRY_CONVERT(TINYINT, RIGHT(a.Periodo, 2)) = @PeriodoHastaTrabajo THEN CASE WHEN @MonedaTrabajo = 'USD' THEN d.TotalImporteD ELSE d.TotalImporteS END ELSE 0 END),
            SUM(CASE WHEN d.DH = 'D' THEN CASE WHEN @MonedaTrabajo = 'USD' THEN d.TotalImporteD ELSE d.TotalImporteS END ELSE 0 END),
            SUM(CASE WHEN d.DH = 'H' THEN CASE WHEN @MonedaTrabajo = 'USD' THEN d.TotalImporteD ELSE d.TotalImporteS END ELSE 0 END)
        FROM dbo.CON_Asiento AS a
        INNER JOIN dbo.CON_AsientoDetalle AS d
            ON d.IdAsiento = a.IdAsiento
        INNER JOIN dbo.CON_PlanCuenta AS p
            ON p.IdPlanCuenta = d.IdPlanCuenta
        WHERE a.IdEmpresa = @IdEmpresa
          AND LEFT(a.Periodo, 4) = @AnioTrabajo
          AND TRY_CONVERT(TINYINT, RIGHT(a.Periodo, 2)) BETWEEN @PeriodoDesdeTrabajo AND @PeriodoHastaTrabajo
          AND p.Estado = 1
          AND p.AceptaMovimiento = 1
        GROUP BY
            p.CodigoCuenta;

        SELECT
            pc.CodigoCuenta,
            pc.NombreCuenta,
            pc.ColBalance,
            CAST(pc.NivelCuenta AS TINYINT) AS GradoCuenta,
            SUM(m.DebAnt) AS DebAnt,
            SUM(m.HabAnt) AS HabAnt,
            SUM(m.DebMes) AS DebMes,
            SUM(m.HabMes) AS HabMes,
            SUM(m.Debe) AS Debe,
            SUM(m.Haber) AS Haber
        FROM dbo.CON_PlanCuenta AS pc
        INNER JOIN #Movimientos AS m
            ON m.CodigoCuentaMovimiento LIKE pc.CodigoCuenta + '%'
        WHERE pc.IdEmpresa = @IdEmpresa
          AND pc.Estado = 1
          AND pc.NivelCuenta <= @GradoTrabajo
          AND (
                @TodasLasCuentas = 1
                OR (
                    pc.CodigoCuenta >= @CuentaDesdeTrabajo
                    AND pc.CodigoCuenta <= @CuentaHastaTrabajo
                )
              )
          AND (@FiltrarGrado = 0 OR pc.NivelCuenta = @GradoTrabajo)
        GROUP BY
            pc.CodigoCuenta,
            pc.NombreCuenta,
            pc.ColBalance,
            pc.NivelCuenta
        HAVING SUM(m.Debe) <> 0 OR SUM(m.Haber) <> 0
        ORDER BY
            pc.CodigoCuenta;

    END TRY

    BEGIN CATCH

        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);

    END CATCH

END
