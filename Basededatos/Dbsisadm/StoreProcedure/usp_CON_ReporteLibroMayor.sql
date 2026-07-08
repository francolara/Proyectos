-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   06/07/2026
-- Description:   Replica el libro mayor del legacy en HTML, usando NumeroDocumento como auxiliar funcional y conservando saldos iniciales por cuenta.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   08/07/2026
-- Description:   Ajusta Libro Mayor para trabajar por periodo contable, mostrar doble moneda y segmentar el saldo final solo al cierre de cada cuenta.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_ReporteLibroMayor
    @IdEmpresa INT,
    @Periodo CHAR(6),
    @CuentaDesde VARCHAR(20) = NULL,
    @CuentaHasta VARCHAR(20) = NULL,
    @NumeroDocumento VARCHAR(20) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @AnioPeriodo CHAR(4) = LEFT(@Periodo, 4);
        DECLARE @CuentaDesdeTrabajo VARCHAR(20);
        DECLARE @CuentaHastaTrabajo VARCHAR(20);
        DECLARE @NumeroDocumentoTrabajo VARCHAR(20) = NULLIF(LTRIM(RTRIM(@NumeroDocumento)), '');

        SELECT
            @CuentaDesdeTrabajo = COALESCE(NULLIF(LTRIM(RTRIM(@CuentaDesde)), ''), MIN(p.CodigoCuenta)),
            @CuentaHastaTrabajo = COALESCE(NULLIF(LTRIM(RTRIM(@CuentaHasta)), ''), MAX(p.CodigoCuenta))
        FROM dbo.CON_PlanCuenta AS p
        WHERE p.IdEmpresa = @IdEmpresa
          AND p.Estado = 1
          AND p.AceptaMovimiento = 1;

        CREATE TABLE #SaldosIniciales
        (
            CodigoCuenta VARCHAR(20) NOT NULL PRIMARY KEY,
            NombreCuenta NVARCHAR(200) NOT NULL,
            SaldoInicial DECIMAL(18, 2) NOT NULL,
            SaldoInicialDolares DECIMAL(18, 2) NOT NULL
        );

        INSERT INTO #SaldosIniciales (CodigoCuenta, NombreCuenta, SaldoInicial, SaldoInicialDolares)
        SELECT
            p.CodigoCuenta,
            p.NombreCuenta,
            SUM(CASE WHEN d.DH = 'H' THEN d.TotalImporteS * -1 ELSE d.TotalImporteS END),
            SUM(CASE WHEN d.DH = 'H' THEN d.TotalImporteD * -1 ELSE d.TotalImporteD END)
        FROM dbo.CON_Asiento AS a
        INNER JOIN dbo.CON_AsientoDetalle AS d
            ON d.IdAsiento = a.IdAsiento
        INNER JOIN dbo.CON_PlanCuenta AS p
            ON p.IdPlanCuenta = d.IdPlanCuenta
        WHERE a.IdEmpresa = @IdEmpresa
          AND a.Periodo < @Periodo
          AND LEFT(a.Periodo, 4) = @AnioPeriodo
          AND p.CodigoCuenta >= @CuentaDesdeTrabajo
          AND p.CodigoCuenta <= @CuentaHastaTrabajo
          AND (@NumeroDocumentoTrabajo IS NULL OR d.NumeroDocumento = @NumeroDocumentoTrabajo)
        GROUP BY
            p.CodigoCuenta,
            p.NombreCuenta;

        CREATE TABLE #Movimientos
        (
            CodigoCuenta VARCHAR(20) NOT NULL,
            NombreCuenta NVARCHAR(200) NOT NULL,
            CodigoOrigen VARCHAR(10) NOT NULL,
            NombreOrigen NVARCHAR(150) NOT NULL,
            Periodo CHAR(6) NOT NULL,
            NumeroAsiento INT NOT NULL,
            Item SMALLINT NOT NULL,
            FechaEmision DATE NOT NULL,
            TipoDocumento NVARCHAR(150) NOT NULL,
            Serie VARCHAR(10) NOT NULL,
            Referencia NVARCHAR(100) NOT NULL,
            NumeroDocumento VARCHAR(20) NOT NULL,
            NombreAuxiliar NVARCHAR(250) NOT NULL,
            Glosa NVARCHAR(500) NOT NULL,
            TipoCambio DECIMAL(18, 6) NOT NULL,
            Debe DECIMAL(18, 2) NOT NULL,
            Haber DECIMAL(18, 2) NOT NULL,
            DebeDolares DECIMAL(18, 2) NOT NULL,
            HaberDolares DECIMAL(18, 2) NOT NULL
        );

        INSERT INTO #Movimientos
        (
            CodigoCuenta,
            NombreCuenta,
            CodigoOrigen,
            NombreOrigen,
            Periodo,
            NumeroAsiento,
            Item,
            FechaEmision,
            TipoDocumento,
            Serie,
            Referencia,
            NumeroDocumento,
            NombreAuxiliar,
            Glosa,
            TipoCambio,
            Debe,
            Haber,
            DebeDolares,
            HaberDolares
        )
        SELECT
            p.CodigoCuenta,
            p.NombreCuenta,
            o.CodigoOrigen,
            o.NombreOrigen,
            a.Periodo,
            a.NumeroAsiento,
            d.Item,
            a.FechaEmision,
            ISNULL(NULLIF(LTRIM(RTRIM(d.TipoDocumento)), ''), ''),
            ISNULL(NULLIF(LTRIM(RTRIM(d.Serie)), ''), ''),
            ISNULL(NULLIF(LTRIM(RTRIM(d.ReferenciaLinea)), ''), ''),
            ISNULL(NULLIF(LTRIM(RTRIM(d.NumeroDocumento)), ''), ''),
            ISNULL(NULLIF(LTRIM(RTRIM(COALESCE(per.NombreCompleto, per.RazonSocial, d.NumeroDocumento))), ''), ''),
            COALESCE(NULLIF(LTRIM(RTRIM(d.GlosaDetalle)), ''), a.Glosa, N''),
            ISNULL(NULLIF(d.TipoCambioLinea, 0), a.TipoCambio),
            CASE WHEN d.DH = 'D' THEN d.TotalImporteS ELSE 0 END,
            CASE WHEN d.DH = 'H' THEN d.TotalImporteS ELSE 0 END,
            CASE WHEN d.DH = 'D' THEN d.TotalImporteD ELSE 0 END,
            CASE WHEN d.DH = 'H' THEN d.TotalImporteD ELSE 0 END
        FROM dbo.CON_Asiento AS a
        INNER JOIN dbo.CON_AsientoDetalle AS d
            ON d.IdAsiento = a.IdAsiento
        INNER JOIN dbo.CON_PlanCuenta AS p
            ON p.IdPlanCuenta = d.IdPlanCuenta
        INNER JOIN dbo.CON_Origen AS o
            ON o.IdOrigen = a.IdOrigen
        LEFT JOIN dbo.ADM_Persona AS per
            ON per.IdEmpresa = a.IdEmpresa
           AND per.NumeroDocumento = d.NumeroDocumento
        WHERE a.IdEmpresa = @IdEmpresa
          AND a.Periodo = @Periodo
          AND p.CodigoCuenta >= @CuentaDesdeTrabajo
          AND p.CodigoCuenta <= @CuentaHastaTrabajo
          AND (@NumeroDocumentoTrabajo IS NULL OR d.NumeroDocumento = @NumeroDocumentoTrabajo);

        SELECT
            c.CodigoCuenta,
            c.NombreCuenta,
            CAST('' AS VARCHAR(10)) AS CodigoOrigen,
            CAST(N'Saldo inicial' AS NVARCHAR(150)) AS NombreOrigen,
            CAST('' AS CHAR(6)) AS Periodo,
            0 AS NumeroAsiento,
            CAST(0 AS SMALLINT) AS Item,
            CAST(NULL AS DATE) AS FechaEmision,
            CAST(N'' AS NVARCHAR(150)) AS TipoDocumento,
            CAST('' AS VARCHAR(10)) AS Serie,
            CAST(N'' AS NVARCHAR(100)) AS Referencia,
            CAST('' AS VARCHAR(20)) AS NumeroDocumento,
            CAST(N'' AS NVARCHAR(250)) AS NombreAuxiliar,
            CAST(N'Saldo acumulado anterior al rango consultado.' AS NVARCHAR(500)) AS Glosa,
            CAST(0 AS DECIMAL(18, 6)) AS TipoCambio,
            CAST(0 AS DECIMAL(18, 2)) AS Debe,
            CAST(0 AS DECIMAL(18, 2)) AS Haber,
            CAST(0 AS DECIMAL(18, 2)) AS DebeDolares,
            CAST(0 AS DECIMAL(18, 2)) AS HaberDolares,
            ISNULL(s.SaldoInicial, 0) AS SaldoInicial,
            ISNULL(s.SaldoInicialDolares, 0) AS SaldoInicialDolares,
            CAST(1 AS BIT) AS EsSaldoInicial
        FROM
        (
            SELECT p.CodigoCuenta, p.NombreCuenta
            FROM dbo.CON_PlanCuenta AS p
            WHERE p.IdEmpresa = @IdEmpresa
              AND p.Estado = 1
              AND p.AceptaMovimiento = 1
              AND p.CodigoCuenta >= @CuentaDesdeTrabajo
              AND p.CodigoCuenta <= @CuentaHastaTrabajo
        ) AS c
        LEFT JOIN #SaldosIniciales AS s
            ON s.CodigoCuenta = c.CodigoCuenta
        WHERE EXISTS
        (
            SELECT 1
            FROM #Movimientos AS m
            WHERE m.CodigoCuenta = c.CodigoCuenta
        )
           OR EXISTS
        (
            SELECT 1
            FROM #SaldosIniciales AS x
            WHERE x.CodigoCuenta = c.CodigoCuenta
              AND (x.SaldoInicial <> 0 OR x.SaldoInicialDolares <> 0)
        )

        UNION ALL

        SELECT
            m.CodigoCuenta,
            m.NombreCuenta,
            m.CodigoOrigen,
            m.NombreOrigen,
            m.Periodo,
            m.NumeroAsiento,
            m.Item,
            m.FechaEmision,
            m.TipoDocumento,
            m.Serie,
            m.Referencia,
            m.NumeroDocumento,
            m.NombreAuxiliar,
            m.Glosa,
            m.TipoCambio,
            m.Debe,
            m.Haber,
            m.DebeDolares,
            m.HaberDolares,
            ISNULL(s.SaldoInicial, 0) AS SaldoInicial,
            ISNULL(s.SaldoInicialDolares, 0) AS SaldoInicialDolares,
            CAST(0 AS BIT) AS EsSaldoInicial
        FROM #Movimientos AS m
        LEFT JOIN #SaldosIniciales AS s
            ON s.CodigoCuenta = m.CodigoCuenta
        ORDER BY
            CodigoCuenta,
            EsSaldoInicial DESC,
            FechaEmision,
            NumeroAsiento,
            Item;

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
