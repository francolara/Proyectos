-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   06/07/2026
-- Description:   Replica el reporte legacy Analisis de cuentas sobre CON_Asiento y CON_AsientoDetalle usando NumeroDocumento como auxiliar funcional y la clave documental NumeroDocumento + TipoDocumento + Serie + ReferenciaLinea.
-- =============================================
-- Firma: FRANCO LARA - 11/07/2026 | Corrige el calculo multimoneda del analisis de cuentas para que los importes dolarizados del reporte siempre usen TotalImporteD por linea y no dependan de la moneda fija del plan de cuentas.

CREATE OR ALTER PROCEDURE dbo.usp_CON_ReporteAnalisisCuentas
    @IdEmpresa INT,
    @Periodo CHAR(6),
    @CuentaDesde VARCHAR(20) = NULL,
    @CuentaHasta VARCHAR(20) = NULL,
    @Auxiliar VARCHAR(20) = NULL,
    @Moneda VARCHAR(3) = 'PEN',
    @Estado CHAR(1) = 'T',
    @Tipo CHAR(1) = '0'
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @Ejercicio CHAR(4) = LEFT(@Periodo, 4);
        DECLARE @CuentaDesdeTrabajo VARCHAR(20);
        DECLARE @CuentaHastaTrabajo VARCHAR(20);
        DECLARE @AuxiliarTrabajo VARCHAR(20) = NULLIF(LTRIM(RTRIM(@Auxiliar)), '');
        DECLARE @MonedaTrabajo VARCHAR(3) = CASE WHEN UPPER(LTRIM(RTRIM(ISNULL(@Moneda, 'PEN')))) = 'USD' THEN 'USD' ELSE 'PEN' END;
        DECLARE @EstadoTrabajo CHAR(1) = CASE WHEN UPPER(LTRIM(RTRIM(ISNULL(@Estado, 'T')))) IN ('P', 'C') THEN UPPER(LTRIM(RTRIM(@Estado))) ELSE 'T' END;
        DECLARE @TipoTrabajo CHAR(1) = CASE WHEN @Tipo IN ('1', '2') THEN @Tipo ELSE '0' END;

        SELECT
            @CuentaDesdeTrabajo = COALESCE(NULLIF(LTRIM(RTRIM(@CuentaDesde)), ''), MIN(p.CodigoCuenta)),
            @CuentaHastaTrabajo = COALESCE(NULLIF(LTRIM(RTRIM(@CuentaHasta)), ''), MAX(p.CodigoCuenta))
        FROM dbo.CON_PlanCuenta AS p
        WHERE p.IdEmpresa = @IdEmpresa
          AND p.Estado = 1
          AND p.AceptaMovimiento = 1
          AND p.GeneraDiferenciaPorAnalisis = 1;

        IF @CuentaDesdeTrabajo IS NULL OR @CuentaHastaTrabajo IS NULL
        BEGIN
            SELECT
                CAST('' AS VARCHAR(20)) AS CodigoCuenta,
                CAST(N'' AS NVARCHAR(200)) AS NombreCuenta,
                CAST('' AS VARCHAR(20)) AS Auxiliar,
                CAST(N'' AS NVARCHAR(250)) AS NombreAuxiliar,
                CAST(N'' AS NVARCHAR(150)) AS TipoDocumento,
                CAST('' AS VARCHAR(10)) AS Serie,
                CAST(N'' AS NVARCHAR(100)) AS NumeroReferencia,
                CAST('' AS CHAR(6)) AS Periodo,
                CAST('' AS VARCHAR(30)) AS Comprobante,
                CAST(N'' AS NVARCHAR(300)) AS GlosaDetalle,
                CAST(NULL AS DATE) AS FechaEmision,
                CAST(0 AS DECIMAL(18, 6)) AS TipoCambio,
                CAST(0 AS DECIMAL(18, 2)) AS Debe,
                CAST(0 AS DECIMAL(18, 2)) AS Haber,
                CAST(0 AS DECIMAL(18, 2)) AS DebeDolares,
                CAST(0 AS DECIMAL(18, 2)) AS HaberDolares
            WHERE 1 = 0;

            RETURN;
        END;

        CREATE TABLE #BaseAnalisis
        (
            CodigoCuenta VARCHAR(20) NOT NULL,
            NombreCuenta NVARCHAR(200) NOT NULL,
            Auxiliar VARCHAR(20) NOT NULL,
            NombreAuxiliar NVARCHAR(250) NOT NULL,
            TipoDocumento NVARCHAR(150) NOT NULL,
            Serie VARCHAR(10) NOT NULL,
            NumeroReferencia NVARCHAR(100) NOT NULL,
            Periodo CHAR(6) NOT NULL,
            Comprobante VARCHAR(30) NOT NULL,
            GlosaDetalle NVARCHAR(300) NOT NULL,
            FechaEmision DATE NULL,
            TipoCambio DECIMAL(18, 6) NOT NULL,
            Debe DECIMAL(18, 2) NOT NULL,
            Haber DECIMAL(18, 2) NOT NULL,
            DebeDolares DECIMAL(18, 2) NOT NULL,
            HaberDolares DECIMAL(18, 2) NOT NULL
        );

        INSERT INTO #BaseAnalisis
        (
            CodigoCuenta,
            NombreCuenta,
            Auxiliar,
            NombreAuxiliar,
            TipoDocumento,
            Serie,
            NumeroReferencia,
            Periodo,
            Comprobante,
            GlosaDetalle,
            FechaEmision,
            TipoCambio,
            Debe,
            Haber,
            DebeDolares,
            HaberDolares
        )
        SELECT
            p.CodigoCuenta,
            p.NombreCuenta,
            ISNULL(NULLIF(LTRIM(RTRIM(d.NumeroDocumento)), ''), ''),
            ISNULL(
                NULLIF(
                    LTRIM(RTRIM(
                        COALESCE(per.NombreCompleto, per.RazonSocial, d.NumeroDocumento)
                    )),
                    ''
                ),
                ''
            ) AS NombreAuxiliar,
            ISNULL(NULLIF(LTRIM(RTRIM(d.TipoDocumento)), ''), ''),
            ISNULL(NULLIF(LTRIM(RTRIM(d.Serie)), ''), ''),
            ISNULL(NULLIF(LTRIM(RTRIM(d.ReferenciaLinea)), ''), ''),
            a.Periodo,
            CONCAT(ori.CodigoOrigen, '-', RIGHT('00000000' + CONVERT(VARCHAR(8), a.NumeroAsiento), 8)),
            ISNULL(d.GlosaDetalle, N''),
            a.FechaEmision,
            ISNULL(NULLIF(d.TipoCambioLinea, 0), a.TipoCambio),
            CASE
                WHEN d.DH = 'D' THEN CASE WHEN @MonedaTrabajo = 'USD' THEN d.TotalImporteD ELSE d.TotalImporteS END
                ELSE 0
            END,
            CASE
                WHEN d.DH = 'H' THEN CASE WHEN @MonedaTrabajo = 'USD' THEN d.TotalImporteD ELSE d.TotalImporteS END
                ELSE 0
            END,
            CASE
                WHEN d.DH = 'D' THEN d.TotalImporteD
                ELSE 0
            END,
            CASE
                WHEN d.DH = 'H' THEN d.TotalImporteD
                ELSE 0
            END
        FROM dbo.CON_AsientoDetalle AS d
        INNER JOIN dbo.CON_Asiento AS a
            ON a.IdAsiento = d.IdAsiento
        INNER JOIN dbo.CON_PlanCuenta AS p
            ON p.IdPlanCuenta = d.IdPlanCuenta
        INNER JOIN dbo.CON_Origen AS ori
            ON ori.IdOrigen = a.IdOrigen
        LEFT JOIN dbo.ADM_Persona AS per
            ON per.IdEmpresa = a.IdEmpresa
           AND per.NumeroDocumento = d.NumeroDocumento
        WHERE a.IdEmpresa = @IdEmpresa
          AND LEFT(a.Periodo, 4) = @Ejercicio
          AND a.Periodo <= @Periodo
          AND p.Estado = 1
          AND p.AceptaMovimiento = 1
          AND p.GeneraDiferenciaPorAnalisis = 1
          AND p.CodigoCuenta >= @CuentaDesdeTrabajo
          AND p.CodigoCuenta <= @CuentaHastaTrabajo
          AND (@AuxiliarTrabajo IS NULL OR d.NumeroDocumento = @AuxiliarTrabajo);

        IF @TipoTrabajo = '2'
        BEGIN
            CREATE TABLE #GruposAuxiliar
            (
                CodigoCuenta VARCHAR(20) NOT NULL,
                Auxiliar VARCHAR(20) NOT NULL,
                Saldo DECIMAL(18, 2) NOT NULL
            );

            INSERT INTO #GruposAuxiliar (CodigoCuenta, Auxiliar, Saldo)
            SELECT
                b.CodigoCuenta,
                b.Auxiliar,
                SUM(b.Debe - b.Haber)
            FROM #BaseAnalisis AS b
            GROUP BY
                b.CodigoCuenta,
                b.Auxiliar
            HAVING @EstadoTrabajo = 'T'
                OR (@EstadoTrabajo = 'P' AND SUM(b.Debe - b.Haber) <> 0)
                OR (@EstadoTrabajo = 'C' AND SUM(b.Debe - b.Haber) = 0);

            SELECT
                b.CodigoCuenta,
                MIN(b.NombreCuenta) AS NombreCuenta,
                b.Auxiliar,
                MIN(b.NombreAuxiliar) AS NombreAuxiliar,
                CAST(N'' AS NVARCHAR(150)) AS TipoDocumento,
                CAST('' AS VARCHAR(10)) AS Serie,
                CAST(N'' AS NVARCHAR(100)) AS NumeroReferencia,
                MIN(b.Periodo) AS Periodo,
                MIN(b.Comprobante) AS Comprobante,
                MIN(b.GlosaDetalle) AS GlosaDetalle,
                MIN(b.FechaEmision) AS FechaEmision,
                MIN(b.TipoCambio) AS TipoCambio,
                SUM(b.Debe) AS Debe,
                SUM(b.Haber) AS Haber,
                SUM(b.DebeDolares) AS DebeDolares,
                SUM(b.HaberDolares) AS HaberDolares
            FROM #BaseAnalisis AS b
            INNER JOIN #GruposAuxiliar AS g
                ON g.CodigoCuenta = b.CodigoCuenta
               AND g.Auxiliar = b.Auxiliar
            GROUP BY
                b.CodigoCuenta,
                b.Auxiliar
            ORDER BY
                b.CodigoCuenta,
                b.Auxiliar;

            RETURN;
        END;

        CREATE TABLE #GruposDocumento
        (
            CodigoCuenta VARCHAR(20) NOT NULL,
            Auxiliar VARCHAR(20) NOT NULL,
            TipoDocumento NVARCHAR(150) NOT NULL,
            Serie VARCHAR(10) NOT NULL,
            NumeroReferencia NVARCHAR(100) NOT NULL,
            Saldo DECIMAL(18, 2) NOT NULL
        );

        INSERT INTO #GruposDocumento
        (
            CodigoCuenta,
            Auxiliar,
            TipoDocumento,
            Serie,
            NumeroReferencia,
            Saldo
        )
        SELECT
            b.CodigoCuenta,
            b.Auxiliar,
            b.TipoDocumento,
            b.Serie,
            b.NumeroReferencia,
            SUM(b.Debe - b.Haber)
        FROM #BaseAnalisis AS b
        GROUP BY
            b.CodigoCuenta,
            b.Auxiliar,
            b.TipoDocumento,
            b.Serie,
            b.NumeroReferencia
        HAVING @EstadoTrabajo = 'T'
            OR (@EstadoTrabajo = 'P' AND SUM(b.Debe - b.Haber) <> 0)
            OR (@EstadoTrabajo = 'C' AND SUM(b.Debe - b.Haber) = 0);

        IF @TipoTrabajo = '1'
        BEGIN
            SELECT
                b.CodigoCuenta,
                MIN(b.NombreCuenta) AS NombreCuenta,
                b.Auxiliar,
                MIN(b.NombreAuxiliar) AS NombreAuxiliar,
                b.TipoDocumento,
                b.Serie,
                b.NumeroReferencia,
                MIN(b.Periodo) AS Periodo,
                MIN(b.Comprobante) AS Comprobante,
                MIN(b.GlosaDetalle) AS GlosaDetalle,
                MIN(b.FechaEmision) AS FechaEmision,
                MIN(b.TipoCambio) AS TipoCambio,
                SUM(b.Debe) AS Debe,
                SUM(b.Haber) AS Haber,
                SUM(b.DebeDolares) AS DebeDolares,
                SUM(b.HaberDolares) AS HaberDolares
            FROM #BaseAnalisis AS b
            INNER JOIN #GruposDocumento AS g
                ON g.CodigoCuenta = b.CodigoCuenta
               AND g.Auxiliar = b.Auxiliar
               AND g.TipoDocumento = b.TipoDocumento
               AND g.Serie = b.Serie
               AND g.NumeroReferencia = b.NumeroReferencia
            GROUP BY
                b.CodigoCuenta,
                b.Auxiliar,
                b.TipoDocumento,
                b.Serie,
                b.NumeroReferencia
            ORDER BY
                b.CodigoCuenta,
                b.Auxiliar,
                b.TipoDocumento,
                b.Serie,
                b.NumeroReferencia;

            RETURN;
        END;

        SELECT
            b.CodigoCuenta,
            b.NombreCuenta,
            b.Auxiliar,
            b.NombreAuxiliar,
            b.TipoDocumento,
            b.Serie,
            b.NumeroReferencia,
            b.Periodo,
            b.Comprobante,
            b.GlosaDetalle,
            b.FechaEmision,
            b.TipoCambio,
            b.Debe,
            b.Haber,
            b.DebeDolares,
            b.HaberDolares
        FROM #BaseAnalisis AS b
        INNER JOIN #GruposDocumento AS g
            ON g.CodigoCuenta = b.CodigoCuenta
           AND g.Auxiliar = b.Auxiliar
           AND g.TipoDocumento = b.TipoDocumento
           AND g.Serie = b.Serie
           AND g.NumeroReferencia = b.NumeroReferencia
        ORDER BY
            b.CodigoCuenta,
            b.Auxiliar,
            b.TipoDocumento,
            b.Serie,
            b.NumeroReferencia,
            b.Periodo,
            b.FechaEmision,
            b.Comprobante;

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
