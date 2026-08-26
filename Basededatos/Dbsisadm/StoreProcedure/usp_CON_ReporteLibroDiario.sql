-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   06/07/2026
-- Description:   Replica el libro diario legacy en HTML, cubriendo diario auxiliar y diario por origen detallado/resumido.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   08/07/2026
-- Description:   Ajusta Libro Diario para fijar moneda base en soles, eliminar filtro por origen y separar las vistas totalizadas Por Cuenta y Por Origen.
-- =============================================
-- Firma: FRANCO LARA - 26/08/2026 | Agrega filtros opcionales CuentaDesde y CuentaHasta sobre el codigo contable antes de totalizar el Libro Diario.

CREATE OR ALTER PROCEDURE dbo.usp_CON_ReporteLibroDiario
    @IdEmpresa INT,
    @Periodo CHAR(6),
    @Moneda VARCHAR(3) = 'PEN',
    @Modo CHAR(1) = 'A',
    @OrigenDesde VARCHAR(10) = NULL,
    @OrigenHasta VARCHAR(10) = NULL,
    @CuentaDesde VARCHAR(20) = NULL,
    @CuentaHasta VARCHAR(20) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @MonedaTrabajo VARCHAR(3) = CASE WHEN UPPER(LTRIM(RTRIM(ISNULL(@Moneda, 'PEN')))) = 'USD' THEN 'USD' ELSE 'PEN' END;
        DECLARE @ModoTrabajo CHAR(1) = CASE WHEN UPPER(ISNULL(@Modo, 'A')) IN ('D', 'R') THEN UPPER(@Modo) ELSE 'A' END;
        DECLARE @OrigenDesdeTrabajo VARCHAR(10);
        DECLARE @OrigenHastaTrabajo VARCHAR(10);
        DECLARE @CuentaDesdeTrabajo VARCHAR(20) = NULLIF(LTRIM(RTRIM(@CuentaDesde)), '');
        DECLARE @CuentaHastaTrabajo VARCHAR(20) = NULLIF(LTRIM(RTRIM(@CuentaHasta)), '');

        SELECT
            @OrigenDesdeTrabajo = COALESCE(NULLIF(LTRIM(RTRIM(@OrigenDesde)), ''), MIN(o.CodigoOrigen)),
            @OrigenHastaTrabajo = COALESCE(NULLIF(LTRIM(RTRIM(@OrigenHasta)), ''), MAX(o.CodigoOrigen))
        FROM dbo.CON_Origen AS o
        WHERE o.IdEmpresa = @IdEmpresa
          AND o.Estado = 1;

        CREATE TABLE #BaseDiario
        (
            Modo CHAR(1) NOT NULL,
            CodigoOrigen VARCHAR(10) NOT NULL,
            NombreOrigen NVARCHAR(150) NOT NULL,
            Periodo CHAR(6) NOT NULL,
            NumeroAsiento INT NOT NULL,
            Item SMALLINT NOT NULL,
            FechaEmision DATE NOT NULL,
            CodigoCuenta VARCHAR(20) NOT NULL,
            NombreCuenta NVARCHAR(200) NOT NULL,
            NumeroDocumento VARCHAR(20) NOT NULL,
            NombreAuxiliar NVARCHAR(250) NOT NULL,
            TipoDocumento NVARCHAR(150) NOT NULL,
            Serie VARCHAR(10) NOT NULL,
            Referencia NVARCHAR(100) NOT NULL,
            Glosa NVARCHAR(500) NOT NULL,
            TipoCambio DECIMAL(18, 6) NOT NULL,
            Debe DECIMAL(18, 2) NOT NULL,
            Haber DECIMAL(18, 2) NOT NULL,
            DebeDolares DECIMAL(18, 2) NOT NULL,
            HaberDolares DECIMAL(18, 2) NOT NULL
        );

        INSERT INTO #BaseDiario
        (
            Modo,
            CodigoOrigen,
            NombreOrigen,
            Periodo,
            NumeroAsiento,
            Item,
            FechaEmision,
            CodigoCuenta,
            NombreCuenta,
            NumeroDocumento,
            NombreAuxiliar,
            TipoDocumento,
            Serie,
            Referencia,
            Glosa,
            TipoCambio,
            Debe,
            Haber,
            DebeDolares,
            HaberDolares
        )
        SELECT
            @ModoTrabajo,
            o.CodigoOrigen,
            o.NombreOrigen,
            a.Periodo,
            a.NumeroAsiento,
            d.Item,
            a.FechaEmision,
            p.CodigoCuenta,
            p.NombreCuenta,
            ISNULL(NULLIF(LTRIM(RTRIM(d.NumeroDocumento)), ''), ''),
            ISNULL(NULLIF(LTRIM(RTRIM(COALESCE(per.NombreCompleto, per.RazonSocial, d.NumeroDocumento))), ''), ''),
            ISNULL(NULLIF(LTRIM(RTRIM(d.TipoDocumento)), ''), ''),
            ISNULL(NULLIF(LTRIM(RTRIM(d.Serie)), ''), ''),
            ISNULL(NULLIF(LTRIM(RTRIM(d.ReferenciaLinea)), ''), ''),
            COALESCE(NULLIF(LTRIM(RTRIM(d.GlosaDetalle)), ''), a.Glosa, N''),
            ISNULL(NULLIF(d.TipoCambioLinea, 0), a.TipoCambio),
            CASE WHEN d.DH = 'D' THEN CASE WHEN @MonedaTrabajo = 'USD' THEN d.TotalImporteD ELSE d.TotalImporteS END ELSE 0 END,
            CASE WHEN d.DH = 'H' THEN CASE WHEN @MonedaTrabajo = 'USD' THEN d.TotalImporteD ELSE d.TotalImporteS END ELSE 0 END,
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
          AND (@CuentaDesdeTrabajo IS NULL OR p.CodigoCuenta >= @CuentaDesdeTrabajo)
          AND (@CuentaHastaTrabajo IS NULL OR p.CodigoCuenta <= @CuentaHastaTrabajo)
          AND (
                @ModoTrabajo = 'A'
                OR (
                    o.CodigoOrigen >= ISNULL(@OrigenDesdeTrabajo, o.CodigoOrigen)
                    AND o.CodigoOrigen <= ISNULL(@OrigenHastaTrabajo, o.CodigoOrigen)
                )
              );

        IF @ModoTrabajo = 'D'
        BEGIN
            SELECT
                b.Modo,
                CAST('' AS VARCHAR(10)) AS CodigoOrigen,
                CAST(N'' AS NVARCHAR(150)) AS NombreOrigen,
                MIN(b.Periodo) AS Periodo,
                0 AS NumeroAsiento,
                CAST(0 AS SMALLINT) AS Item,
                MIN(b.FechaEmision) AS FechaEmision,
                b.CodigoCuenta,
                MIN(b.NombreCuenta) AS NombreCuenta,
                CAST('' AS VARCHAR(20)) AS NumeroDocumento,
                CAST(N'' AS NVARCHAR(250)) AS NombreAuxiliar,
                CAST(N'' AS NVARCHAR(150)) AS TipoDocumento,
                CAST('' AS VARCHAR(10)) AS Serie,
                CAST(N'' AS NVARCHAR(100)) AS Referencia,
                CAST(N'TOTAL POR CUENTA' AS NVARCHAR(500)) AS Glosa,
                CAST(0 AS DECIMAL(18, 6)) AS TipoCambio,
                SUM(b.Debe) AS Debe,
                SUM(b.Haber) AS Haber,
                SUM(b.DebeDolares) AS DebeDolares,
                SUM(b.HaberDolares) AS HaberDolares
            FROM #BaseDiario AS b
            GROUP BY
                b.Modo,
                b.CodigoCuenta
            ORDER BY
                b.CodigoCuenta;

            RETURN;
        END;

        IF @ModoTrabajo = 'R'
        BEGIN
            SELECT
                b.Modo,
                b.CodigoOrigen,
                MIN(b.NombreOrigen) AS NombreOrigen,
                MIN(b.Periodo) AS Periodo,
                0 AS NumeroAsiento,
                CAST(0 AS SMALLINT) AS Item,
                MIN(b.FechaEmision) AS FechaEmision,
                CAST('' AS VARCHAR(20)) AS CodigoCuenta,
                CAST(N'' AS NVARCHAR(200)) AS NombreCuenta,
                CAST('' AS VARCHAR(20)) AS NumeroDocumento,
                CAST(N'' AS NVARCHAR(250)) AS NombreAuxiliar,
                CAST(N'' AS NVARCHAR(150)) AS TipoDocumento,
                CAST('' AS VARCHAR(10)) AS Serie,
                CAST(N'' AS NVARCHAR(100)) AS Referencia,
                CAST(N'TOTAL POR ORIGEN' AS NVARCHAR(500)) AS Glosa,
                CAST(0 AS DECIMAL(18, 6)) AS TipoCambio,
                SUM(b.Debe) AS Debe,
                SUM(b.Haber) AS Haber,
                SUM(b.DebeDolares) AS DebeDolares,
                SUM(b.HaberDolares) AS HaberDolares
            FROM #BaseDiario AS b
            GROUP BY
                b.Modo,
                b.CodigoOrigen
            ORDER BY
                b.CodigoOrigen;

            RETURN;
        END;

        SELECT
            b.Modo,
            b.CodigoOrigen,
            b.NombreOrigen,
            b.Periodo,
            b.NumeroAsiento,
            b.Item,
            b.FechaEmision,
            b.CodigoCuenta,
            b.NombreCuenta,
            b.NumeroDocumento,
            b.NombreAuxiliar,
            b.TipoDocumento,
            b.Serie,
            b.Referencia,
            b.Glosa,
            b.TipoCambio,
            b.Debe,
            b.Haber,
            b.DebeDolares,
            b.HaberDolares
        FROM #BaseDiario AS b
        ORDER BY
            b.CodigoOrigen,
            b.FechaEmision,
            b.NumeroAsiento,
            b.Item,
            b.CodigoCuenta;

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
