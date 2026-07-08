-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   07/07/2026
-- Description:   Lista la estructura simplificada del Libro Diario PLE formato 5.2 desde los mismos movimientos contables del diario.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_PLE_LibroDiario52_Listar
    @IdEmpresa INT,
    @IdAnno SMALLINT,
    @Mes TINYINT,
    @Moneda VARCHAR(3) = 'PEN',
    @Estado VARCHAR(10) = NULL,
    @FechaDesde DATE = NULL,
    @FechaHasta DATE = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @Periodo CHAR(6) = CONVERT(CHAR(4), @IdAnno) + RIGHT('0' + CONVERT(VARCHAR(2), @Mes), 2);
        DECLARE @PeriodoPle CHAR(8) = @Periodo + '00';
        DECLARE @MonedaTrabajo VARCHAR(3) = CASE WHEN UPPER(LTRIM(RTRIM(ISNULL(@Moneda, 'PEN')))) = 'USD' THEN 'USD' ELSE 'PEN' END;
        DECLARE @EstadoTrabajo VARCHAR(10) = NULLIF(LTRIM(RTRIM(@Estado)), '');

        SELECT
            @PeriodoPle AS PeriodoPle,
            RIGHT(REPLICATE('0', 8) + CONVERT(VARCHAR(20), a.IdAsiento), 8) AS Cuo,
            RIGHT(REPLICATE('0', 5) + CONVERT(VARCHAR(20), a.NumeroAsiento), 5) AS CorrelativoAsiento,
            a.FechaAsiento AS FechaOperacion,
            REPLACE(REPLACE(COALESCE(NULLIF(LTRIM(RTRIM(d.GlosaDetalle)), ''), a.Glosa, N''), CHAR(13), ' '), CHAR(10), ' ') AS Glosa,
            p.CodigoCuenta AS CodigoCuentaContable,
            m.CodigoMoneda AS CodigoMoneda,
            CASE WHEN d.DH = 'D' THEN CASE WHEN @MonedaTrabajo = 'USD' THEN d.TotalImporteD ELSE d.TotalImporteS END ELSE 0 END AS Debe,
            CASE WHEN d.DH = 'H' THEN CASE WHEN @MonedaTrabajo = 'USD' THEN d.TotalImporteD ELSE d.TotalImporteS END ELSE 0 END AS Haber,
            CASE
                WHEN @EstadoTrabajo IS NULL OR @EstadoTrabajo = 'Todos' THEN '1'
                WHEN @EstadoTrabajo IN ('1', '6', '8', '9') THEN @EstadoTrabajo
                ELSE '1'
            END AS EstadoOperacion,
            a.NumeroAsiento
        FROM dbo.CON_Asiento AS a
        INNER JOIN dbo.CON_AsientoDetalle AS d
            ON d.IdAsiento = a.IdAsiento
        INNER JOIN dbo.CON_PlanCuenta AS p
            ON p.IdPlanCuenta = d.IdPlanCuenta
        INNER JOIN dbo.ADM_Moneda AS m
            ON m.IdMoneda = a.IdMoneda
        WHERE a.IdEmpresa = @IdEmpresa
          AND a.Periodo = @Periodo
          AND (@FechaDesde IS NULL OR a.FechaAsiento >= @FechaDesde)
          AND (@FechaHasta IS NULL OR a.FechaAsiento <= @FechaHasta)
          AND (
                @EstadoTrabajo IS NULL
                OR @EstadoTrabajo = 'Todos'
                OR @EstadoTrabajo IN ('1', '6', '8', '9')
              )
        ORDER BY
            a.FechaAsiento,
            a.IdAsiento,
            d.Item;

    END TRY

    BEGIN CATCH

        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE()

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)

    END CATCH

END
