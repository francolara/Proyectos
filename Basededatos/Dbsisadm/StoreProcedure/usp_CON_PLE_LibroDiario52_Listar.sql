-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   07/07/2026
-- Description:   Lista la estructura simplificada del Libro Diario PLE formato 5.2 desde los mismos movimientos contables del diario.
-- =============================================
-- Firma: FRANCO LARA - 13/07/2026 | Fija la exportacion PLE 5.2 a moneda PEN, elimina la bifurcacion por USD y usa siempre TotalImporteS como importe base de Debe y Haber.
-- Firma: FRANCO LARA - 14/07/2026 | Incluye el periodo 00 al exportar enero y agrega en diciembre los periodos 12, 13, 14 y 15 para que el libro simplificado considere apertura y cierre anual, manteniendo la nocion de cierre anual solo en 14 y 15.

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
        DECLARE @EstadoTrabajo VARCHAR(10) = NULLIF(LTRIM(RTRIM(@Estado)), '');
        DECLARE @PeriodoApertura CHAR(6) = CONVERT(CHAR(4), @IdAnno) + '00';
        DECLARE @PeriodoAjusteFinal CHAR(6) = CONVERT(CHAR(4), @IdAnno) + '13';
        DECLARE @PeriodoCierreResultados CHAR(6) = CONVERT(CHAR(4), @IdAnno) + '14';
        DECLARE @PeriodoCierreInventarios CHAR(6) = CONVERT(CHAR(4), @IdAnno) + '15';

        SELECT
            @PeriodoPle AS PeriodoPle,
            RIGHT(REPLICATE('0', 8) + CONVERT(VARCHAR(20), a.IdAsiento), 8) AS Cuo,
            RIGHT(REPLICATE('0', 5) + CONVERT(VARCHAR(20), a.NumeroAsiento), 5) AS CorrelativoAsiento,
            a.FechaAsiento AS FechaOperacion,
            REPLACE(REPLACE(COALESCE(NULLIF(LTRIM(RTRIM(d.GlosaDetalle)), ''), a.Glosa, N''), CHAR(13), ' '), CHAR(10), ' ') AS Glosa,
            p.CodigoCuenta AS CodigoCuentaContable,
            'PEN' AS CodigoMoneda,
            CASE WHEN d.DH = 'D' THEN d.TotalImporteS ELSE 0 END AS Debe,
            CASE WHEN d.DH = 'H' THEN d.TotalImporteS ELSE 0 END AS Haber,
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
        WHERE a.IdEmpresa = @IdEmpresa
          AND (
                a.Periodo = @Periodo
                OR (@Mes = 1 AND a.Periodo = @PeriodoApertura)
                OR (@Mes = 12 AND a.Periodo IN (@PeriodoAjusteFinal, @PeriodoCierreResultados, @PeriodoCierreInventarios))
              )
          AND (@FechaDesde IS NULL OR a.FechaAsiento >= @FechaDesde)
          AND (@FechaHasta IS NULL OR a.FechaAsiento <= @FechaHasta)
          AND (
                @EstadoTrabajo IS NULL
                OR @EstadoTrabajo = 'Todos'
                OR @EstadoTrabajo IN ('1', '6', '8', '9')
              )
        ORDER BY
            a.Periodo,
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
