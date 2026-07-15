-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   07/07/2026
-- Description:   Lista la estructura base del Libro Mayor PLE formato 6.1 ordenada por cuenta, fecha, CUO y correlativo.
-- =============================================
-- Firma: FRANCO LARA - 13/07/2026 | Fija la exportacion PLE 6.1 a moneda PEN, elimina la bifurcacion por USD y usa siempre TotalImporteS como importe base de Debe y Haber.
-- Firma: FRANCO LARA - 14/07/2026 | Incluye el asiento de apertura periodo 00 al exportar enero, agrega en diciembre los periodos 12, 13, 14 y 15 y marca el correlativo de movimiento como A/M/C, dejando C solo para los asientos de cierre de los periodos 14 y 15.

CREATE OR ALTER PROCEDURE dbo.usp_CON_PLE_LibroMayor61_Listar
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
            CASE
                WHEN a.Periodo = @PeriodoApertura THEN 'A'
                WHEN a.Periodo IN (@PeriodoCierreResultados, @PeriodoCierreInventarios) THEN 'C'
                ELSE 'M'
            END + RIGHT(REPLICATE('0', 4) + CONVERT(VARCHAR(10), d.Item), 4) AS CorrelativoMovimiento,
            p.CodigoCuenta AS CodigoCuentaContable,
            a.FechaAsiento AS FechaOperacion,
            REPLACE(REPLACE(COALESCE(NULLIF(LTRIM(RTRIM(d.GlosaDetalle)), ''), a.Glosa, N''), CHAR(13), ' '), CHAR(10), ' ') AS Glosa,
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
            p.CodigoCuenta,
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
