-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   07/07/2026
-- Description:   Lista la estructura simplificada del Libro Diario PLE formato 5.2 desde los mismos movimientos contables del diario.
-- =============================================
-- Firma: FRANCO LARA - 13/07/2026 | Fija la exportacion PLE 5.2 a moneda PEN, elimina la bifurcacion por USD y usa siempre TotalImporteS como importe base de Debe y Haber.
-- Firma: FRANCO LARA - 14/07/2026 | Incluye el periodo 00 al exportar enero y agrega en diciembre los periodos 12, 13, 14 y 15 para que el libro simplificado considere apertura y cierre anual, manteniendo la nocion de cierre anual solo en 14 y 15.
-- Firma: FRANCO LARA - 03/08/2026 | Completa los 21 campos base del PLE 5.2 usando la misma fuente documentaria validada por el Libro Diario 5.1 y genera correlativos A/M/C por movimiento.

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
            CASE
                WHEN a.Periodo = @PeriodoApertura THEN 'A'
                WHEN a.Periodo IN (@PeriodoCierreResultados, @PeriodoCierreInventarios) THEN 'C'
                ELSE 'M'
            END + RIGHT(REPLICATE('0', 4) + CONVERT(VARCHAR(10), d.Item), 4) AS CorrelativoAsiento,
            p.CodigoCuenta AS CodigoCuentaContable,
            '' AS CodigoUnidadOperacion,
            ISNULL(NULLIF(LTRIM(RTRIM(d.CodigoCentroCosto)), ''), '') AS CodigoCentroCosto,
            m.CodigoMoneda AS CodigoMoneda,
            '050200' AS CodigoLibroRelacionado,
            ISNULL(NULLIF(LTRIM(RTRIM(COALESCE(perAsiento.TipoDocumento, perCompra.TipoDocumento, perVenta.TipoDocumento))), ''), '') AS TipoDocumentoEmisor,
            ISNULL(NULLIF(LTRIM(RTRIM(COALESCE(perAsiento.NumeroDocumento, perCompra.NumeroDocumento, perVenta.NumeroDocumento, d.NumeroDocumento))), ''), '') AS NumeroDocumentoEmisor,
            ISNULL(NULLIF(LTRIM(RTRIM(COALESCE(c.TipoComprobante, v.TipoComprobante, d.TipoDocumento))), ''), '') AS TipoComprobante,
            ISNULL(NULLIF(LTRIM(RTRIM(COALESCE(c.Serie, v.Serie, d.Serie))), ''), '') AS SerieComprobante,
            ISNULL(NULLIF(LTRIM(RTRIM(COALESCE(c.Numero, v.Numero, d.ReferenciaLinea, d.NumeroDocumento))), ''), '') AS NumeroComprobante,
            CASE
                WHEN doc.FechaOperacionBase IS NOT NULL
                 AND CONVERT(CHAR(6), doc.FechaOperacionBase, 112) = a.Periodo
                    THEN doc.FechaOperacionBase
                ELSE a.FechaAsiento
            END AS FechaContable,
            CAST(NULL AS DATE) AS FechaVencimiento,
            doc.FechaOperacionBase AS FechaOperacion,
            REPLACE(REPLACE(COALESCE(NULLIF(LTRIM(RTRIM(d.GlosaDetalle)), ''), a.Glosa, N''), CHAR(13), ' '), CHAR(10), ' ') AS Glosa,
            ISNULL(NULLIF(LTRIM(RTRIM(a.ReferenciaExterna)), ''), '') AS GlosaReferencial,
            CASE WHEN d.DH = 'D' THEN d.TotalImporteS ELSE 0 END AS Debe,
            CASE WHEN d.DH = 'H' THEN d.TotalImporteS ELSE 0 END AS Haber,
            ISNULL(NULLIF(LTRIM(RTRIM(d.ReferenciaLinea)), ''), '') AS InformacionComplementaria,
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
        LEFT JOIN dbo.ADM_Persona AS perAsiento
            ON perAsiento.IdEmpresa = a.IdEmpresa
           AND perAsiento.NumeroDocumento = d.NumeroDocumento
        LEFT JOIN dbo.COM_Compra AS c
            ON c.IdAsiento = a.IdAsiento
           AND c.IdEmpresa = a.IdEmpresa
        LEFT JOIN dbo.ADM_Proveedor AS pr
            ON pr.IdProveedor = c.IdProveedor
        LEFT JOIN dbo.ADM_Persona AS perCompra
            ON perCompra.IdPersona = pr.IdPersona
        LEFT JOIN dbo.VEN_Venta AS v
            ON v.IdAsiento = a.IdAsiento
           AND v.IdEmpresa = a.IdEmpresa
        LEFT JOIN dbo.ADM_Cliente AS cl
            ON cl.IdCliente = v.IdCliente
        LEFT JOIN dbo.ADM_Persona AS perVenta
            ON perVenta.IdPersona = cl.IdPersona
        CROSS APPLY
        (
            SELECT COALESCE(c.FechaEmision, v.FechaEmision, a.FechaEmision) AS FechaOperacionBase
        ) AS doc
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
