-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   07/07/2026
-- Description:   Lista la estructura base del Libro Diario PLE formato 5.1 desde asientos contables, compras y ventas.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_PLE_LibroDiario51_Listar
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
            'M' + RIGHT(REPLICATE('0', 4) + CONVERT(VARCHAR(10), d.Item), 4) AS CorrelativoMovimiento,
            p.CodigoCuenta AS CodigoCuentaContable,
            ISNULL(NULLIF(LTRIM(RTRIM(d.CodigoCentroCosto)), ''), '') AS CodigoUnidadOperacion,
            m.CodigoMoneda AS CodigoMoneda,
            ISNULL(NULLIF(LTRIM(RTRIM(COALESCE(perAsiento.TipoDocumento, perCompra.TipoDocumento, perVenta.TipoDocumento))), ''), '') AS TipoDocumentoEmisor,
            ISNULL(NULLIF(LTRIM(RTRIM(COALESCE(perAsiento.NumeroDocumento, perCompra.NumeroDocumento, perVenta.NumeroDocumento, d.NumeroDocumento))), ''), '') AS NumeroDocumentoEmisor,
            ISNULL(NULLIF(LTRIM(RTRIM(COALESCE(c.TipoComprobante, v.TipoComprobante))), ''), '') AS TipoComprobante,
            ISNULL(NULLIF(LTRIM(RTRIM(COALESCE(c.Serie, v.Serie, d.Serie))), ''), '') AS SerieComprobante,
            ISNULL(NULLIF(LTRIM(RTRIM(COALESCE(c.Numero, v.Numero, d.NumeroDocumento))), ''), '') AS NumeroComprobante,
            a.FechaAsiento AS FechaContable,
            CAST(NULL AS DATE) AS FechaVencimiento,
            COALESCE(c.FechaEmision, v.FechaEmision, a.FechaEmision) AS FechaOperacion,
            REPLACE(REPLACE(COALESCE(NULLIF(LTRIM(RTRIM(d.GlosaDetalle)), ''), a.Glosa, N''), CHAR(13), ' '), CHAR(10), ' ') AS Glosa,
            ISNULL(NULLIF(LTRIM(RTRIM(a.ReferenciaExterna)), ''), '') AS GlosaReferencial,
            CASE WHEN d.DH = 'D' THEN CASE WHEN @MonedaTrabajo = 'USD' THEN d.TotalImporteD ELSE d.TotalImporteS END ELSE 0 END AS Debe,
            CASE WHEN d.DH = 'H' THEN CASE WHEN @MonedaTrabajo = 'USD' THEN d.TotalImporteD ELSE d.TotalImporteS END ELSE 0 END AS Haber,
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
