-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   07/07/2026
-- Description:   Lista la estructura base del Libro Mayor PLE formato 6.1 ordenada por cuenta, fecha, CUO y correlativo.
-- =============================================
-- Firma: FRANCO LARA - 13/07/2026 | Fija la exportacion PLE 6.1 a moneda PEN, elimina la bifurcacion por USD y usa siempre TotalImporteS como importe base de Debe y Haber.
-- Firma: FRANCO LARA - 14/07/2026 | Incluye el asiento de apertura periodo 00 al exportar enero, agrega en diciembre los periodos 12, 13, 14 y 15 y marca el correlativo de movimiento como A/M/C, dejando C solo para los asientos de cierre de los periodos 14 y 15.
-- Firma: FRANCO LARA - 03/08/2026 | Completa los 21 campos base del PLE 6.1 usando la misma fuente documentaria validada por el Libro Diario 5.1.
-- Firma: FRANCO LARA - 04/08/2026 | Genera el CUO con origen, periodo y numero de asiento; referencia Compras o Ventas, incluyendo bancos, detracciones y percepciones asociadas.

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
            CONCAT(
                LTRIM(RTRIM(o.CodigoOrigen)),
                a.Periodo,
                CASE
                    WHEN LEN(CONVERT(VARCHAR(20), a.NumeroAsiento)) < 8
                        THEN RIGHT(REPLICATE('0', 8) + CONVERT(VARCHAR(20), a.NumeroAsiento), 8)
                    ELSE CONVERT(VARCHAR(20), a.NumeroAsiento)
                END
            ) AS Cuo,
            CASE
                WHEN a.Periodo = @PeriodoApertura THEN 'A'
                WHEN a.Periodo IN (@PeriodoCierreResultados, @PeriodoCierreInventarios) THEN 'C'
                ELSE 'M'
            END + RIGHT(REPLICATE('0', 4) + CONVERT(VARCHAR(10), d.Item), 4) AS CorrelativoMovimiento,
            p.CodigoCuenta AS CodigoCuentaContable,
            '' AS CodigoUnidadOperacion,
            ISNULL(NULLIF(LTRIM(RTRIM(d.CodigoCentroCosto)), ''), '') AS CodigoCentroCosto,
            m.CodigoMoneda AS CodigoMoneda,
            origen.CodigoLibroRelacionado,
            origen.PeriodoReferencia,
            origen.CuoReferencia,
            origen.CorrelativoReferencia,
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
            '' AS InformacionComplementaria,
            CASE
                WHEN @EstadoTrabajo IS NULL OR @EstadoTrabajo = 'Todos' THEN '1'
                WHEN @EstadoTrabajo IN ('1', '6', '8', '9') THEN @EstadoTrabajo
                ELSE '1'
            END AS EstadoOperacion,
            a.NumeroAsiento
        FROM dbo.CON_Asiento AS a
        INNER JOIN dbo.CON_AsientoDetalle AS d
            ON d.IdAsiento = a.IdAsiento
        INNER JOIN dbo.CON_Origen AS o
            ON o.IdOrigen = a.IdOrigen
           AND o.IdEmpresa = a.IdEmpresa
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
        LEFT JOIN dbo.COM_CompraDetraccion AS detraccionAsiento
            ON detraccionAsiento.IdAsiento = a.IdAsiento
           AND detraccionAsiento.IdEmpresa = a.IdEmpresa
        LEFT JOIN dbo.COM_CompraPercepcion AS percepcionAsiento
            ON percepcionAsiento.IdAsiento = a.IdAsiento
           AND percepcionAsiento.IdEmpresa = a.IdEmpresa
        LEFT JOIN dbo.COM_Compra AS compraProceso
            ON compraProceso.IdCompra = COALESCE(detraccionAsiento.IdCompra, percepcionAsiento.IdCompra)
           AND compraProceso.IdEmpresa = a.IdEmpresa
        LEFT JOIN dbo.BAN_MovimientoBanco AS banco
            ON banco.IdAsiento = a.IdAsiento
           AND banco.IdEmpresa = a.IdEmpresa
           AND banco.Activo = 1
        LEFT JOIN dbo.BAN_MovimientoBancoDetalle AS bancoDetalle
            ON bancoDetalle.IdMovimientoBanco = banco.IdMovimientoBanco
           AND bancoDetalle.Item = d.Item - 1
           AND bancoDetalle.ModuloOperacionComprobante IN ('COM', 'VEN', 'DET', 'PER', 'R4T')
           AND bancoDetalle.IdRegistroComprobante IS NOT NULL
        LEFT JOIN dbo.COM_CompraDetraccion AS detraccionBanco
            ON bancoDetalle.ModuloOperacionComprobante = 'DET'
           AND detraccionBanco.IdCompraDetraccion = bancoDetalle.IdRegistroComprobante
           AND detraccionBanco.IdEmpresa = a.IdEmpresa
        LEFT JOIN dbo.COM_CompraPercepcion AS percepcionBanco
            ON bancoDetalle.ModuloOperacionComprobante = 'PER'
           AND percepcionBanco.IdCompraPercepcion = bancoDetalle.IdRegistroComprobante
           AND percepcionBanco.IdEmpresa = a.IdEmpresa
        LEFT JOIN dbo.COM_CompraRetencion AS retencionBanco
            ON bancoDetalle.ModuloOperacionComprobante = 'R4T'
           AND retencionBanco.IdCompraRetencion = bancoDetalle.IdRegistroComprobante
           AND retencionBanco.IdEmpresa = a.IdEmpresa
        LEFT JOIN dbo.COM_Compra AS compraBanco
            ON compraBanco.IdCompra = CASE bancoDetalle.ModuloOperacionComprobante
                                          WHEN 'COM' THEN bancoDetalle.IdRegistroComprobante
                                          WHEN 'DET' THEN detraccionBanco.IdCompra
                                          WHEN 'PER' THEN percepcionBanco.IdCompra
                                          WHEN 'R4T' THEN retencionBanco.IdCompra
                                          ELSE NULL
                                      END
           AND compraBanco.IdEmpresa = a.IdEmpresa
        LEFT JOIN dbo.VEN_Venta AS ventaBanco
            ON bancoDetalle.ModuloOperacionComprobante = 'VEN'
           AND ventaBanco.IdVenta = bancoDetalle.IdRegistroComprobante
           AND ventaBanco.IdEmpresa = a.IdEmpresa
        LEFT JOIN dbo.CON_Asiento AS asientoReferencia
            ON asientoReferencia.IdAsiento = COALESCE(c.IdAsiento, compraProceso.IdAsiento, compraBanco.IdAsiento, v.IdAsiento, ventaBanco.IdAsiento)
           AND asientoReferencia.IdEmpresa = a.IdEmpresa
        LEFT JOIN dbo.CON_Origen AS origenReferencia
            ON origenReferencia.IdOrigen = asientoReferencia.IdOrigen
           AND origenReferencia.IdEmpresa = asientoReferencia.IdEmpresa
        OUTER APPLY
        (
            SELECT
                CASE
                    WHEN COALESCE(c.IdAsiento, compraProceso.IdAsiento, compraBanco.IdAsiento) IS NOT NULL
                     AND COALESCE(c.TipoComprobante, compraProceso.TipoComprobante, compraBanco.TipoComprobante) IN ('91', '97', '98') THEN '080200'
                    WHEN COALESCE(c.IdAsiento, compraProceso.IdAsiento, compraBanco.IdAsiento) IS NOT NULL THEN '080100'
                    WHEN COALESCE(v.IdAsiento, ventaBanco.IdAsiento) IS NOT NULL THEN '140100'
                    ELSE ''
                END AS CodigoLibroRelacionado,
                CASE
                    WHEN COALESCE(c.IdAsiento, compraProceso.IdAsiento, compraBanco.IdAsiento) IS NOT NULL
                        THEN CONVERT(CHAR(6), COALESCE(c.FechaContabilizacion, compraProceso.FechaContabilizacion, compraBanco.FechaContabilizacion), 112) + '00'
                    WHEN COALESCE(v.IdAsiento, ventaBanco.IdAsiento) IS NOT NULL
                        THEN CONVERT(CHAR(6), COALESCE(v.FechaContabilizacion, ventaBanco.FechaContabilizacion), 112) + '00'
                    ELSE ''
                END AS PeriodoReferencia,
                CASE
                    WHEN asientoReferencia.IdAsiento IS NOT NULL
                        THEN CONCAT(
                            LTRIM(RTRIM(origenReferencia.CodigoOrigen)),
                            asientoReferencia.Periodo,
                            CASE
                                WHEN LEN(CONVERT(VARCHAR(20), asientoReferencia.NumeroAsiento)) < 8
                                    THEN RIGHT(REPLICATE('0', 8) + CONVERT(VARCHAR(20), asientoReferencia.NumeroAsiento), 8)
                                ELSE CONVERT(VARCHAR(20), asientoReferencia.NumeroAsiento)
                            END
                        )
                    ELSE ''
                END AS CuoReferencia,
                CASE
                    WHEN asientoReferencia.IdAsiento IS NOT NULL THEN 'M0001'
                    ELSE ''
                END AS CorrelativoReferencia
        ) AS origen
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
