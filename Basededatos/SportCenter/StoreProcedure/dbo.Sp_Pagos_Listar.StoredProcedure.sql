
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 35_Maestros_FormasPago.sql (linea 156)
-- Firma: Codex - 09/04/2026 | Lista pagos agrupados por reserva con filtro/paginacion backend, corrige alcance de consulta materializando filtrado en tabla temporal, muestra monto de reserva, saldo/simbolo y banderas PagadaCompleta/TieneComprobanteActivo para habilitar emision de comprobantes.
-- Firma: Codex - 12/04/2026 | Agrega columna Referencia en listado de pagos con el ultimo comprobante principal activo por reserva (boleta/factura/recibo interno), tomando el ultimo generado por Id; oculta referencia cuando el comprobante principal esta anulado o cuando boleta/factura tiene NC/ND activas.
-- Firma: Codex - 12/04/2026 | Usa abreviatura del documento (TiposDocumentoComprobanteSuperMaestro.Abreviatura) en columna Referencia.
-- Firma: Codex - 13/04/2026 | Agrega filtro opcional por rango de fecha de reserva (Desde/Hasta) para listado de pagos.
-- Firma: Codex - 18/06/2026 | Cambia el filtro del listado de pagos para usar FechaPago real y expone la ultima fecha de pago por reserva en la grilla.
-- Firma: FRANCO LARA - 01/10/2026 | Usa el correlativo visible por negocio para identificar y buscar la reserva en el listado de pagos.
-- Firma: FRANCO LARA - 06/10/2026 | Lee simbolo/codigo de moneda y comprobantes desde codigos canonicos historicos.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Pagos_Listar]
    @NegocioId INT,
    @SedeId INT = NULL,
    @Buscar NVARCHAR(120) = NULL,
    @FechaDesde DATE = NULL,
    @FechaHasta DATE = NULL,
    @Pagina INT = 1,
    @TamanoPagina INT = 20,
    @TotalRegistros INT OUTPUT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        IF @Pagina < 1 SET @Pagina = 1;
        IF @TamanoPagina < 1 SET @TamanoPagina = 20;

        DECLARE @Offset INT = (@Pagina - 1) * @TamanoPagina;
        DECLARE @BuscarTrim NVARCHAR(120) = NULLIF(LTRIM(RTRIM(@Buscar)), N'');
        DECLARE @BuscarNumeroReserva INT = TRY_CONVERT(
            INT,
            CASE
                WHEN UPPER(LEFT(@BuscarTrim, 2)) = N'R-' THEN SUBSTRING(@BuscarTrim, 3, 120)
                ELSE NULL
            END
        );

        CREATE TABLE #ReservasFiltradas
        (
            ReservaId INT NOT NULL,
            ReservaCodigo NVARCHAR(25) NOT NULL,
            Sede NVARCHAR(200) NOT NULL,
            Espacio NVARCHAR(200) NOT NULL,
            Cliente NVARCHAR(200) NOT NULL,
            Fecha DATE NOT NULL,
            MontoTotal DECIMAL(10,2) NOT NULL,
            SaldoPendiente DECIMAL(10,2) NOT NULL,
            FormaPagoResumen NVARCHAR(500) NOT NULL,
            CantidadPagos INT NOT NULL,
            MonedaSimbolo NVARCHAR(10) NOT NULL,
            PagadaCompleta BIT NOT NULL,
            TieneComprobanteActivo BIT NOT NULL,
            Referencia NVARCHAR(120) NOT NULL,
            CodigoMoneda NVARCHAR(10) NOT NULL
        );

        ;WITH ReservasConPago AS
        (
            SELECT
                r.Id AS ReservaId,
                r.NumeroPorNegocio AS ReservaNumeroPorNegocio,
                s.Nombre AS Sede,
                e.Nombre AS Espacio,
                c.NombresORazonSocial AS Cliente,
                MAX(CAST(p.FechaPago AS DATE)) AS Fecha,
                CAST(r.Total AS DECIMAL(10,2)) AS MontoTotal,
                CAST(r.Total - SUM(p.Monto) AS DECIMAL(10,2)) AS SaldoPendiente,
                CAST(CASE WHEN r.Estado = 4 AND (r.Total - SUM(p.Monto)) <= 0 THEN 1 ELSE 0 END AS BIT) AS PagadaCompleta,
                COUNT(p.Id) AS CantidadPagos,
                STRING_AGG(fp.Nombre, N', ') WITHIN GROUP (ORDER BY fp.Nombre) AS FormaPagoResumen,
                COALESCE(ms.Simbolo, r.CodigoMoneda) AS MonedaSimbolo,
                r.CodigoMoneda
            FROM dbo.Reservas r
            INNER JOIN dbo.Pagos p ON p.ReservaId = r.Id
            INNER JOIN dbo.FormasPago fp ON fp.Id = p.FormaPago
            INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
            LEFT JOIN dbo.MonedasSuperMaestro ms ON ms.Codigo = r.CodigoMoneda
            INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
            WHERE s.NegocioId = @NegocioId
              AND (@SedeId IS NULL OR s.Id = @SedeId)
              AND (@FechaDesde IS NULL OR CAST(p.FechaPago AS DATE) >= @FechaDesde)
              AND (@FechaHasta IS NULL OR CAST(p.FechaPago AS DATE) <= @FechaHasta)
            GROUP BY r.Id, r.NumeroPorNegocio, s.Nombre, e.Nombre, c.NombresORazonSocial, r.Total, r.Estado, r.CodigoMoneda, ms.Simbolo
        ),
        ComprobantesPrincipales AS
        (
            SELECT
                ce.ReservaId,
                ce.Id AS ComprobanteId,
                ce.CodigoTipoComprobante AS CodigoDocumento,
                COALESCE(tdsm.Abreviatura, tdsm.Nombre, N'Comp.') AS TipoDocumentoNombre,
                ce.Serie,
                ce.Numero,
                CAST(
                    CASE WHEN EXISTS
                    (
                        SELECT 1
                        FROM dbo.ComprobantesElectronicos n
                        WHERE n.NegocioId = ce.NegocioId
                          AND n.ComprobanteReferenciaId = ce.Id
                          AND n.Estado <> 5
                          AND n.CodigoTipoComprobante IN (N'07', N'08')
                    ) THEN 1 ELSE 0 END
                AS BIT) AS TieneNotaActiva,
                ROW_NUMBER() OVER (PARTITION BY ce.ReservaId ORDER BY ce.Id DESC) AS rn
            FROM dbo.ComprobantesElectronicos ce
            LEFT JOIN dbo.TiposDocumentoComprobanteSuperMaestro tdsm ON tdsm.CodigoSunat = ce.CodigoTipoComprobante
            WHERE ce.NegocioId = @NegocioId
              AND ce.ReservaId IS NOT NULL
              AND ce.ComprobanteReferenciaId IS NULL
              AND ce.Estado <> 5
              AND ce.CodigoTipoComprobante IN (N'01', N'03', N'RI')
        ),
        UltimoComprobantePrincipal AS
        (
            SELECT
                ReservaId,
                CodigoDocumento,
                TipoDocumentoNombre,
                Serie,
                Numero,
                TieneNotaActiva
            FROM ComprobantesPrincipales
            WHERE rn = 1
        )
        INSERT INTO #ReservasFiltradas
        (
            ReservaId,
            ReservaCodigo,
            Sede,
            Espacio,
            Cliente,
            Fecha,
            MontoTotal,
            SaldoPendiente,
            FormaPagoResumen,
            CantidadPagos,
            MonedaSimbolo,
            PagadaCompleta,
            TieneComprobanteActivo,
            Referencia,
            CodigoMoneda
        )
        SELECT
            x.ReservaId,
            CONCAT(N'R-', RIGHT(N'000000' + CONVERT(NVARCHAR(20), x.ReservaNumeroPorNegocio), 6)) AS ReservaCodigo,
            x.Sede,
            x.Espacio,
            x.Cliente,
            x.Fecha,
            x.MontoTotal,
            x.SaldoPendiente,
            x.FormaPagoResumen,
            x.CantidadPagos,
            x.MonedaSimbolo,
            x.PagadaCompleta,
            CAST(CASE
                WHEN EXISTS
                (
                    SELECT 1
                    FROM dbo.ComprobantesElectronicos ce
                    WHERE ce.NegocioId = @NegocioId
                      AND ce.ReservaId = x.ReservaId
                      AND ce.ComprobanteReferenciaId IS NULL
                      AND ce.Estado <> 5
                      AND NOT EXISTS
                      (
                          SELECT 1
                          FROM dbo.ComprobantesElectronicos nc
                            WHERE nc.NegocioId = ce.NegocioId
                            AND nc.ComprobanteReferenciaId = ce.Id
                            AND nc.Estado <> 5
                            AND nc.CodigoTipoComprobante = N'07'
                      )
                ) THEN 1 ELSE 0
            END AS BIT) AS TieneComprobanteActivo,
            CASE
                WHEN u.ReservaId IS NULL THEN N''
                WHEN u.CodigoDocumento IN (N'01', N'03') AND u.TieneNotaActiva = 1 THEN N''
                ELSE CONCAT(u.TipoDocumentoNombre, N' ', u.Serie, N'-', FORMAT(u.Numero, '00000000'))
            END AS Referencia,
            x.CodigoMoneda
        FROM ReservasConPago x
        LEFT JOIN UltimoComprobantePrincipal u ON u.ReservaId = x.ReservaId
        WHERE @BuscarTrim IS NULL
           OR CONVERT(NVARCHAR(20), x.ReservaNumeroPorNegocio) LIKE N'%' + @BuscarTrim + N'%'
           OR (@BuscarNumeroReserva IS NOT NULL AND x.ReservaNumeroPorNegocio = @BuscarNumeroReserva)
           OR x.Sede LIKE N'%' + @BuscarTrim + N'%'
           OR x.Espacio LIKE N'%' + @BuscarTrim + N'%'
           OR x.Cliente LIKE N'%' + @BuscarTrim + N'%'
           OR x.FormaPagoResumen LIKE N'%' + @BuscarTrim + N'%'
           OR CONVERT(NVARCHAR(10), x.Fecha, 103) LIKE N'%' + @BuscarTrim + N'%';

        SELECT @TotalRegistros = COUNT(1)
        FROM #ReservasFiltradas;

        SELECT
            ReservaId,
            ReservaCodigo,
            Sede,
            Espacio,
            Cliente,
            Fecha,
            MontoTotal,
            SaldoPendiente,
            FormaPagoResumen,
            CantidadPagos,
            MonedaSimbolo,
            PagadaCompleta,
            TieneComprobanteActivo,
            Referencia,
            CodigoMoneda
        FROM #ReservasFiltradas
        ORDER BY Fecha DESC, ReservaId DESC
        OFFSET @Offset ROWS FETCH NEXT @TamanoPagina ROWS ONLY;

        DROP TABLE #ReservasFiltradas;
    END TRY
    BEGIN CATCH
        IF OBJECT_ID('tempdb..#ReservasFiltradas') IS NOT NULL
            DROP TABLE #ReservasFiltradas;

        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
