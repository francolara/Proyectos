USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 35_Maestros_FormasPago.sql (linea 156)
-- Firma: Codex - 09/04/2026 | Lista pagos agrupados por reserva con filtro/paginacion backend, corrige alcance de consulta materializando filtrado en tabla temporal, muestra monto de reserva, saldo/simbolo y banderas PagadaCompleta/TieneComprobanteActivo para habilitar emision de comprobantes.
-- Firma: Codex - 12/04/2026 | Agrega columna Referencia en listado de pagos con el ultimo comprobante principal activo por reserva (boleta/factura/recibo interno), tomando el ultimo generado por Id; oculta referencia cuando el comprobante principal esta anulado o cuando boleta/factura tiene NC/ND activas.
-- Firma: Codex - 12/04/2026 | Usa abreviatura del documento (TiposDocumentoComprobanteSuperMaestro.Abreviatura) en columna Referencia.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Pagos_Listar]
    @NegocioId INT,
    @SedeId INT = NULL,
    @Buscar NVARCHAR(120) = NULL,
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
            Referencia NVARCHAR(120) NOT NULL
        );

        ;WITH ReservasConPago AS
        (
            SELECT
                r.Id AS ReservaId,
                s.Nombre AS Sede,
                e.Nombre AS Espacio,
                c.NombresORazonSocial AS Cliente,
                r.Fecha,
                CAST(r.Total AS DECIMAL(10,2)) AS MontoTotal,
                CAST(r.Total - SUM(p.Monto) AS DECIMAL(10,2)) AS SaldoPendiente,
                CAST(CASE WHEN r.Estado = 4 AND (r.Total - SUM(p.Monto)) <= 0 THEN 1 ELSE 0 END AS BIT) AS PagadaCompleta,
                COUNT(p.Id) AS CantidadPagos,
                STRING_AGG(fp.Nombre, N', ') WITHIN GROUP (ORDER BY fp.Nombre) AS FormaPagoResumen,
                COALESCE(ms.Simbolo, N'S/') AS MonedaSimbolo
            FROM dbo.Reservas r
            INNER JOIN dbo.Pagos p ON p.ReservaId = r.Id
            INNER JOIN dbo.FormasPago fp ON fp.Id = p.FormaPago
            INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
            INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
            LEFT JOIN dbo.Monedas m ON m.Id = n.MonedaId
            LEFT JOIN dbo.MonedasSuperMaestro ms ON ms.Id = m.MonedaSuperId
            INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
            WHERE s.NegocioId = @NegocioId
              AND (@SedeId IS NULL OR s.Id = @SedeId)
            GROUP BY r.Id, s.Nombre, e.Nombre, c.NombresORazonSocial, r.Fecha, r.Total, r.Estado, ms.Simbolo
        ),
        ComprobantesPrincipales AS
        (
            SELECT
                ce.ReservaId,
                ce.Id AS ComprobanteId,
                tdc.CodigoSunat AS CodigoDocumento,
                COALESCE(tdsm.Abreviatura, tdsm.Nombre, N'Comp.') AS TipoDocumentoNombre,
                ce.Serie,
                ce.Numero,
                CAST(
                    CASE WHEN EXISTS
                    (
                        SELECT 1
                        FROM dbo.ComprobantesElectronicos n
                        INNER JOIN dbo.NegociosTiposDocumentoComprobante tdn ON tdn.Id = n.TipoComprobante
                        WHERE n.NegocioId = ce.NegocioId
                          AND n.ComprobanteReferenciaId = ce.Id
                          AND n.Estado <> 5
                          AND tdn.CodigoSunat IN (N'07', N'08')
                    ) THEN 1 ELSE 0 END
                AS BIT) AS TieneNotaActiva,
                ROW_NUMBER() OVER (PARTITION BY ce.ReservaId ORDER BY ce.Id DESC) AS rn
            FROM dbo.ComprobantesElectronicos ce
            INNER JOIN dbo.NegociosTiposDocumentoComprobante tdc ON tdc.Id = ce.TipoComprobante
            LEFT JOIN dbo.TiposDocumentoComprobanteSuperMaestro tdsm ON tdsm.CodigoSunat = tdc.CodigoSunat
            WHERE ce.NegocioId = @NegocioId
              AND ce.ReservaId IS NOT NULL
              AND ce.ComprobanteReferenciaId IS NULL
              AND ce.Estado <> 5
              AND tdc.CodigoSunat IN (N'01', N'03', N'RI')
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
            Referencia
        )
        SELECT
            x.ReservaId,
            CONCAT(N'#', CONVERT(NVARCHAR(20), x.ReservaId)) AS ReservaCodigo,
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
                          INNER JOIN dbo.NegociosTiposDocumentoComprobante ntdNc ON ntdNc.Id = nc.TipoComprobante
                          WHERE nc.NegocioId = ce.NegocioId
                            AND nc.ComprobanteReferenciaId = ce.Id
                            AND nc.Estado <> 5
                            AND ntdNc.CodigoSunat = N'07'
                      )
                ) THEN 1 ELSE 0
            END AS BIT) AS TieneComprobanteActivo,
            CASE
                WHEN u.ReservaId IS NULL THEN N''
                WHEN u.CodigoDocumento IN (N'01', N'03') AND u.TieneNotaActiva = 1 THEN N''
                ELSE CONCAT(u.TipoDocumentoNombre, N' ', u.Serie, N'-', FORMAT(u.Numero, '00000000'))
            END AS Referencia
        FROM ReservasConPago x
        LEFT JOIN UltimoComprobantePrincipal u ON u.ReservaId = x.ReservaId
        WHERE @BuscarTrim IS NULL
           OR CONVERT(NVARCHAR(20), x.ReservaId) LIKE N'%' + @BuscarTrim + N'%'
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
            Referencia
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
