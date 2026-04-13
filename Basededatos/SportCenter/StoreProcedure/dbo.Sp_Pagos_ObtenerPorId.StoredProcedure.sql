USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 04_Reservas_Pagos_Comprobantes.sql (linea 286)
-- Firma: Codex - 09/04/2026 | Obtiene cabecera de reserva (incluye horario, moneda y politica) y detalle de pagos para crear/editar pagos.
-- Firma: Codex - 12/04/2026 | Incluye bandera de bloqueo por comprobante activo y referencia del ultimo comprobante principal (ultimo generado por Id) para forzar edicion solo lectura en pagos cuando ya se emitio documento.
-- Firma: Codex - 12/04/2026 | Usa abreviatura del documento (TiposDocumentoComprobanteSuperMaestro.Abreviatura) en ReferenciaComprobante.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Pagos_ObtenerPorId]
    @NegocioId INT,
    @Id INT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            r.Id AS ReservaId,
            CONCAT(N'#', CONVERT(NVARCHAR(20), r.Id)) AS ReservaCodigo,
            s.Nombre AS Sede,
            e.Nombre AS Espacio,
            c.NombresORazonSocial AS Cliente,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin,
            r.Total AS TotalReserva,
            COALESCE(SUM(p.Monto), 0) AS TotalPagado,
            (r.Total - COALESCE(SUM(p.Monto), 0)) AS SaldoPendiente,
            COALESCE(ms.Simbolo, N'S/') AS MonedaSimbolo,
            CAST(ISNULL(n.PoliticaConfirmacionPago, 0) AS INT) AS PoliticaConfirmacionPago,
            n.PorcentajeAdelantoMinimo,
            CAST(
                CASE WHEN EXISTS
                (
                    SELECT 1
                    FROM dbo.ComprobantesElectronicos cex
                    WHERE cex.NegocioId = @NegocioId
                      AND cex.ReservaId = r.Id
                      AND cex.ComprobanteReferenciaId IS NULL
                      AND cex.Estado <> 5
                      AND NOT EXISTS
                      (
                          SELECT 1
                          FROM dbo.ComprobantesElectronicos nc
                          INNER JOIN dbo.NegociosTiposDocumentoComprobante ntdNc ON ntdNc.Id = nc.TipoComprobante
                          WHERE nc.NegocioId = cex.NegocioId
                            AND nc.ComprobanteReferenciaId = cex.Id
                            AND nc.Estado <> 5
                            AND ntdNc.CodigoSunat = N'07'
                      )
                ) THEN 1 ELSE 0 END
            AS BIT) AS TieneComprobanteActivo,
            COALESCE
            (
                (
                    SELECT TOP (1)
                        CASE
                            WHEN ntd.CodigoSunat IN (N'01', N'03') AND EXISTS
                            (
                                SELECT 1
                                FROM dbo.ComprobantesElectronicos nrel
                                INNER JOIN dbo.NegociosTiposDocumentoComprobante ntdRel ON ntdRel.Id = nrel.TipoComprobante
                                WHERE nrel.NegocioId = ce.NegocioId
                                  AND nrel.ComprobanteReferenciaId = ce.Id
                                  AND nrel.Estado <> 5
                                  AND ntdRel.CodigoSunat IN (N'07', N'08')
                            ) THEN N''
                            ELSE CONCAT(
                                COALESCE(tdsm.Abreviatura, tdsm.Nombre, N'Comp.'),
                                N' ',
                                ce.Serie,
                                N'-',
                                FORMAT(ce.Numero, '00000000'))
                        END
                    FROM dbo.ComprobantesElectronicos ce
                    INNER JOIN dbo.NegociosTiposDocumentoComprobante ntd ON ntd.Id = ce.TipoComprobante
                    LEFT JOIN dbo.TiposDocumentoComprobanteSuperMaestro tdsm ON tdsm.CodigoSunat = ntd.CodigoSunat
                    WHERE ce.NegocioId = @NegocioId
                      AND ce.ReservaId = r.Id
                      AND ce.ComprobanteReferenciaId IS NULL
                      AND ce.Estado <> 5
                      AND ntd.CodigoSunat IN (N'01', N'03', N'RI')
                    ORDER BY ce.Id DESC
                ),
                N''
            ) AS ReferenciaComprobante
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
        INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
        LEFT JOIN dbo.Monedas m ON m.Id = n.MonedaId
        LEFT JOIN dbo.MonedasSuperMaestro ms ON ms.Id = m.MonedaSuperId
        LEFT JOIN dbo.Pagos p ON p.ReservaId = r.Id
        WHERE r.Id = @Id
          AND s.NegocioId = @NegocioId
        GROUP BY
            r.Id,
            s.Nombre,
            e.Nombre,
            c.NombresORazonSocial,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin,
            r.Total,
            ms.Simbolo,
            n.PoliticaConfirmacionPago,
            n.PorcentajeAdelantoMinimo;

        SELECT
            p.Id,
            p.FechaPago,
            p.Monto,
            p.FormaPago,
            fp.Nombre AS FormaPagoNombre,
            p.NumeroOperacion,
            p.Observacion
        FROM dbo.Pagos p
        INNER JOIN dbo.FormasPago fp ON fp.Id = p.FormaPago
        INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE r.Id = @Id
          AND s.NegocioId = @NegocioId
        ORDER BY p.FechaPago, p.Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
