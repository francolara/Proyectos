
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 04_Reservas_Pagos_Comprobantes.sql (linea 286)
-- Firma: Codex - 09/04/2026 | Obtiene cabecera de reserva (incluye horario, moneda y politica) y detalle de pagos para crear/editar pagos.
-- Firma: Codex - 12/04/2026 | Incluye bandera de bloqueo por comprobante activo y referencia del ultimo comprobante principal (ultimo generado por Id) para forzar edicion solo lectura en pagos cuando ya se emitio documento.
-- Firma: Codex - 12/04/2026 | Usa abreviatura del documento (TiposDocumentoComprobanteSuperMaestro.Abreviatura) en ReferenciaComprobante.
-- Firma: FRANCO LARA - 17/09/2026 | Incluye trazabilidad de cada pago para su visualizacion durante la edicion.
-- Firma: FRANCO LARA - 01/10/2026 | Muestra correlativos visibles por negocio para la reserva y el detalle de pagos.
-- Firma: FRANCO LARA - 06/10/2026 | Presenta y expone CodigoMoneda historico de la reserva y usa codigos SUNAT canonicos.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Pagos_ObtenerPorId]
    @NegocioId INT,
    @Id INT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            r.Id AS ReservaId,
            CONCAT(N'R-', RIGHT(N'000000' + CONVERT(NVARCHAR(20), r.NumeroPorNegocio), 6)) AS ReservaCodigo,
            s.Nombre AS Sede,
            e.Nombre AS Espacio,
            c.NombresORazonSocial AS Cliente,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin,
            r.Total AS TotalReserva,
            COALESCE(SUM(p.Monto), 0) AS TotalPagado,
            (r.Total - COALESCE(SUM(p.Monto), 0)) AS SaldoPendiente,
            COALESCE(ms.Simbolo, r.CodigoMoneda) AS MonedaSimbolo,
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
                          WHERE nc.NegocioId = cex.NegocioId
                            AND nc.ComprobanteReferenciaId = cex.Id
                            AND nc.Estado <> 5
                            AND nc.CodigoTipoComprobante = N'07'
                      )
                ) THEN 1 ELSE 0 END
            AS BIT) AS TieneComprobanteActivo,
            COALESCE
            (
                (
                    SELECT TOP (1)
                        CASE
                            WHEN ce.CodigoTipoComprobante IN (N'01', N'03') AND EXISTS
                            (
                                SELECT 1
                                FROM dbo.ComprobantesElectronicos nrel
                                WHERE nrel.NegocioId = ce.NegocioId
                                  AND nrel.ComprobanteReferenciaId = ce.Id
                                  AND nrel.Estado <> 5
                                  AND nrel.CodigoTipoComprobante IN (N'07', N'08')
                            ) THEN N''
                            ELSE CONCAT(
                                COALESCE(tdsm.Abreviatura, tdsm.Nombre, N'Comp.'),
                                N' ',
                                ce.Serie,
                                N'-',
                                FORMAT(ce.Numero, '00000000'))
                        END
                    FROM dbo.ComprobantesElectronicos ce
                    LEFT JOIN dbo.TiposDocumentoComprobanteSuperMaestro tdsm ON tdsm.CodigoSunat = ce.CodigoTipoComprobante
                    WHERE ce.NegocioId = @NegocioId
                      AND ce.ReservaId = r.Id
                      AND ce.ComprobanteReferenciaId IS NULL
                      AND ce.Estado <> 5
                      AND ce.CodigoTipoComprobante IN (N'01', N'03', N'RI')
                    ORDER BY ce.Id DESC
                ),
                N''
            ) AS ReferenciaComprobante,
            r.CodigoMoneda
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
        INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
        LEFT JOIN dbo.MonedasSuperMaestro ms ON ms.Codigo = r.CodigoMoneda
        LEFT JOIN dbo.Pagos p ON p.ReservaId = r.Id
        WHERE r.Id = @Id
          AND s.NegocioId = @NegocioId
        GROUP BY
            r.Id,
            r.NumeroPorNegocio,
            s.Nombre,
            e.Nombre,
            c.NombresORazonSocial,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin,
            r.Total,
            r.CodigoMoneda,
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
            p.Observacion,
            p.UsuarioCreacion,
            CAST(p.FechaCreacion AT TIME ZONE 'UTC' AT TIME ZONE 'SA Pacific Standard Time' AS DATETIME2) AS FechaRegistro,
            p.UsuarioActualizacion,
            CAST(p.FechaActualizacion AT TIME ZONE 'UTC' AT TIME ZONE 'SA Pacific Standard Time' AS DATETIME2) AS FechaActualizacion,
            p.NumeroPorNegocio
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
