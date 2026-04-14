USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 04/04/2026 | Actualizacion individual de Sp_ConfiguracionClub_Obtener por ubigeo fiscal y tipo de documento SUNAT centralizado.
-- Firma: Codex - 06/04/2026 | Se agrega politica de confirmacion de reserva por pago y porcentaje minimo de adelanto a nivel negocio.
-- Firma: Codex - 09/04/2026 | Se agrega configuracion de emision (CPE/Recibo interno) y porcentaje IGV.
-- Firma: Codex - 13/04/2026 | Se agrega LogoUrl para administracion de logo del negocio.
CREATE OR ALTER PROCEDURE dbo.Sp_ConfiguracionClub_Obtener
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            n.Id,
            n.NombreComercial,
            n.RazonSocial,
            COALESCE(NULLIF(n.TipoDocumentoFiscal, N''), N'1') AS TipoDocumentoFiscal,
            COALESCE(NULLIF(n.NumeroDocumentoFiscal, N''), n.DocumentoFiscal) AS NumeroDocumentoFiscal,
            n.DireccionFiscal,
            COALESCE(n.MonedaId, 1) AS MonedaId,
            n.CodigoUbigeo,
            CAST(COALESCE(n.PoliticaConfirmacionPago, 0) AS TINYINT) AS PoliticaConfirmacionPago,
            n.PorcentajeAdelantoMinimo,
            CAST(COALESCE(n.EmisionComprobantesElectronicos, 0) AS BIT) AS EmisionComprobantesElectronicos,
            CAST(COALESCE(n.EmisionReciboInterno, 0) AS BIT) AS EmisionReciboInterno,
            CAST(COALESCE(n.PorcentajeIgv, 18) AS INT) AS PorcentajeIgv,
            n.LogoUrl
        FROM dbo.Negocios n
        WHERE n.Id = @NegocioId
          AND n.Activo = 1;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
