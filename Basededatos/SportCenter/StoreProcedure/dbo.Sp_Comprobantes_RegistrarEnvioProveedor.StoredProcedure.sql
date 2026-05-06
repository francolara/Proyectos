USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 05/05/2026 | Registra resultado de envio de comprobante al proveedor electronico (estado/codigo/mensaje/ticket/hash).
-- Firma: Codex - 05/05/2026 | Agrega persistencia de URLs SUNAT (PDF/XML/CDR) por proveedor.
CREATE OR ALTER PROCEDURE dbo.Sp_Comprobantes_RegistrarEnvioProveedor
    @NegocioId INT,
    @Id INT,
    @Estado INT,
    @CodigoRespuesta NVARCHAR(50) = NULL,
    @MensajeRespuesta NVARCHAR(500) = NULL,
    @NumeroTicketSunat NVARCHAR(40) = NULL,
    @CodigoHashCpe NVARCHAR(100) = NULL,
    @UrlPdfSunat NVARCHAR(500) = NULL,
    @UrlXmlSunat NVARCHAR(500) = NULL,
    @UrlCdrSunat NVARCHAR(500) = NULL,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.ComprobantesElectronicos c
            WHERE c.Id = @Id
              AND c.NegocioId = @NegocioId
        )
            RAISERROR('No se encontro el comprobante para registrar envio.', 16, 1);

        UPDATE dbo.ComprobantesElectronicos
        SET
            Estado = @Estado,
            CodigoRespuestaSunat = NULLIF(LTRIM(RTRIM(@CodigoRespuesta)), N''),
            MensajeRespuestaSunat = NULLIF(LTRIM(RTRIM(@MensajeRespuesta)), N''),
            NumeroTicketSunat = NULLIF(LTRIM(RTRIM(@NumeroTicketSunat)), N''),
            CodigoHashCpe = NULLIF(LTRIM(RTRIM(@CodigoHashCpe)), N''),
            UrlPdfSunat = NULLIF(LTRIM(RTRIM(@UrlPdfSunat)), N''),
            UrlXmlSunat = NULLIF(LTRIM(RTRIM(@UrlXmlSunat)), N''),
            UrlCdrSunat = NULLIF(LTRIM(RTRIM(@UrlCdrSunat)), N''),
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id
          AND NegocioId = @NegocioId;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
