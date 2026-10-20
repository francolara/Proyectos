-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Obtiene un documento y sus cuentas por empresa.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantDocumentos_Obtener
    @IdEmpresa CHAR(2),
    @IdDocumento CHAR(2)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            D.IdDocumento, D.GlsDocumento, D.AbreDocumento, D.EstDocumentos,
            DC.IdCtaVtaSoles, DC.IdCtaVtaDolares, DC.IdCtaVtaTerSoles, DC.IdCtaVtaTerDolares,
            DC.IdCtaCompSoles, DC.IdCtaCompDolares, DC.IdCtaCompTerSoles, DC.IdCtaCompTerDolares,
            DC.IdCtaCompTerMatrizS, DC.IdCtaCompTerMatrizD, DC.IdCtaCompTerSubSiS, DC.IdCtaCompTerSubSiD,
            DC.IdCtaVtaTerMatrizS, DC.IdCtaVtaTerMatrizD, DC.IdCtaVtaTerSubSiS, DC.IdCtaVtaTerSubSiD,
            DC.IdCtaVtaPerS, DC.IdCtaVtaPerD, DC.IdCtaCompPerS, DC.IdCtaCompPerD,
            DC.IdCtaVtaSocS, DC.IdCtaVtaSocD, DC.IdCtaCompSocS, DC.IdCtaCompSocD
        FROM dbo.Documentos AS D
        LEFT JOIN dbo.DocumentosCuentas AS DC
            ON DC.IdEmpresa = @IdEmpresa AND D.IdDocumento = DC.IdDocumento
        WHERE D.IdDocumento = @IdDocumento;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
