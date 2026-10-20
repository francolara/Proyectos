-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Obtiene un concepto para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantConceptos_Obtener
    @IdConcepto CHAR(2)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            C.IdConcepto,
            C.GlsConcepto,
            C.TipoConcepto,
            C.IndCosto
        FROM dbo.Conceptos AS C
        WHERE C.IdConcepto = @IdConcepto;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
