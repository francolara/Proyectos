-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 16/10/2026
-- Description: Lista documentos que tienen motivos de traslado configurados.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantMotivosTraslado_Documentos_Listar
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT D.GlsDocumento, D.IdDocumento
        FROM dbo.Documentos AS D
        WHERE EXISTS (
            SELECT 1
            FROM dbo.MotivosTraslados AS MT
            WHERE MT.IdDocumento = D.IdDocumento
        )
        ORDER BY D.GlsDocumento;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(), @ErrorSeverity=ERROR_SEVERITY(), @ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
