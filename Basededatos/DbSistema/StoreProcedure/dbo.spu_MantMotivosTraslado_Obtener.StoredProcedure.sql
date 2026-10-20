-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 16/10/2026
-- Description: Obtiene un motivo de traslado para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantMotivosTraslado_Obtener
    @IdMotivoTraslado CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT MT.IdMotivoTraslado, MT.GlsMotivoTraslado, MT.IndTipoVale,
               MT.IndGeneraVale, MT.IdDocumento, MT.IndEvaluaLC
        FROM dbo.MotivosTraslados AS MT
        WHERE MT.IdMotivoTraslado = @IdMotivoTraslado;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(), @ErrorSeverity=ERROR_SEVERITY(), @ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
