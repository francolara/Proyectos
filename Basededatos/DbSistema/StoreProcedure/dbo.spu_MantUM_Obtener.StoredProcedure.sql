-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Obtiene una unidad de medida para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantUM_Obtener
    @IdUM CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            UM.IdUM,
            UM.GlsUM,
            UM.AbreUM
        FROM dbo.UnidadMedida AS UM
        WHERE UM.IdUM = @IdUM;
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
