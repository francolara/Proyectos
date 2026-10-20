-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Obtiene una marca para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantMarcas_Obtener
    @IdEmpresa CHAR(2),
    @IdMarca CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            M.IdMarca,
            M.GlsMarca,
            M.MaxDcto
        FROM dbo.Marcas AS M
        WHERE M.IdEmpresa = @IdEmpresa
          AND M.IdMarca = @IdMarca;
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
