-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Obtiene un centro de costo para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantCentroCosto_Obtener
    @IdEmpresa CHAR(2),
    @IdCentroCosto CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            CC.IdCentroCosto,
            CC.GlsCentroCosto,
            CC.IndEstado,
            CC.IndGenerado,
            CC.AbrevCcosto,
            CC.IdInterface
        FROM dbo.CentrosCosto AS CC
        WHERE CC.IdEmpresa = @IdEmpresa
          AND CC.IdCentroCosto = @IdCentroCosto;
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
