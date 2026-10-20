-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Obtiene un nivel para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantNiveles_Obtener
    @IdEmpresa CHAR(2),
    @IdNivel CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            N.IdNivel, N.GlsNivel, N.IdTipoNivel, N.IdNivelPred,
            ISNULL(N.Orden, 0) AS Orden, N.Tipo, N.PorcComision,
            N.MaxDcto, N.GlsAbreviatura, N.IdCtaContableC,
            N.IdCtaContableV, N.IdCtaContableVR
        FROM dbo.Niveles AS N
        WHERE N.IdEmpresa = @IdEmpresa AND N.IdNivel = @IdNivel;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
