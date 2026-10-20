-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 16/10/2026
-- Description: Obtiene un perfil de la empresa y sistema indicados.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantPerfiles_Obtener
    @IdEmpresa CHAR(2),
    @CodSistema CHAR(2),
    @IdPerfil CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT P.IdPerfil, P.GlsPerfil
        FROM dbo.Perfil AS P
        WHERE P.IdEmpresa = @IdEmpresa
          AND P.CodSistema = @CodSistema
          AND P.IdPerfil = @IdPerfil;
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
