-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   10/07/2026
-- Description:   Desactiva la relacion entre usuario y empresa solo para la cuenta administradora indicada.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_DesactivarUsuarioEmpresa
    @AspNetUserId NVARCHAR(450),
    @IdEmpresa INT,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        UPDATE dbo.SEG_UsuarioEmpresa
        SET Estado = 0,
            UsuarioRegistro = @UsuarioRegistro
        WHERE AspNetUserId = @AspNetUserId
          AND IdEmpresa = @IdEmpresa;

    END TRY

    BEGIN CATCH

        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE()

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)

    END CATCH

END
