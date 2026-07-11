-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   10/07/2026
-- Description:   Desactiva el acceso del usuario a la cuenta administradora y a sus empresas vinculadas de la misma cuenta.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_DesactivarUsuarioCuentaAdministradora
    @AspNetUserId NVARCHAR(450),
    @IdCuentaAdministradora INT,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        UPDATE dbo.SEG_UsuarioCuentaAdministradora
        SET Estado = 0,
            UsuarioRegistro = @UsuarioRegistro
        WHERE AspNetUserId = @AspNetUserId
          AND IdCuentaAdministradora = @IdCuentaAdministradora;

        UPDATE ue
        SET ue.Estado = 0,
            ue.UsuarioRegistro = @UsuarioRegistro
        FROM dbo.SEG_UsuarioEmpresa AS ue
        INNER JOIN dbo.SEG_Empresa AS e
            ON e.IdEmpresa = ue.IdEmpresa
        WHERE ue.AspNetUserId = @AspNetUserId
          AND e.IdCuentaAdministradora = @IdCuentaAdministradora;

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
