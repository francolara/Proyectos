-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Inserta o actualiza el perfil complementario del usuario autenticado.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_GuardarUsuarioPerfil
    @AspNetUserId NVARCHAR(450),
    @NombreCompleto NVARCHAR(180),
    @Telefono NVARCHAR(30) = NULL,
    @CorreoReferencia NVARCHAR(256) = NULL,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        IF EXISTS
        (
            SELECT 1
            FROM dbo.SEG_UsuarioPerfil AS up
            WHERE up.AspNetUserId = @AspNetUserId
        )
        BEGIN
            UPDATE dbo.SEG_UsuarioPerfil
            SET
                NombreCompleto = @NombreCompleto,
                Telefono = @Telefono,
                CorreoReferencia = @CorreoReferencia,
                Estado = 1,
                UsuarioRegistro = COALESCE(@UsuarioRegistro, UsuarioRegistro)
            WHERE AspNetUserId = @AspNetUserId;
        END
        ELSE
        BEGIN
            INSERT INTO dbo.SEG_UsuarioPerfil
            (
                AspNetUserId,
                NombreCompleto,
                Telefono,
                CorreoReferencia,
                Estado,
                UsuarioRegistro
            )
            VALUES
            (
                @AspNetUserId,
                @NombreCompleto,
                @Telefono,
                @CorreoReferencia,
                1,
                @UsuarioRegistro
            );
        END;

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
