-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Asigna o reactiva la relacion entre usuario y empresa.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_AsignarUsuarioEmpresa
    @AspNetUserId NVARCHAR(450),
    @IdEmpresa INT,
    @EsEmpresaPredeterminada BIT,
    @UsuarioRegistro NVARCHAR(450)
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        IF @EsEmpresaPredeterminada = 1
        BEGIN
            UPDATE dbo.SEG_UsuarioEmpresa
            SET EsEmpresaPredeterminada = 0
            WHERE AspNetUserId = @AspNetUserId;
        END;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.SEG_UsuarioEmpresa AS ue
            WHERE ue.AspNetUserId = @AspNetUserId
              AND ue.IdEmpresa = @IdEmpresa
        )
        BEGIN
            UPDATE dbo.SEG_UsuarioEmpresa
            SET Estado = 1,
                EsEmpresaPredeterminada = @EsEmpresaPredeterminada,
                UsuarioRegistro = @UsuarioRegistro
            WHERE AspNetUserId = @AspNetUserId
              AND IdEmpresa = @IdEmpresa;
        END
        ELSE
        BEGIN
            INSERT INTO dbo.SEG_UsuarioEmpresa
            (
                AspNetUserId,
                IdEmpresa,
                EsEmpresaPredeterminada,
                Estado,
                UsuarioRegistro
            )
            VALUES
            (
                @AspNetUserId,
                @IdEmpresa,
                @EsEmpresaPredeterminada,
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
