-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   25/08/2026
-- Description:   Inicializa en una transaccion el plan, parametros, cuentas destino, impuestos y documentos de una empresa.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_CargarConfiguracionDefaultEmpresa
    @IdEmpresa INT,
    @IdEmpresaBase INT = NULL,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    BEGIN TRY

        BEGIN TRANSACTION;

        EXEC dbo.usp_CON_CargarPlanCuentaDefaultEmpresa
            @IdEmpresa = @IdEmpresa,
            @IdEmpresaBase = @IdEmpresaBase,
            @UsuarioRegistro = @UsuarioRegistro;

        EXEC dbo.usp_ADM_CargarParametrosDefaultEmpresa
            @IdEmpresa = @IdEmpresa,
            @UsuarioRegistro = @UsuarioRegistro;

        EXEC dbo.usp_CON_CargarCuentasDestinoDefaultEmpresa
            @IdEmpresa = @IdEmpresa,
            @UsuarioRegistro = @UsuarioRegistro;

        EXEC dbo.usp_CON_CargarImpuestosDefaultEmpresa
            @IdEmpresa = @IdEmpresa,
            @UsuarioRegistro = @UsuarioRegistro;

        EXEC dbo.usp_CON_CargarDocumentosDefaultEmpresa
            @IdEmpresa = @IdEmpresa,
            @UsuarioRegistro = @UsuarioRegistro;

        COMMIT TRANSACTION;

    END TRY

    BEGIN CATCH

        IF XACT_STATE() <> 0
        BEGIN
            ROLLBACK TRANSACTION;
        END;

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
