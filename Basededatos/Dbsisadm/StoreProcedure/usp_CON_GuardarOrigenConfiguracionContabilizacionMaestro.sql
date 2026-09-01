-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   27/08/2026
-- Description:   Actualiza solamente el origen asignado a una configuracion contable maestra.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_GuardarOrigenConfiguracionContabilizacionMaestro
    @IdConfiguracionContabilizacionMaestro INT,
    @CodigoOrigen VARCHAR(10),
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @Codigo VARCHAR(10) = UPPER(NULLIF(LTRIM(RTRIM(@CodigoOrigen)), ''));

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.CON_ConfiguracionContabilizacionMaestro
            WHERE IdConfiguracionContabilizacionMaestro = @IdConfiguracionContabilizacionMaestro
        )
            RAISERROR(N'La configuracion contable maestra indicada no existe.', 16, 1);

        IF @Codigo IS NULL
           OR NOT EXISTS
           (
               SELECT 1
               FROM dbo.CON_OrigenMaestro
               WHERE CodigoOrigen = @Codigo AND Estado = 1
           )
            RAISERROR(N'El origen seleccionado no existe o esta inactivo.', 16, 1);

        UPDATE dbo.CON_ConfiguracionContabilizacionMaestro
        SET CodigoOrigen = @Codigo,
            UsuarioRegistro = @UsuarioRegistro
        WHERE IdConfiguracionContabilizacionMaestro = @IdConfiguracionContabilizacionMaestro;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE()
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)
    END CATCH
END
