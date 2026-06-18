-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Elimina la configuracion contable automatica y su detalle.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_EliminarConfiguracionContabilizacion
    @IdConfiguracionContabilizacion INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        BEGIN TRAN;

        DELETE FROM dbo.CON_ConfiguracionContabilizacionDetalle
        WHERE IdConfiguracionContabilizacion = @IdConfiguracionContabilizacion;

        DELETE FROM dbo.CON_ConfiguracionContabilizacion
        WHERE IdConfiguracionContabilizacion = @IdConfiguracionContabilizacion;

        COMMIT;

    END TRY

    BEGIN CATCH

        IF @@TRANCOUNT > 0
        BEGIN
            ROLLBACK;
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
