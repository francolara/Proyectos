-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Elimina una regla de cuentas destino y su detalle.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_EliminarCuentaDestinoRegla
    @IdCuentaDestinoRegla INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        BEGIN TRAN;

        DELETE FROM dbo.CON_CuentaDestinoReglaDetalle
        WHERE IdCuentaDestinoRegla = @IdCuentaDestinoRegla;

        DELETE FROM dbo.CON_CuentaDestinoRegla
        WHERE IdCuentaDestinoRegla = @IdCuentaDestinoRegla;

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
