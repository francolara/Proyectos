-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   27/08/2026
-- Description:   Elimina de forma transaccional una regla maestra y todos sus tramos.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_EliminarCuentaDestinoMaestro
    @IdCuentaDestinoReglaMaestro INT
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.CON_CuentaDestinoReglaMaestro WHERE IdCuentaDestinoReglaMaestro = @IdCuentaDestinoReglaMaestro)
            RAISERROR(N'La regla maestra indicada no existe.', 16, 1);

        BEGIN TRANSACTION;
        DELETE FROM dbo.CON_CuentaDestinoReglaDetalleMaestro WHERE IdCuentaDestinoReglaMaestro = @IdCuentaDestinoReglaMaestro;
        DELETE FROM dbo.CON_CuentaDestinoReglaMaestro WHERE IdCuentaDestinoReglaMaestro = @IdCuentaDestinoReglaMaestro;
        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE()
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)
    END CATCH
END
