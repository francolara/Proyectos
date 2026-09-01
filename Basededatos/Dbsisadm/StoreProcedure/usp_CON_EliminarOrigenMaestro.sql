-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   27/08/2026
-- Description:   Elimina un origen maestro solamente cuando no esta asignado a una configuracion maestra.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_EliminarOrigenMaestro
    @IdOrigenMaestro INT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @CodigoOrigen VARCHAR(10);

        SELECT @CodigoOrigen = origen.CodigoOrigen
        FROM dbo.CON_OrigenMaestro AS origen
        WHERE origen.IdOrigenMaestro = @IdOrigenMaestro;

        IF @CodigoOrigen IS NULL
            RAISERROR(N'El origen maestro indicado no existe.', 16, 1);

        IF EXISTS
        (
            SELECT 1
            FROM dbo.CON_ConfiguracionContabilizacionMaestro AS configuracion
            WHERE configuracion.CodigoOrigen = @CodigoOrigen
        )
            RAISERROR(N'No se puede eliminar el origen porque esta asignado a una configuracion contable maestra.', 16, 1);

        DELETE FROM dbo.CON_OrigenMaestro
        WHERE IdOrigenMaestro = @IdOrigenMaestro;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE()
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)
    END CATCH
END
