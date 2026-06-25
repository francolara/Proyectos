-- =============================================
-- Author:        FRANCO LARA
-- Create date:   22/06/2026
-- Description:   Elimina un asiento manual y bloquea la eliminacion directa de asientos automaticos generados desde su modulo de origen.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_EliminarAsiento
    @IdAsiento INT,
    @IdEmpresa INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @PermiteRegistroManual BIT

        SELECT
            @PermiteRegistroManual = o.PermiteRegistroManual
        FROM dbo.CON_Asiento AS a
        INNER JOIN dbo.CON_Origen AS o
            ON o.IdOrigen = a.IdOrigen
        WHERE a.IdAsiento = @IdAsiento
          AND a.IdEmpresa = @IdEmpresa;

        IF @PermiteRegistroManual IS NULL
        BEGIN
            RAISERROR(N'El asiento indicado no existe para la empresa activa.', 16, 1);
        END;

        IF @PermiteRegistroManual = 0
        BEGIN
            RAISERROR(N'el asiento fue generado de forma automática , elimínelo desde el módulo de origen.', 16, 1);
        END;

        BEGIN TRAN;

        DELETE FROM dbo.CON_AsientoDetalle
        WHERE IdAsiento = @IdAsiento;

        DELETE FROM dbo.CON_Asiento
        WHERE IdAsiento = @IdAsiento
          AND IdEmpresa = @IdEmpresa;

        COMMIT TRAN;

    END TRY

    BEGIN CATCH

        IF @@TRANCOUNT > 0
        BEGIN
            ROLLBACK TRAN;
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
