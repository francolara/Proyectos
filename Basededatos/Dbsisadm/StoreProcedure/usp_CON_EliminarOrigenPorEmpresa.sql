-- =============================================
-- Author:        FRANCO LARA
-- Create date:   29/08/2026
-- Description:   Elimina un origen de empresa solo cuando no tiene configuraciones, correlativos, asientos o procesos relacionados.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.usp_CON_EliminarOrigenPorEmpresa
    @IdEmpresa INT,
    @IdOrigen INT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.CON_Origen AS o
            WHERE o.IdOrigen = @IdOrigen
              AND o.IdEmpresa = @IdEmpresa
        )
        BEGIN
            RAISERROR (N'El origen no existe en la empresa activa.', 16, 1);
        END;

        DELETE o
        FROM dbo.CON_Origen AS o
        WHERE o.IdOrigen = @IdOrigen
          AND o.IdEmpresa = @IdEmpresa;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        IF ERROR_NUMBER() = 547
        BEGIN
            SET @ErrorMessage = N'No se puede eliminar el origen porque tiene configuraciones, correlativos, asientos o procesos relacionados.';
            SET @ErrorSeverity = 16;
            SET @ErrorState = 1;
        END
        ELSE
        BEGIN
            SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        END;

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH;
END;
