-- =============================================
-- Author:        FRANCO LARA
-- Create date:   29/08/2026
-- Description:   Elimina una cuenta corriente solo cuando no tiene movimientos bancarios relacionados.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.usp_CON_EliminarBancoConfiguracionEmpresa
    @IdEmpresa INT,
    @IdBancoConfiguracionEmpresa INT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.CON_BancosConfiguracionEmpresa AS bce
            WHERE bce.IdBancoConfiguracionEmpresa = @IdBancoConfiguracionEmpresa
              AND bce.IdEmpresa = @IdEmpresa
        )
        BEGIN
            RAISERROR (N'La cuenta corriente no existe en la empresa activa.', 16, 1);
        END;

        DELETE bce
        FROM dbo.CON_BancosConfiguracionEmpresa AS bce
        WHERE bce.IdBancoConfiguracionEmpresa = @IdBancoConfiguracionEmpresa
          AND bce.IdEmpresa = @IdEmpresa;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        IF ERROR_NUMBER() = 547
        BEGIN
            SET @ErrorMessage = N'No se puede eliminar la cuenta corriente porque tiene movimientos bancarios relacionados.';
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
