-- =============================================
-- Author:        FRANCO LARA
-- Create date:   29/08/2026
-- Description:   Elimina una cuenta de empresa solo cuando no tiene dependencias contables u operativas.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.usp_CON_EliminarPlanCuentaPorEmpresa
    @IdEmpresa INT,
    @IdPlanCuenta INT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.CON_PlanCuenta AS pc
            WHERE pc.IdPlanCuenta = @IdPlanCuenta
              AND pc.IdEmpresa = @IdEmpresa
        )
        BEGIN
            RAISERROR (N'La cuenta contable no existe en la empresa activa.', 16, 1);
        END;

        DELETE pc
        FROM dbo.CON_PlanCuenta AS pc
        WHERE pc.IdPlanCuenta = @IdPlanCuenta
          AND pc.IdEmpresa = @IdEmpresa;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        IF ERROR_NUMBER() = 547
        BEGIN
            SET @ErrorMessage = N'No se puede eliminar la cuenta contable porque tiene cuentas hijas, configuraciones o movimientos relacionados.';
            SET @ErrorSeverity = 16;
            SET @ErrorState = 1;
        END
        ELSE
        BEGIN
            SELECT
                @ErrorMessage = ERROR_MESSAGE(),
                @ErrorSeverity = ERROR_SEVERITY(),
                @ErrorState = ERROR_STATE();
        END;

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH;
END;
