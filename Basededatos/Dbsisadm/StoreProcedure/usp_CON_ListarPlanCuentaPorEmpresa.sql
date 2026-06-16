-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Lista el plan de cuentas de una empresa con filtro opcional para cuentas de movimiento.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_ListarPlanCuentaPorEmpresa
    @IdEmpresa INT,
    @SoloMovimiento BIT = 0
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            pc.IdPlanCuenta,
            pc.IdPlanCuentaPadre,
            pc.CodigoCuenta,
            pc.NombreCuenta,
            pc.NivelCuenta,
            pc.NaturalezaSaldo,
            pc.AceptaMovimiento,
            pc.RequiereCentroCosto
        FROM dbo.CON_PlanCuenta AS pc
        WHERE pc.IdEmpresa = @IdEmpresa
          AND pc.Estado = 1
          AND (@SoloMovimiento = 0 OR pc.AceptaMovimiento = 1)
        ORDER BY pc.CodigoCuenta ASC;

    END TRY

    BEGIN CATCH

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
