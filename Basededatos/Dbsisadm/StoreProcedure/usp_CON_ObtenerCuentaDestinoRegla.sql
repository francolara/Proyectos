-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Obtiene la cabecera y detalle de una regla de cuentas destino.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   25/06/2026
-- Description:   Simplifica la cabecera de cuentas destino para operar sin dependencia funcional del ejercicio.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_ObtenerCuentaDestinoRegla
    @IdCuentaDestinoRegla INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            r.IdCuentaDestinoRegla,
            r.IdEmpresa,
            r.IdPlanCuentaOrigen,
            po.CodigoCuenta AS CodigoCuentaOrigen,
            po.NombreCuenta AS NombreCuentaOrigen,
            r.Activo,
            r.Observacion
        FROM dbo.CON_CuentaDestinoRegla AS r
        INNER JOIN dbo.CON_PlanCuenta AS po
            ON po.IdPlanCuenta = r.IdPlanCuentaOrigen
        WHERE r.IdCuentaDestinoRegla = @IdCuentaDestinoRegla;

        SELECT
            d.IdCuentaDestinoReglaDetalle,
            d.IdCuentaDestinoRegla,
            d.Orden,
            d.IdPlanCuentaDestinoCargo,
            pc1.CodigoCuenta AS CodigoCuentaDestinoCargo,
            pc1.NombreCuenta AS NombreCuentaDestinoCargo,
            d.IdPlanCuentaDestinoAbono,
            pc2.CodigoCuenta AS CodigoCuentaDestinoAbono,
            pc2.NombreCuenta AS NombreCuentaDestinoAbono,
            d.Porcentaje,
            d.Activo
        FROM dbo.CON_CuentaDestinoReglaDetalle AS d
        INNER JOIN dbo.CON_PlanCuenta AS pc1
            ON pc1.IdPlanCuenta = d.IdPlanCuentaDestinoCargo
        INNER JOIN dbo.CON_PlanCuenta AS pc2
            ON pc2.IdPlanCuenta = d.IdPlanCuentaDestinoAbono
        WHERE d.IdCuentaDestinoRegla = @IdCuentaDestinoRegla
        ORDER BY
            d.Orden ASC;

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
