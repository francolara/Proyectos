-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Lista las reglas de cuentas de destino por empresa y ejercicio con resumen de porcentajes.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_ListarCuentasDestinoReglaPorEmpresa
    @IdEmpresa INT,
    @Ejercicio SMALLINT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            r.IdCuentaDestinoRegla,
            r.IdEmpresa,
            r.Ejercicio,
            r.IdPlanCuentaOrigen,
            po.CodigoCuenta AS CodigoCuentaOrigen,
            po.NombreCuenta AS NombreCuentaOrigen,
            r.Activo,
            r.Observacion,
            COUNT(d.IdCuentaDestinoReglaDetalle) AS CantidadTramos,
            COALESCE(SUM(d.Porcentaje), 0) AS PorcentajeTotal
        FROM dbo.CON_CuentaDestinoRegla AS r
        INNER JOIN dbo.CON_PlanCuenta AS po
            ON po.IdPlanCuenta = r.IdPlanCuentaOrigen
        LEFT JOIN dbo.CON_CuentaDestinoReglaDetalle AS d
            ON d.IdCuentaDestinoRegla = r.IdCuentaDestinoRegla
           AND d.Activo = 1
        WHERE r.IdEmpresa = @IdEmpresa
          AND r.Ejercicio = @Ejercicio
        GROUP BY
            r.IdCuentaDestinoRegla,
            r.IdEmpresa,
            r.Ejercicio,
            r.IdPlanCuentaOrigen,
            po.CodigoCuenta,
            po.NombreCuenta,
            r.Activo,
            r.Observacion
        ORDER BY
            po.CodigoCuenta ASC;

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
