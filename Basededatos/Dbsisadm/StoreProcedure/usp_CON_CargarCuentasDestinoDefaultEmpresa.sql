-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Copia reglas maestras internas de cuentas destino hacia una empresa.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_CargarCuentasDestinoDefaultEmpresa
    @IdEmpresa INT,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        INSERT INTO dbo.CON_CuentaDestinoRegla
        (
            IdEmpresa,
            Ejercicio,
            IdPlanCuentaOrigen,
            Activo,
            Observacion,
            UsuarioRegistro
        )
        SELECT
            @IdEmpresa,
            rm.Ejercicio,
            origen.IdPlanCuenta,
            rm.Activo,
            rm.Observacion,
            @UsuarioRegistro
        FROM dbo.CON_CuentaDestinoReglaMaestro AS rm
        INNER JOIN dbo.CON_PlanCuenta AS origen
            ON origen.IdEmpresa = @IdEmpresa
           AND origen.CodigoCuenta = rm.CodigoCuentaOrigen
        WHERE rm.Activo = 1
          AND NOT EXISTS
          (
              SELECT 1
              FROM dbo.CON_CuentaDestinoRegla AS r
              WHERE r.IdEmpresa = @IdEmpresa
                AND r.Ejercicio = rm.Ejercicio
                AND r.IdPlanCuentaOrigen = origen.IdPlanCuenta
          );

        INSERT INTO dbo.CON_CuentaDestinoReglaDetalle
        (
            IdCuentaDestinoRegla,
            Orden,
            IdPlanCuentaDestinoCargo,
            IdPlanCuentaDestinoAbono,
            Porcentaje,
            Activo,
            UsuarioRegistro
        )
        SELECT
            r.IdCuentaDestinoRegla,
            dm.Orden,
            cargo.IdPlanCuenta,
            abono.IdPlanCuenta,
            dm.Porcentaje,
            dm.Activo,
            @UsuarioRegistro
        FROM dbo.CON_CuentaDestinoReglaDetalleMaestro AS dm
        INNER JOIN dbo.CON_CuentaDestinoReglaMaestro AS rm
            ON rm.IdCuentaDestinoReglaMaestro = dm.IdCuentaDestinoReglaMaestro
        INNER JOIN dbo.CON_PlanCuenta AS origen
            ON origen.IdEmpresa = @IdEmpresa
           AND origen.CodigoCuenta = rm.CodigoCuentaOrigen
        INNER JOIN dbo.CON_CuentaDestinoRegla AS r
            ON r.IdEmpresa = @IdEmpresa
           AND r.Ejercicio = rm.Ejercicio
           AND r.IdPlanCuentaOrigen = origen.IdPlanCuenta
        INNER JOIN dbo.CON_PlanCuenta AS cargo
            ON cargo.IdEmpresa = @IdEmpresa
           AND cargo.CodigoCuenta = dm.CodigoCuentaDestinoCargo
        INNER JOIN dbo.CON_PlanCuenta AS abono
            ON abono.IdEmpresa = @IdEmpresa
           AND abono.CodigoCuenta = dm.CodigoCuentaDestinoAbono
        WHERE dm.Activo = 1
          AND NOT EXISTS
          (
              SELECT 1
              FROM dbo.CON_CuentaDestinoReglaDetalle AS d
              WHERE d.IdCuentaDestinoRegla = r.IdCuentaDestinoRegla
                AND d.Orden = dm.Orden
          );

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
