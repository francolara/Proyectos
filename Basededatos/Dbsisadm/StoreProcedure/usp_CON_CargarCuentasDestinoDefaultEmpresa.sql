-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Copia reglas maestras internas de cuentas destino hacia una empresa.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   25/06/2026
-- Description:   Carga una sola configuracion de cuentas destino por empresa y cuenta origen, sin depender de un ejercicio.
-- =============================================
-- Firma: FRANCO LARA - 25/08/2026 | Valida los codigos maestros y carga reglas por empresa sin dependencia del ejercicio.

CREATE OR ALTER PROCEDURE dbo.usp_CON_CargarCuentasDestinoDefaultEmpresa
    @IdEmpresa INT,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @CodigoCuentaFaltante VARCHAR(20)

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.SEG_Empresa AS empresa
            WHERE empresa.IdEmpresa = @IdEmpresa
        )
        BEGIN
            RAISERROR(N'La empresa indicada no existe.', 16, 1);
        END;

        SELECT TOP (1)
            @CodigoCuentaFaltante = regla.CodigoCuentaOrigen
        FROM dbo.CON_CuentaDestinoReglaMaestro AS regla
        LEFT JOIN dbo.CON_PlanCuenta AS cuenta
            ON cuenta.IdEmpresa = @IdEmpresa
           AND cuenta.CodigoCuenta = regla.CodigoCuentaOrigen
           AND cuenta.Estado = 1
           AND cuenta.AceptaMovimiento = 1
        WHERE regla.Activo = 1
          AND cuenta.IdPlanCuenta IS NULL
        ORDER BY regla.CodigoCuentaOrigen;

        IF @CodigoCuentaFaltante IS NOT NULL
        BEGIN
            RAISERROR(N'La cuenta origen maestra %s no existe, esta inactiva o no acepta movimiento en el plan de la empresa.', 16, 1, @CodigoCuentaFaltante);
        END;

        SET @CodigoCuentaFaltante = NULL;

        SELECT TOP (1)
            @CodigoCuentaFaltante = codigos.CodigoCuenta
        FROM dbo.CON_CuentaDestinoReglaDetalleMaestro AS detalle
        INNER JOIN dbo.CON_CuentaDestinoReglaMaestro AS regla
            ON regla.IdCuentaDestinoReglaMaestro = detalle.IdCuentaDestinoReglaMaestro
        CROSS APPLY
        (
            VALUES
                (detalle.CodigoCuentaDestinoCargo),
                (detalle.CodigoCuentaDestinoAbono)
        ) AS codigos (CodigoCuenta)
        LEFT JOIN dbo.CON_PlanCuenta AS cuenta
            ON cuenta.IdEmpresa = @IdEmpresa
           AND cuenta.CodigoCuenta = codigos.CodigoCuenta
           AND cuenta.Estado = 1
           AND cuenta.AceptaMovimiento = 1
        WHERE regla.Activo = 1
          AND detalle.Activo = 1
          AND cuenta.IdPlanCuenta IS NULL
        ORDER BY codigos.CodigoCuenta;

        IF @CodigoCuentaFaltante IS NOT NULL
        BEGIN
            RAISERROR(N'La cuenta destino maestra %s no existe, esta inactiva o no acepta movimiento en el plan de la empresa.', 16, 1, @CodigoCuentaFaltante);
        END;

        INSERT INTO dbo.CON_CuentaDestinoRegla
        (
            IdEmpresa,
            IdPlanCuentaOrigen,
            Activo,
            Observacion,
            UsuarioRegistro
        )
        SELECT
            @IdEmpresa,
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
