-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   25/08/2026
-- Description:   Carga por empresa las cuentas de impuestos resolviendo el CodigoCuenta del maestro contra el plan contable empresarial.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_CargarImpuestosDefaultEmpresa
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
            @CodigoCuentaFaltante = impuesto.CodigoCuenta
        FROM dbo.CON_TipoImpuesto AS impuesto
        LEFT JOIN dbo.CON_PlanCuenta AS cuenta
            ON cuenta.IdEmpresa = @IdEmpresa
           AND cuenta.CodigoCuenta = impuesto.CodigoCuenta
           AND cuenta.Estado = 1
           AND cuenta.AceptaMovimiento = 1
        WHERE impuesto.Estado = 1
          AND NULLIF(LTRIM(RTRIM(impuesto.CodigoCuenta)), '') IS NOT NULL
          AND cuenta.IdPlanCuenta IS NULL
        ORDER BY impuesto.CodigoSunat;

        IF @CodigoCuentaFaltante IS NOT NULL
        BEGIN
            RAISERROR(N'La cuenta maestra de impuesto %s no existe, esta inactiva o no acepta movimiento en el plan de la empresa.', 16, 1, @CodigoCuentaFaltante);
        END;

        INSERT INTO dbo.CON_TipoImpuestoConfiguracionEmpresa
        (
            IdEmpresa,
            IdTipoImpuesto,
            IdPlanCuenta,
            Activo,
            UsuarioRegistro
        )
        SELECT
            @IdEmpresa,
            impuesto.IdTipoImpuesto,
            cuenta.IdPlanCuenta,
            1,
            @UsuarioRegistro
        FROM dbo.CON_TipoImpuesto AS impuesto
        LEFT JOIN dbo.CON_PlanCuenta AS cuenta
            ON cuenta.IdEmpresa = @IdEmpresa
           AND cuenta.CodigoCuenta = impuesto.CodigoCuenta
           AND cuenta.Estado = 1
           AND cuenta.AceptaMovimiento = 1
        WHERE impuesto.Estado = 1
          AND NOT EXISTS
          (
              SELECT 1
              FROM dbo.CON_TipoImpuestoConfiguracionEmpresa AS configuracion
              WHERE configuracion.IdEmpresa = @IdEmpresa
                AND configuracion.IdTipoImpuesto = impuesto.IdTipoImpuesto
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
