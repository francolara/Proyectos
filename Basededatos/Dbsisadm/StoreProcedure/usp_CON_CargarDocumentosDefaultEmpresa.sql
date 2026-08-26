-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   25/08/2026
-- Description:   Carga por empresa las cuentas de documentos resolviendo los codigos maestros contra el plan contable empresarial.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_CargarDocumentosDefaultEmpresa
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
            @CodigoCuentaFaltante = codigos.CodigoCuenta
        FROM dbo.ADM_TipoComprobante AS tipo
        CROSS APPLY
        (
            VALUES
                (tipo.CodigoCuentaVentaSoles),
                (tipo.CodigoCuentaVentaDolares),
                (tipo.CodigoCuentaCompraSoles),
                (tipo.CodigoCuentaCompraDolares)
        ) AS codigos (CodigoCuenta)
        LEFT JOIN dbo.CON_PlanCuenta AS cuenta
            ON cuenta.IdEmpresa = @IdEmpresa
           AND cuenta.CodigoCuenta = codigos.CodigoCuenta
           AND cuenta.Estado = 1
           AND cuenta.AceptaMovimiento = 1
        WHERE tipo.Estado = 1
          AND NULLIF(LTRIM(RTRIM(codigos.CodigoCuenta)), '') IS NOT NULL
          AND cuenta.IdPlanCuenta IS NULL
        ORDER BY codigos.CodigoCuenta;

        IF @CodigoCuentaFaltante IS NOT NULL
        BEGIN
            RAISERROR(N'La cuenta maestra de documento %s no existe, esta inactiva o no acepta movimiento en el plan de la empresa.', 16, 1, @CodigoCuentaFaltante);
        END;

        INSERT INTO dbo.CON_DocumentoConfiguracionEmpresa
        (
            IdEmpresa,
            IdTipoComprobante,
            IdCuentaVentaSoles,
            IdCuentaVentaDolares,
            IdCuentaCompraSoles,
            IdCuentaCompraDolares,
            Activo,
            UsuarioRegistro
        )
        SELECT
            @IdEmpresa,
            tipo.IdTipoComprobante,
            cuentaVentaSoles.IdPlanCuenta,
            cuentaVentaDolares.IdPlanCuenta,
            cuentaCompraSoles.IdPlanCuenta,
            cuentaCompraDolares.IdPlanCuenta,
            1,
            @UsuarioRegistro
        FROM dbo.ADM_TipoComprobante AS tipo
        LEFT JOIN dbo.CON_PlanCuenta AS cuentaVentaSoles
            ON cuentaVentaSoles.IdEmpresa = @IdEmpresa
           AND cuentaVentaSoles.CodigoCuenta = tipo.CodigoCuentaVentaSoles
           AND cuentaVentaSoles.Estado = 1
           AND cuentaVentaSoles.AceptaMovimiento = 1
        LEFT JOIN dbo.CON_PlanCuenta AS cuentaVentaDolares
            ON cuentaVentaDolares.IdEmpresa = @IdEmpresa
           AND cuentaVentaDolares.CodigoCuenta = tipo.CodigoCuentaVentaDolares
           AND cuentaVentaDolares.Estado = 1
           AND cuentaVentaDolares.AceptaMovimiento = 1
        LEFT JOIN dbo.CON_PlanCuenta AS cuentaCompraSoles
            ON cuentaCompraSoles.IdEmpresa = @IdEmpresa
           AND cuentaCompraSoles.CodigoCuenta = tipo.CodigoCuentaCompraSoles
           AND cuentaCompraSoles.Estado = 1
           AND cuentaCompraSoles.AceptaMovimiento = 1
        LEFT JOIN dbo.CON_PlanCuenta AS cuentaCompraDolares
            ON cuentaCompraDolares.IdEmpresa = @IdEmpresa
           AND cuentaCompraDolares.CodigoCuenta = tipo.CodigoCuentaCompraDolares
           AND cuentaCompraDolares.Estado = 1
           AND cuentaCompraDolares.AceptaMovimiento = 1
        WHERE tipo.Estado = 1
          AND NOT EXISTS
          (
              SELECT 1
              FROM dbo.CON_DocumentoConfiguracionEmpresa AS configuracion
              WHERE configuracion.IdEmpresa = @IdEmpresa
                AND configuracion.IdTipoComprobante = tipo.IdTipoComprobante
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
