-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   25/08/2026
-- Description:   Migra las cuentas maestras de documentos e impuestos desde IdPlanCuenta hacia CodigoCuenta portable entre empresas.
-- =============================================

SET XACT_ABORT ON;

BEGIN TRY

BEGIN TRANSACTION;

IF COL_LENGTH(N'dbo.ADM_TipoComprobante', N'CodigoCuentaVentaSoles') IS NULL
BEGIN
    ALTER TABLE dbo.ADM_TipoComprobante ADD CodigoCuentaVentaSoles VARCHAR(20) NULL;
END;

IF COL_LENGTH(N'dbo.ADM_TipoComprobante', N'CodigoCuentaVentaDolares') IS NULL
BEGIN
    ALTER TABLE dbo.ADM_TipoComprobante ADD CodigoCuentaVentaDolares VARCHAR(20) NULL;
END;

IF COL_LENGTH(N'dbo.ADM_TipoComprobante', N'CodigoCuentaCompraSoles') IS NULL
BEGIN
    ALTER TABLE dbo.ADM_TipoComprobante ADD CodigoCuentaCompraSoles VARCHAR(20) NULL;
END;

IF COL_LENGTH(N'dbo.ADM_TipoComprobante', N'CodigoCuentaCompraDolares') IS NULL
BEGIN
    ALTER TABLE dbo.ADM_TipoComprobante ADD CodigoCuentaCompraDolares VARCHAR(20) NULL;
END;

IF COL_LENGTH(N'dbo.ADM_TipoComprobante', N'IdCuentaVentaSoles') IS NOT NULL
BEGIN
    EXEC sys.sp_executesql N'
        UPDATE tipo
        SET CodigoCuentaVentaSoles = COALESCE(tipo.CodigoCuentaVentaSoles, cuenta.CodigoCuenta)
        FROM dbo.ADM_TipoComprobante AS tipo
        LEFT JOIN dbo.CON_PlanCuenta AS cuenta
            ON cuenta.IdPlanCuenta = tipo.IdCuentaVentaSoles;';

    EXEC sys.sp_executesql N'
        IF EXISTS
        (
            SELECT 1
            FROM dbo.ADM_TipoComprobante AS tipo
            WHERE tipo.IdCuentaVentaSoles IS NOT NULL
              AND NULLIF(LTRIM(RTRIM(tipo.CodigoCuentaVentaSoles)), '''') IS NULL
        )
        BEGIN
            RAISERROR(N''No se pudo convertir una cuenta maestra de venta en soles a CodigoCuenta.'', 16, 1);
        END;';

    EXEC sys.sp_executesql N'ALTER TABLE dbo.ADM_TipoComprobante DROP COLUMN IdCuentaVentaSoles;';
END;

IF COL_LENGTH(N'dbo.ADM_TipoComprobante', N'IdCuentaVentaDolares') IS NOT NULL
BEGIN
    EXEC sys.sp_executesql N'
        UPDATE tipo
        SET CodigoCuentaVentaDolares = COALESCE(tipo.CodigoCuentaVentaDolares, cuenta.CodigoCuenta)
        FROM dbo.ADM_TipoComprobante AS tipo
        LEFT JOIN dbo.CON_PlanCuenta AS cuenta
            ON cuenta.IdPlanCuenta = tipo.IdCuentaVentaDolares;';

    EXEC sys.sp_executesql N'
        IF EXISTS
        (
            SELECT 1
            FROM dbo.ADM_TipoComprobante AS tipo
            WHERE tipo.IdCuentaVentaDolares IS NOT NULL
              AND NULLIF(LTRIM(RTRIM(tipo.CodigoCuentaVentaDolares)), '''') IS NULL
        )
        BEGIN
            RAISERROR(N''No se pudo convertir una cuenta maestra de venta en dolares a CodigoCuenta.'', 16, 1);
        END;';

    EXEC sys.sp_executesql N'ALTER TABLE dbo.ADM_TipoComprobante DROP COLUMN IdCuentaVentaDolares;';
END;

IF COL_LENGTH(N'dbo.ADM_TipoComprobante', N'IdCuentaCompraSoles') IS NOT NULL
BEGIN
    EXEC sys.sp_executesql N'
        UPDATE tipo
        SET CodigoCuentaCompraSoles = COALESCE(tipo.CodigoCuentaCompraSoles, cuenta.CodigoCuenta)
        FROM dbo.ADM_TipoComprobante AS tipo
        LEFT JOIN dbo.CON_PlanCuenta AS cuenta
            ON cuenta.IdPlanCuenta = tipo.IdCuentaCompraSoles;';

    EXEC sys.sp_executesql N'
        IF EXISTS
        (
            SELECT 1
            FROM dbo.ADM_TipoComprobante AS tipo
            WHERE tipo.IdCuentaCompraSoles IS NOT NULL
              AND NULLIF(LTRIM(RTRIM(tipo.CodigoCuentaCompraSoles)), '''') IS NULL
        )
        BEGIN
            RAISERROR(N''No se pudo convertir una cuenta maestra de compra en soles a CodigoCuenta.'', 16, 1);
        END;';

    EXEC sys.sp_executesql N'ALTER TABLE dbo.ADM_TipoComprobante DROP COLUMN IdCuentaCompraSoles;';
END;

IF COL_LENGTH(N'dbo.ADM_TipoComprobante', N'IdCuentaCompraDolares') IS NOT NULL
BEGIN
    EXEC sys.sp_executesql N'
        UPDATE tipo
        SET CodigoCuentaCompraDolares = COALESCE(tipo.CodigoCuentaCompraDolares, cuenta.CodigoCuenta)
        FROM dbo.ADM_TipoComprobante AS tipo
        LEFT JOIN dbo.CON_PlanCuenta AS cuenta
            ON cuenta.IdPlanCuenta = tipo.IdCuentaCompraDolares;';

    EXEC sys.sp_executesql N'
        IF EXISTS
        (
            SELECT 1
            FROM dbo.ADM_TipoComprobante AS tipo
            WHERE tipo.IdCuentaCompraDolares IS NOT NULL
              AND NULLIF(LTRIM(RTRIM(tipo.CodigoCuentaCompraDolares)), '''') IS NULL
        )
        BEGIN
            RAISERROR(N''No se pudo convertir una cuenta maestra de compra en dolares a CodigoCuenta.'', 16, 1);
        END;';

    EXEC sys.sp_executesql N'ALTER TABLE dbo.ADM_TipoComprobante DROP COLUMN IdCuentaCompraDolares;';
END;

IF COL_LENGTH(N'dbo.CON_TipoImpuesto', N'CodigoCuenta') IS NULL
BEGIN
    ALTER TABLE dbo.CON_TipoImpuesto ADD CodigoCuenta VARCHAR(20) NULL;
END;

IF COL_LENGTH(N'dbo.CON_TipoImpuesto', N'IdPlanCuenta') IS NOT NULL
BEGIN
    EXEC sys.sp_executesql N'
        UPDATE impuesto
        SET CodigoCuenta = COALESCE(impuesto.CodigoCuenta, cuenta.CodigoCuenta)
        FROM dbo.CON_TipoImpuesto AS impuesto
        LEFT JOIN dbo.CON_PlanCuenta AS cuenta
            ON cuenta.IdPlanCuenta = impuesto.IdPlanCuenta;';

    EXEC sys.sp_executesql N'
        IF EXISTS
        (
            SELECT 1
            FROM dbo.CON_TipoImpuesto AS impuesto
            WHERE impuesto.IdPlanCuenta IS NOT NULL
              AND NULLIF(LTRIM(RTRIM(impuesto.CodigoCuenta)), '''') IS NULL
        )
        BEGIN
            RAISERROR(N''No se pudo convertir una cuenta maestra de impuesto a CodigoCuenta.'', 16, 1);
        END;';

    EXEC sys.sp_executesql N'ALTER TABLE dbo.CON_TipoImpuesto DROP COLUMN IdPlanCuenta;';
END;

COMMIT TRANSACTION;

END TRY

BEGIN CATCH

    IF XACT_STATE() <> 0
    BEGIN
        ROLLBACK TRANSACTION;
    END;

    DECLARE @ErrorMessage NVARCHAR(4000)
    DECLARE @ErrorSeverity INT
    DECLARE @ErrorState INT

    SELECT
        @ErrorMessage = ERROR_MESSAGE(),
        @ErrorSeverity = ERROR_SEVERITY(),
        @ErrorState = ERROR_STATE()

    RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)

END CATCH;
