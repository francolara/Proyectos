-- =============================================
-- Author:        FRANCO LARA
-- Create date:   20/06/2026
-- Description:   Migra configuracion de documentos a cuentas separadas de compra y venta por moneda.
-- =============================================
-- Firma: FRANCO LARA - 25/08/2026 | Define las cuentas maestras de comprobantes como codigos, conserva IdPlanCuenta por empresa y migra columnas antiguas mediante SQL dinamico.

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

IF COL_LENGTH(N'dbo.CON_DocumentoConfiguracionEmpresa', N'IdCuentaVentaSoles') IS NULL
BEGIN
    ALTER TABLE dbo.CON_DocumentoConfiguracionEmpresa ADD IdCuentaVentaSoles INT NULL;
END;

IF COL_LENGTH(N'dbo.CON_DocumentoConfiguracionEmpresa', N'IdCuentaVentaDolares') IS NULL
BEGIN
    ALTER TABLE dbo.CON_DocumentoConfiguracionEmpresa ADD IdCuentaVentaDolares INT NULL;
END;

IF COL_LENGTH(N'dbo.CON_DocumentoConfiguracionEmpresa', N'IdCuentaCompraSoles') IS NULL
BEGIN
    ALTER TABLE dbo.CON_DocumentoConfiguracionEmpresa ADD IdCuentaCompraSoles INT NULL;
END;

IF COL_LENGTH(N'dbo.CON_DocumentoConfiguracionEmpresa', N'IdCuentaCompraDolares') IS NULL
BEGIN
    ALTER TABLE dbo.CON_DocumentoConfiguracionEmpresa ADD IdCuentaCompraDolares INT NULL;
END;

IF COL_LENGTH(N'dbo.CON_DocumentoConfiguracionEmpresa', N'IdPlanCuentaSoles') IS NOT NULL
BEGIN
    EXEC sys.sp_executesql N'
        UPDATE cfg
        SET IdCuentaVentaSoles = ISNULL(cfg.IdCuentaVentaSoles, cfg.IdPlanCuentaSoles),
            IdCuentaCompraSoles = ISNULL(cfg.IdCuentaCompraSoles, cfg.IdPlanCuentaSoles)
        FROM dbo.CON_DocumentoConfiguracionEmpresa AS cfg;';
END;

IF COL_LENGTH(N'dbo.CON_DocumentoConfiguracionEmpresa', N'IdPlanCuentaDolares') IS NOT NULL
BEGIN
    EXEC sys.sp_executesql N'
        UPDATE cfg
        SET IdCuentaVentaDolares = ISNULL(cfg.IdCuentaVentaDolares, cfg.IdPlanCuentaDolares),
            IdCuentaCompraDolares = ISNULL(cfg.IdCuentaCompraDolares, cfg.IdPlanCuentaDolares)
        FROM dbo.CON_DocumentoConfiguracionEmpresa AS cfg;';
END;

IF OBJECT_ID(N'dbo.FK_CON_DocumentoConfiguracionEmpresa_PlanCuentaSoles', N'F') IS NOT NULL
BEGIN
    ALTER TABLE dbo.CON_DocumentoConfiguracionEmpresa DROP CONSTRAINT FK_CON_DocumentoConfiguracionEmpresa_PlanCuentaSoles;
END;

IF OBJECT_ID(N'dbo.FK_CON_DocumentoConfiguracionEmpresa_PlanCuentaDolares', N'F') IS NOT NULL
BEGIN
    ALTER TABLE dbo.CON_DocumentoConfiguracionEmpresa DROP CONSTRAINT FK_CON_DocumentoConfiguracionEmpresa_PlanCuentaDolares;
END;

IF COL_LENGTH(N'dbo.CON_DocumentoConfiguracionEmpresa', N'IdPlanCuentaSoles') IS NOT NULL
BEGIN
    EXEC sys.sp_executesql N'
        ALTER TABLE dbo.CON_DocumentoConfiguracionEmpresa DROP COLUMN IdPlanCuentaSoles;';
END;

IF COL_LENGTH(N'dbo.CON_DocumentoConfiguracionEmpresa', N'IdPlanCuentaDolares') IS NOT NULL
BEGIN
    EXEC sys.sp_executesql N'
        ALTER TABLE dbo.CON_DocumentoConfiguracionEmpresa DROP COLUMN IdPlanCuentaDolares;';
END;

IF OBJECT_ID(N'dbo.FK_CON_DocumentoConfiguracionEmpresa_CuentaVentaSoles', N'F') IS NULL
BEGIN
    ALTER TABLE dbo.CON_DocumentoConfiguracionEmpresa
        ADD CONSTRAINT FK_CON_DocumentoConfiguracionEmpresa_CuentaVentaSoles
            FOREIGN KEY (IdCuentaVentaSoles) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);
END;

IF OBJECT_ID(N'dbo.FK_CON_DocumentoConfiguracionEmpresa_CuentaVentaDolares', N'F') IS NULL
BEGIN
    ALTER TABLE dbo.CON_DocumentoConfiguracionEmpresa
        ADD CONSTRAINT FK_CON_DocumentoConfiguracionEmpresa_CuentaVentaDolares
            FOREIGN KEY (IdCuentaVentaDolares) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);
END;

IF OBJECT_ID(N'dbo.FK_CON_DocumentoConfiguracionEmpresa_CuentaCompraSoles', N'F') IS NULL
BEGIN
    ALTER TABLE dbo.CON_DocumentoConfiguracionEmpresa
        ADD CONSTRAINT FK_CON_DocumentoConfiguracionEmpresa_CuentaCompraSoles
            FOREIGN KEY (IdCuentaCompraSoles) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);
END;

IF OBJECT_ID(N'dbo.FK_CON_DocumentoConfiguracionEmpresa_CuentaCompraDolares', N'F') IS NULL
BEGIN
    ALTER TABLE dbo.CON_DocumentoConfiguracionEmpresa
        ADD CONSTRAINT FK_CON_DocumentoConfiguracionEmpresa_CuentaCompraDolares
            FOREIGN KEY (IdCuentaCompraDolares) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);
END;
