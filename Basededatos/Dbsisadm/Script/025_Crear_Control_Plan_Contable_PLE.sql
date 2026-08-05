-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   04/08/2026
-- Description:   Crea el control idempotente de huellas para los planes contables PLE 5.3 y 5.4.
-- =============================================

IF OBJECT_ID(N'dbo.CON_PLE_PlanContableControl', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_PLE_PlanContableControl
    (
        IdPLEPlanContableControl INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_PLE_PlanContableControl PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        Anio SMALLINT NOT NULL,
        CodigoFormato VARCHAR(10) NOT NULL,
        HuellaPlanContable CHAR(64) NOT NULL,
        FechaUltimaGeneracion DATETIME2(0) NOT NULL CONSTRAINT DF_CON_PLE_PlanContableControl_Fecha DEFAULT (SYSDATETIME()),
        UsuarioGeneracion NVARCHAR(450) NULL,
        CONSTRAINT UQ_CON_PLE_PlanContableControl_EmpresaAnioFormato UNIQUE (IdEmpresa, Anio, CodigoFormato),
        CONSTRAINT FK_CON_PLE_PlanContableControl_SEG_Empresa FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa),
        CONSTRAINT CK_CON_PLE_PlanContableControl_Anio CHECK (Anio BETWEEN 2000 AND 2199),
        CONSTRAINT CK_CON_PLE_PlanContableControl_Formato CHECK (CodigoFormato IN ('5.3', '5.4'))
    );
END;
