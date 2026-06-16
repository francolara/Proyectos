-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Cabecera de reglas de cuentas de destino por empresa, ejercicio y cuenta contable origen.
-- =============================================

IF OBJECT_ID(N'dbo.CON_CuentaDestinoRegla', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_CuentaDestinoRegla
    (
        IdCuentaDestinoRegla INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_CuentaDestinoRegla PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        Ejercicio SMALLINT NOT NULL,
        IdPlanCuentaOrigen INT NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_CON_CuentaDestinoRegla_Activo DEFAULT (1),
        Observacion NVARCHAR(500) NULL,
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_CuentaDestinoRegla_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_CuentaDestinoRegla
        ADD CONSTRAINT FK_CON_CuentaDestinoRegla_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);

    ALTER TABLE dbo.CON_CuentaDestinoRegla
        ADD CONSTRAINT FK_CON_CuentaDestinoRegla_CON_PlanCuentaOrigen
            FOREIGN KEY (IdPlanCuentaOrigen) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);

    ALTER TABLE dbo.CON_CuentaDestinoRegla
        ADD CONSTRAINT UQ_CON_CuentaDestinoRegla
            UNIQUE (IdEmpresa, Ejercicio, IdPlanCuentaOrigen);
END;
