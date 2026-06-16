-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Detalle de cuentas destino y porcentaje de distribucion por regla contable.
-- =============================================

IF OBJECT_ID(N'dbo.CON_CuentaDestinoReglaDetalle', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_CuentaDestinoReglaDetalle
    (
        IdCuentaDestinoReglaDetalle INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_CuentaDestinoReglaDetalle PRIMARY KEY,
        IdCuentaDestinoRegla INT NOT NULL,
        Orden SMALLINT NOT NULL,
        IdPlanCuentaDestinoCargo INT NOT NULL,
        IdPlanCuentaDestinoAbono INT NOT NULL,
        Porcentaje DECIMAL(7,4) NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_CON_CuentaDestinoReglaDetalle_Activo DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_CuentaDestinoReglaDetalle_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_CuentaDestinoReglaDetalle
        ADD CONSTRAINT FK_CON_CuentaDestinoReglaDetalle_CON_CuentaDestinoRegla
            FOREIGN KEY (IdCuentaDestinoRegla) REFERENCES dbo.CON_CuentaDestinoRegla (IdCuentaDestinoRegla);

    ALTER TABLE dbo.CON_CuentaDestinoReglaDetalle
        ADD CONSTRAINT FK_CON_CuentaDestinoReglaDetalle_CON_PlanCuentaDestinoCargo
            FOREIGN KEY (IdPlanCuentaDestinoCargo) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);

    ALTER TABLE dbo.CON_CuentaDestinoReglaDetalle
        ADD CONSTRAINT FK_CON_CuentaDestinoReglaDetalle_CON_PlanCuentaDestinoAbono
            FOREIGN KEY (IdPlanCuentaDestinoAbono) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);

    ALTER TABLE dbo.CON_CuentaDestinoReglaDetalle
        ADD CONSTRAINT CK_CON_CuentaDestinoReglaDetalle_Orden
            CHECK (Orden >= 1);

    ALTER TABLE dbo.CON_CuentaDestinoReglaDetalle
        ADD CONSTRAINT CK_CON_CuentaDestinoReglaDetalle_Porcentaje
            CHECK (Porcentaje > 0 AND Porcentaje <= 100);

    ALTER TABLE dbo.CON_CuentaDestinoReglaDetalle
        ADD CONSTRAINT UQ_CON_CuentaDestinoReglaDetalle
            UNIQUE (IdCuentaDestinoRegla, Orden);
END;
