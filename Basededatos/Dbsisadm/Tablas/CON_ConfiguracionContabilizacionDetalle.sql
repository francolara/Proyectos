-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Detalle de componentes contables y cuentas asociadas para compras y ventas automaticas.
-- =============================================

IF OBJECT_ID(N'dbo.CON_ConfiguracionContabilizacionDetalle', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_ConfiguracionContabilizacionDetalle
    (
        IdConfiguracionContabilizacionDetalle INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_ConfiguracionContabilizacionDetalle PRIMARY KEY,
        IdConfiguracionContabilizacion INT NOT NULL,
        Orden SMALLINT NOT NULL,
        ComponenteContable VARCHAR(20) NOT NULL,
        IdPlanCuenta INT NOT NULL,
        NaturalezaMovimiento CHAR(1) NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_CON_ConfiguracionContabilizacionDetalle_Activo DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_ConfiguracionContabilizacionDetalle_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_ConfiguracionContabilizacionDetalle
        ADD CONSTRAINT FK_CON_ConfiguracionContabilizacionDetalle_Cabecera
            FOREIGN KEY (IdConfiguracionContabilizacion) REFERENCES dbo.CON_ConfiguracionContabilizacion (IdConfiguracionContabilizacion);

    ALTER TABLE dbo.CON_ConfiguracionContabilizacionDetalle
        ADD CONSTRAINT FK_CON_ConfiguracionContabilizacionDetalle_PlanCuenta
            FOREIGN KEY (IdPlanCuenta) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);

    ALTER TABLE dbo.CON_ConfiguracionContabilizacionDetalle
        ADD CONSTRAINT CK_CON_ConfiguracionContabilizacionDetalle_Orden
            CHECK (Orden >= 1);

    ALTER TABLE dbo.CON_ConfiguracionContabilizacionDetalle
        ADD CONSTRAINT CK_CON_ConfiguracionContabilizacionDetalle_ComponenteContable
            CHECK (ComponenteContable IN ('BRUTO', 'IGV', 'TOTAL', 'REDONDEO', 'ISC', 'OTROS'));

    ALTER TABLE dbo.CON_ConfiguracionContabilizacionDetalle
        ADD CONSTRAINT CK_CON_ConfiguracionContabilizacionDetalle_NaturalezaMovimiento
            CHECK (NaturalezaMovimiento IN ('D', 'H'));

    ALTER TABLE dbo.CON_ConfiguracionContabilizacionDetalle
        ADD CONSTRAINT UQ_CON_ConfiguracionContabilizacionDetalle_Orden
            UNIQUE (IdConfiguracionContabilizacion, Orden);
END;
