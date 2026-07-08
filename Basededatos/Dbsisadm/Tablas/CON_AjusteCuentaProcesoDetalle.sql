-- =============================================
-- Author:        FRANCO LARA
-- Create date:   02/07/2026
-- Description:   Detalle por cuenta del proceso de ajuste de cuentas y asiento generado.
-- =============================================
-- Firma: FRANCO LARA - 02/07/2026 | Guarda una fila por cuenta analitica procesada indicando moneda, cantidad de analisis residuales y el asiento AJU generado.

IF OBJECT_ID(N'dbo.CON_AjusteCuentaProcesoDetalle', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_AjusteCuentaProcesoDetalle
    (
        IdAjusteCuentaProcesoDetalle INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_AjusteCuentaProcesoDetalle PRIMARY KEY,
        IdAjusteCuentaProceso INT NOT NULL,
        IdPlanCuenta INT NOT NULL,
        CodigoMoneda VARCHAR(3) NOT NULL CONSTRAINT DF_CON_AjusteCuentaProcesoDetalle_CodigoMoneda DEFAULT ('PEN'),
        TipoCambioAplicado DECIMAL(18,6) NOT NULL CONSTRAINT DF_CON_AjusteCuentaProcesoDetalle_TipoCambioAplicado DEFAULT (1),
        TotalAnalisis INT NOT NULL CONSTRAINT DF_CON_AjusteCuentaProcesoDetalle_TotalAnalisis DEFAULT (0),
        IdAsiento INT NULL,
        NumeroAsiento INT NULL,
        TotalDebe DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_AjusteCuentaProcesoDetalle_TotalDebe DEFAULT (0),
        TotalHaber DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_AjusteCuentaProcesoDetalle_TotalHaber DEFAULT (0),
        Estado NVARCHAR(30) NOT NULL CONSTRAINT DF_CON_AjusteCuentaProcesoDetalle_Estado DEFAULT (N'PENDIENTE'),
        Observacion NVARCHAR(300) NULL,
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_AjusteCuentaProcesoDetalle_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_AjusteCuentaProcesoDetalle
        ADD CONSTRAINT FK_CON_AjusteCuentaProcesoDetalle_CON_AjusteCuentaProceso
            FOREIGN KEY (IdAjusteCuentaProceso) REFERENCES dbo.CON_AjusteCuentaProceso (IdAjusteCuentaProceso);

    ALTER TABLE dbo.CON_AjusteCuentaProcesoDetalle
        ADD CONSTRAINT FK_CON_AjusteCuentaProcesoDetalle_CON_PlanCuenta
            FOREIGN KEY (IdPlanCuenta) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);

    ALTER TABLE dbo.CON_AjusteCuentaProcesoDetalle
        ADD CONSTRAINT FK_CON_AjusteCuentaProcesoDetalle_CON_Asiento
            FOREIGN KEY (IdAsiento) REFERENCES dbo.CON_Asiento (IdAsiento);

    ALTER TABLE dbo.CON_AjusteCuentaProcesoDetalle
        ADD CONSTRAINT UQ_CON_AjusteCuentaProcesoDetalle_Proceso_Cuenta
            UNIQUE (IdAjusteCuentaProceso, IdPlanCuenta);
END;
