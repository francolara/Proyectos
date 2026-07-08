-- =============================================
-- Author:        FRANCO LARA
-- Create date:   01/07/2026
-- Description:   Detalle por cuenta del proceso de diferencia en cambio y asiento generado.
-- =============================================
-- Firma: FRANCO LARA - 01/07/2026 | Guarda una fila por cuenta en dolares procesada, indicando si se calculo por saldo o por analisis y el asiento generado.

IF OBJECT_ID(N'dbo.CON_DiferenciaCambioProcesoDetalle', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_DiferenciaCambioProcesoDetalle
    (
        IdDiferenciaCambioProcesoDetalle INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_DiferenciaCambioProcesoDetalle PRIMARY KEY,
        IdDiferenciaCambioProceso INT NOT NULL,
        IdPlanCuenta INT NOT NULL,
        GeneraPorAnalisis BIT NOT NULL CONSTRAINT DF_CON_DiferenciaCambioProcesoDetalle_GeneraPorAnalisis DEFAULT (0),
        TipoCambioAplicado DECIMAL(18,6) NOT NULL CONSTRAINT DF_CON_DiferenciaCambioProcesoDetalle_TipoCambioAplicado DEFAULT (0),
        IdAsiento INT NULL,
        NumeroAsiento INT NULL,
        TotalDebe DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_DiferenciaCambioProcesoDetalle_TotalDebe DEFAULT (0),
        TotalHaber DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_DiferenciaCambioProcesoDetalle_TotalHaber DEFAULT (0),
        Estado NVARCHAR(30) NOT NULL CONSTRAINT DF_CON_DiferenciaCambioProcesoDetalle_Estado DEFAULT (N'PENDIENTE'),
        Observacion NVARCHAR(300) NULL,
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_DiferenciaCambioProcesoDetalle_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_DiferenciaCambioProcesoDetalle
        ADD CONSTRAINT FK_CON_DiferenciaCambioProcesoDetalle_CON_DiferenciaCambioProceso
            FOREIGN KEY (IdDiferenciaCambioProceso) REFERENCES dbo.CON_DiferenciaCambioProceso (IdDiferenciaCambioProceso);

    ALTER TABLE dbo.CON_DiferenciaCambioProcesoDetalle
        ADD CONSTRAINT FK_CON_DiferenciaCambioProcesoDetalle_CON_PlanCuenta
            FOREIGN KEY (IdPlanCuenta) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);

    ALTER TABLE dbo.CON_DiferenciaCambioProcesoDetalle
        ADD CONSTRAINT FK_CON_DiferenciaCambioProcesoDetalle_CON_Asiento
            FOREIGN KEY (IdAsiento) REFERENCES dbo.CON_Asiento (IdAsiento);

    ALTER TABLE dbo.CON_DiferenciaCambioProcesoDetalle
        ADD CONSTRAINT UQ_CON_DiferenciaCambioProcesoDetalle_Proceso_Cuenta
            UNIQUE (IdDiferenciaCambioProceso, IdPlanCuenta);
END;
