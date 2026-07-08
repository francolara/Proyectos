-- =============================================
-- Author:        FRANCO LARA
-- Create date:   02/07/2026
-- Description:   Detalle de lineas del proceso de asiento de apertura.
-- =============================================
-- Firma: FRANCO LARA - 02/07/2026 | Guarda el detalle de lineas generadas en la apertura, separando saldos resumidos y saldos por referencia documental.

IF OBJECT_ID(N'dbo.CON_AperturaProcesoDetalle', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_AperturaProcesoDetalle
    (
        IdAperturaProcesoDetalle INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_AperturaProcesoDetalle PRIMARY KEY,
        IdAperturaProceso INT NOT NULL,
        Item SMALLINT NOT NULL,
        TipoDetalle NVARCHAR(20) NOT NULL,
        IdPlanCuenta INT NOT NULL,
        CodigoMoneda VARCHAR(3) NOT NULL CONSTRAINT DF_CON_AperturaProcesoDetalle_CodigoMoneda DEFAULT ('PEN'),
        TipoCambioAplicado DECIMAL(18,6) NOT NULL CONSTRAINT DF_CON_AperturaProcesoDetalle_TipoCambioAplicado DEFAULT (1),
        TipoDocumento NVARCHAR(150) NULL,
        Serie VARCHAR(10) NULL,
        NumeroDocumento VARCHAR(20) NULL,
        Debe DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_AperturaProcesoDetalle_Debe DEFAULT (0),
        Haber DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_AperturaProcesoDetalle_Haber DEFAULT (0),
        TotalImporteS DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_AperturaProcesoDetalle_TotalImporteS DEFAULT (0),
        TotalImporteD DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_AperturaProcesoDetalle_TotalImporteD DEFAULT (0),
        Observacion NVARCHAR(300) NULL,
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_AperturaProcesoDetalle_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_AperturaProcesoDetalle
        ADD CONSTRAINT FK_CON_AperturaProcesoDetalle_CON_AperturaProceso
            FOREIGN KEY (IdAperturaProceso) REFERENCES dbo.CON_AperturaProceso (IdAperturaProceso);

    ALTER TABLE dbo.CON_AperturaProcesoDetalle
        ADD CONSTRAINT FK_CON_AperturaProcesoDetalle_CON_PlanCuenta
            FOREIGN KEY (IdPlanCuenta) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);

    ALTER TABLE dbo.CON_AperturaProcesoDetalle
        ADD CONSTRAINT CK_CON_AperturaProcesoDetalle_Item
            CHECK (Item >= 1);

    ALTER TABLE dbo.CON_AperturaProcesoDetalle
        ADD CONSTRAINT CK_CON_AperturaProcesoDetalle_TipoDetalle
            CHECK (TipoDetalle IN (N'RESUMEN', N'ANALISIS'));

    ALTER TABLE dbo.CON_AperturaProcesoDetalle
        ADD CONSTRAINT UQ_CON_AperturaProcesoDetalle_Proceso_Item
            UNIQUE (IdAperturaProceso, Item);
END;
