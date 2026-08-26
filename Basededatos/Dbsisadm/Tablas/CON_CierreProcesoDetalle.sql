-- =============================================
-- Author:        FRANCO LARA
-- Create date:   02/07/2026
-- Description:   Detalle por cuenta del proceso de asiento de cierre y asiento generado.
-- =============================================
-- Firma: FRANCO LARA - 02/07/2026 | Guarda una fila por cuenta cerrada indicando si corresponde al periodo 14 o 15 y que asiento se genero.
-- Firma: FRANCO LARA - 13/08/2026 | Convierte el detalle en las lineas del asiento compuesto, conservando item, sentido contable e importes acumulados en soles y dolares.
-- Firma: FRANCO LARA - 22/08/2026 | Restringe el tipo de cierre al periodo 14, reservado para el cierre de Inventario.

IF OBJECT_ID(N'dbo.CON_CierreProcesoDetalle', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_CierreProcesoDetalle
    (
        IdCierreProcesoDetalle INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_CierreProcesoDetalle PRIMARY KEY,
        IdCierreProceso INT NOT NULL,
        Item SMALLINT NOT NULL,
        TipoCierre CHAR(2) NOT NULL,
        IdPlanCuenta INT NOT NULL,
        CodigoMoneda VARCHAR(3) NOT NULL CONSTRAINT DF_CON_CierreProcesoDetalle_CodigoMoneda DEFAULT ('PEN'),
        TipoCambioAplicado DECIMAL(18,6) NOT NULL CONSTRAINT DF_CON_CierreProcesoDetalle_TipoCambioAplicado DEFAULT (1),
        IdAsiento INT NULL,
        NumeroAsiento INT NULL,
        DH CHAR(1) NOT NULL,
        TotalDebe DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_CierreProcesoDetalle_TotalDebe DEFAULT (0),
        TotalHaber DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_CierreProcesoDetalle_TotalHaber DEFAULT (0),
        TotalImporteS DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_CierreProcesoDetalle_TotalImporteS DEFAULT (0),
        TotalImporteD DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_CierreProcesoDetalle_TotalImporteD DEFAULT (0),
        Estado NVARCHAR(30) NOT NULL CONSTRAINT DF_CON_CierreProcesoDetalle_Estado DEFAULT (N'PENDIENTE'),
        Observacion NVARCHAR(300) NULL,
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_CierreProcesoDetalle_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_CierreProcesoDetalle
        ADD CONSTRAINT FK_CON_CierreProcesoDetalle_CON_CierreProceso
            FOREIGN KEY (IdCierreProceso) REFERENCES dbo.CON_CierreProceso (IdCierreProceso);

    ALTER TABLE dbo.CON_CierreProcesoDetalle
        ADD CONSTRAINT FK_CON_CierreProcesoDetalle_CON_PlanCuenta
            FOREIGN KEY (IdPlanCuenta) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);

    ALTER TABLE dbo.CON_CierreProcesoDetalle
        ADD CONSTRAINT FK_CON_CierreProcesoDetalle_CON_Asiento
            FOREIGN KEY (IdAsiento) REFERENCES dbo.CON_Asiento (IdAsiento);

    ALTER TABLE dbo.CON_CierreProcesoDetalle
        ADD CONSTRAINT CK_CON_CierreProcesoDetalle_TipoCierre
            CHECK (TipoCierre = '14');

    ALTER TABLE dbo.CON_CierreProcesoDetalle
        ADD CONSTRAINT CK_CON_CierreProcesoDetalle_Item
            CHECK (Item >= 1);

    ALTER TABLE dbo.CON_CierreProcesoDetalle
        ADD CONSTRAINT CK_CON_CierreProcesoDetalle_DH
            CHECK (DH IN ('D', 'H'));

    ALTER TABLE dbo.CON_CierreProcesoDetalle
        ADD CONSTRAINT UQ_CON_CierreProcesoDetalle_Proceso_Tipo_Cuenta
            UNIQUE (IdCierreProceso, TipoCierre, IdPlanCuenta);
END;
