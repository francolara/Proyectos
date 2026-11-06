-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Snapshots de presupuesto/impresión y regla de cantidades/precios a dos decimales.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF COL_LENGTH(N'dbo.PresupuestoRevision',N'ZonaTexto') IS NULL
    ALTER TABLE dbo.PresupuestoRevision ADD ZonaTexto NVARCHAR(200) NULL;
IF COL_LENGTH(N'dbo.PresupuestoRevision',N'EmpresaSnapshotJson') IS NULL
    ALTER TABLE dbo.PresupuestoRevision ADD EmpresaSnapshotJson NVARCHAR(MAX) NULL;
IF COL_LENGTH(N'dbo.PresupuestoRevision',N'PoliticaCalculo') IS NULL
    ALTER TABLE dbo.PresupuestoRevision ADD PoliticaCalculo NVARCHAR(20) NOT NULL CONSTRAINT DF_Revision_Politica DEFAULT N'LINEA_2_V1' WITH VALUES;
IF COL_LENGTH(N'dbo.PresupuestoRevision',N'FechaAprobacion') IS NULL
    ALTER TABLE dbo.PresupuestoRevision ADD FechaAprobacion DATE NULL;
IF COL_LENGTH(N'dbo.PresupuestoRevision',N'DescripcionTrabajo') IS NULL
    ALTER TABLE dbo.PresupuestoRevision ADD DescripcionTrabajo NVARCHAR(1500) NULL;
IF COL_LENGTH(N'dbo.PresupuestoRevision',N'SimboloMoneda') IS NULL
    ALTER TABLE dbo.PresupuestoRevision ADD SimboloMoneda NVARCHAR(5) NULL;
IF COL_LENGTH(N'dbo.PresupuestoRevision',N'TiendaCodigo') IS NULL
    ALTER TABLE dbo.PresupuestoRevision ADD TiendaCodigo NVARCHAR(40) NULL;
IF COL_LENGTH(N'dbo.PresupuestoPartida',N'ZonasTexto') IS NULL
    ALTER TABLE dbo.PresupuestoPartida ADD ZonasTexto NVARCHAR(1000) NULL;
IF COL_LENGTH(N'dbo.PresupuestoPartida',N'OrigenPrecio') IS NULL
    ALTER TABLE dbo.PresupuestoPartida ADD OrigenPrecio NVARCHAR(20) NOT NULL CONSTRAINT DF_Partida_Origen DEFAULT N'MANUAL' WITH VALUES;
IF COL_LENGTH(N'dbo.PresupuestoPartida',N'MotivoPrecioManual') IS NULL
    ALTER TABLE dbo.PresupuestoPartida ADD MotivoPrecioManual NVARCHAR(300) NULL;
IF COL_LENGTH(N'dbo.ImpresionPresupuesto',N'ContenidoHTML') IS NULL
    ALTER TABLE dbo.ImpresionPresupuesto ADD ContenidoHTML NVARCHAR(MAX) NULL;
IF COL_LENGTH(N'dbo.ImpresionPresupuesto',N'SHA256') IS NULL
    ALTER TABLE dbo.ImpresionPresupuesto ADD SHA256 CHAR(64) NULL;
GO
IF NOT EXISTS (SELECT 1 FROM sys.check_constraints WHERE name=N'CK_Revision_EmpresaJSON')
    ALTER TABLE dbo.PresupuestoRevision WITH CHECK ADD CONSTRAINT CK_Revision_EmpresaJSON CHECK (EmpresaSnapshotJson IS NULL OR ISJSON(EmpresaSnapshotJson)=1);
IF NOT EXISTS (SELECT 1 FROM sys.check_constraints WHERE name=N'CK_Tarifa_DosDecimales')
    ALTER TABLE dbo.TarifarioPartida WITH CHECK ADD CONSTRAINT CK_Tarifa_DosDecimales CHECK
        ((PrecioPEN IS NULL OR PrecioPEN=ROUND(PrecioPEN,2)) AND (PrecioUSD IS NULL OR PrecioUSD=ROUND(PrecioUSD,2)));
IF NOT EXISTS (SELECT 1 FROM sys.check_constraints WHERE name=N'CK_Partida_DosDecimales')
    ALTER TABLE dbo.PresupuestoPartida WITH CHECK ADD CONSTRAINT CK_Partida_DosDecimales CHECK
        (Cantidad=ROUND(Cantidad,2) AND (PrecioOrigen IS NULL OR PrecioOrigen=ROUND(PrecioOrigen,2)) AND (PrecioUnitario IS NULL OR PrecioUnitario=ROUND(PrecioUnitario,2)));
IF NOT EXISTS (SELECT 1 FROM sys.check_constraints WHERE name=N'CK_GuiaPartida_DosDecimales')
    ALTER TABLE dbo.GuiaPartida WITH CHECK ADD CONSTRAINT CK_GuiaPartida_DosDecimales CHECK (Cantidad=ROUND(Cantidad,2));
GO
