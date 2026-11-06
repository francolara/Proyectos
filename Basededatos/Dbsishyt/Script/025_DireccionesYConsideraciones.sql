-- Firma: FRANCO LARA | Create date: 05/11/2026 | Ubigeo de FralseCont y direcciones de clientes/tiendas.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF COL_LENGTH(N'dbo.Cliente',N'CodigoUbigeo') IS NULL ALTER TABLE dbo.Cliente ADD CodigoUbigeo CHAR(6) NULL CONSTRAINT FK_Cliente_Ubigeo REFERENCES dbo.UbigeoDistritos(CodigoUbigeo);
IF COL_LENGTH(N'dbo.Tienda',N'CodigoUbigeo') IS NULL ALTER TABLE dbo.Tienda ADD CodigoUbigeo CHAR(6) NULL CONSTRAINT FK_Tienda_Ubigeo REFERENCES dbo.UbigeoDistritos(CodigoUbigeo);
IF COL_LENGTH(N'dbo.ConfiguracionHYT',N'ConsideracionesPresupuesto') IS NULL ALTER TABLE dbo.ConfiguracionHYT ADD ConsideracionesPresupuesto NVARCHAR(MAX) NULL;
ALTER TABLE dbo.PresupuestoRevision ALTER COLUMN ClienteDireccion NVARCHAR(600) NULL;
ALTER TABLE dbo.PresupuestoRevision ALTER COLUMN TiendaDireccion NVARCHAR(600) NULL;
GO
