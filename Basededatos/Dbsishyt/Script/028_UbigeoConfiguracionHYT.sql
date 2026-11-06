-- Firma: FRANCO LARA | Create date: 06/11/2026 | Ubigeo de la dirección de HYT, con referencia al catálogo de distritos.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF COL_LENGTH(N'dbo.ConfiguracionHYT',N'CodigoUbigeo') IS NULL ALTER TABLE dbo.ConfiguracionHYT ADD CodigoUbigeo CHAR(6) NULL;
GO
IF NOT EXISTS(SELECT 1 FROM sys.foreign_keys AS fk WHERE fk.name=N'FK_ConfiguracionHYT_Ubigeo' AND fk.parent_object_id=OBJECT_ID(N'dbo.ConfiguracionHYT')) ALTER TABLE dbo.ConfiguracionHYT ADD CONSTRAINT FK_ConfiguracionHYT_Ubigeo FOREIGN KEY(CodigoUbigeo) REFERENCES dbo.UbigeoDistritos(CodigoUbigeo);
GO
