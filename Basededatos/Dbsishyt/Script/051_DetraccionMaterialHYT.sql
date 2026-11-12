-- Firma: FRANCO LARA | Create date: 10/11/2026 | Asigna un código de detracción opcional en el maestro de materiales y servicios.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF COL_LENGTH(N'dbo.ServicioMaterial',N'CodigoDetraccion') IS NULL
    ALTER TABLE dbo.ServicioMaterial ADD CodigoDetraccion NVARCHAR(3) NULL;
GO
IF OBJECT_ID(N'dbo.FK_ServicioMaterial_DetraccionCatalogoHYT',N'F') IS NULL
    ALTER TABLE dbo.ServicioMaterial WITH CHECK ADD CONSTRAINT FK_ServicioMaterial_DetraccionCatalogoHYT
        FOREIGN KEY (CodigoDetraccion) REFERENCES dbo.DetraccionCatalogoHYT(Codigo);
GO
