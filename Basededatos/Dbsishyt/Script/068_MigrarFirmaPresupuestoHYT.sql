-- Firma: FRANCO LARA | Create date: 17/11/2026 | Conserva la firma anterior como firma predeterminada del catálogo.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF NOT EXISTS(SELECT 1 FROM dbo.FirmaPresupuestoHYT)
    INSERT INTO dbo.FirmaPresupuestoHYT(Nombre,ArchivoJson,Predeterminada,UsuarioRegistro)
    SELECT COALESCE(NULLIF(TRIM(empresa.NombreFirmante),N''),N'Firma principal'),empresa.FirmaPresupuestoJson,1,empresa.UsuarioRegistro
    FROM dbo.ConfiguracionHYT AS empresa WHERE empresa.Id=1 AND ISJSON(empresa.FirmaPresupuestoJson)=1;
GO
