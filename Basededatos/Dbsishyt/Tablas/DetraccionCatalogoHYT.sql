-- Firma: FRANCO LARA | Create date: 10/11/2026 | Catálogo de detracciones copiado de FralseCont.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.DetraccionCatalogoHYT',N'U') IS NULL
BEGIN
    CREATE TABLE dbo.DetraccionCatalogoHYT
    (
        Codigo NVARCHAR(3) NOT NULL CONSTRAINT PK_DetraccionCatalogoHYT PRIMARY KEY,
        Nombre NVARCHAR(250) NOT NULL,
        Porcentaje DECIMAL(7,4) NOT NULL CHECK (Porcentaje BETWEEN 0 AND 100),
        Activo BIT NOT NULL
    );
END;
GO
