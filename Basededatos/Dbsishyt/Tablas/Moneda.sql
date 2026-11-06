-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define monedas internas por codigo textual, sin Id numerico.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.Moneda', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Moneda
    (
        Codigo NVARCHAR(3) NOT NULL CONSTRAINT PK_Moneda PRIMARY KEY,
        Nombre NVARCHAR(40) NOT NULL,
        Simbolo NVARCHAR(5) NOT NULL
    );
END;
GO
