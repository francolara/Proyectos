-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   BancoPeruHYT
-- =============================================
USE [SistemaWebHYT];
GO
IF OBJECT_ID(N'dbo.BancoPeruHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.BancoPeruHYT
    (
        Nombre NVARCHAR(100) NOT NULL CONSTRAINT PK_BancoPeruHYT PRIMARY KEY
    );
END;
GO
