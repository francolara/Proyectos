-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define Permiso para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.Permiso', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Permiso
    (
        Codigo NVARCHAR(80) NOT NULL CONSTRAINT PK_Permiso PRIMARY KEY,
        Descripcion NVARCHAR(200) NOT NULL
    );
END;
GO
