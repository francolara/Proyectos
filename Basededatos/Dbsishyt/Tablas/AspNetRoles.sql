-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define AspNetRoles para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.AspNetRoles', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.AspNetRoles
    (
        Id NVARCHAR(128) NOT NULL CONSTRAINT PK_AspNetRoles PRIMARY KEY,
        Name NVARCHAR(256) NULL,
        NormalizedName NVARCHAR(256) NULL,
        ConcurrencyStamp NVARCHAR(MAX) NULL
    );
END;
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.AspNetRoles') AND name = N'RoleNameIndex')
    CREATE UNIQUE INDEX RoleNameIndex ON dbo.AspNetRoles(NormalizedName) WHERE NormalizedName IS NOT NULL;
GO
