-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define AspNetRoleClaims para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.AspNetRoleClaims', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.AspNetRoleClaims
    (
        Id INT IDENTITY NOT NULL CONSTRAINT PK_AspNetRoleClaims PRIMARY KEY,
        RoleId NVARCHAR(128) NOT NULL,
        ClaimType NVARCHAR(MAX) NULL,
        ClaimValue NVARCHAR(MAX) NULL,
        CONSTRAINT FK_AspNetRoleClaims_Owner FOREIGN KEY (RoleId) REFERENCES dbo.AspNetRoles(Id)
    );
END;
GO
