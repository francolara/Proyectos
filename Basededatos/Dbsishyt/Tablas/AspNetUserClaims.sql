-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define AspNetUserClaims para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.AspNetUserClaims', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.AspNetUserClaims
    (
        Id INT IDENTITY NOT NULL CONSTRAINT PK_AspNetUserClaims PRIMARY KEY,
        UserId NVARCHAR(128) NOT NULL,
        ClaimType NVARCHAR(MAX) NULL,
        ClaimValue NVARCHAR(MAX) NULL,
        CONSTRAINT FK_AspNetUserClaims_Owner FOREIGN KEY (UserId) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO
