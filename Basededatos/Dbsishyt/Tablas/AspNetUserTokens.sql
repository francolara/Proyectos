-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define AspNetUserTokens para la base inicial de HYT.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.AspNetUserTokens', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.AspNetUserTokens
    (
        UserId NVARCHAR(128) NOT NULL,
        LoginProvider NVARCHAR(128) NOT NULL,
        Name NVARCHAR(128) NOT NULL,
        Value NVARCHAR(MAX) NULL,
        CONSTRAINT PK_AspNetUserTokens PRIMARY KEY (UserId, LoginProvider, Name),
        CONSTRAINT FK_UserTokens_User FOREIGN KEY (UserId) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO
