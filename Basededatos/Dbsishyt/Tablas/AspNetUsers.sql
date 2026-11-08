-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Define AspNetUsers para la base inicial de HYT.
-- =============================================
-- Firma: FRANCO LARA | Create date: 06/11/2026 | Apellidos, cambio obligatorio, último ingreso y responsables de administración.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.AspNetUsers', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.AspNetUsers
    (
        Id NVARCHAR(128) NOT NULL CONSTRAINT PK_AspNetUsers PRIMARY KEY,
        UserName NVARCHAR(256) NULL,
        NormalizedUserName NVARCHAR(256) NULL,
        Email NVARCHAR(256) NULL,
        NormalizedEmail NVARCHAR(256) NULL,
        EmailConfirmed BIT NOT NULL DEFAULT 0,
        PasswordHash NVARCHAR(MAX) NULL,
        SecurityStamp NVARCHAR(MAX) NULL,
        ConcurrencyStamp NVARCHAR(MAX) NULL,
        PhoneNumber NVARCHAR(MAX) NULL,
        PhoneNumberConfirmed BIT NOT NULL DEFAULT 0,
        TwoFactorEnabled BIT NOT NULL DEFAULT 0,
        LockoutEnd DATETIMEOFFSET NULL,
        LockoutEnabled BIT NOT NULL DEFAULT 1,
        AccessFailedCount INT NOT NULL DEFAULT 0,
        Nombres NVARCHAR(150) NOT NULL,
        Apellidos NVARCHAR(150) NOT NULL CONSTRAINT DF_AspNetUsers_Apellidos DEFAULT N'',
        DebeCambiarClave BIT NOT NULL CONSTRAINT DF_AspNetUsers_CambiarClave DEFAULT 0,
        FechaCambioClave DATETIME2(3) NULL,
        UltimoIngreso DATETIME2(3) NULL,
        UsuarioRegistro NVARCHAR(128) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        Activo BIT NOT NULL DEFAULT 1,
        FechaRegistro DATETIME2(3) NOT NULL DEFAULT SYSUTCDATETIME(),
        FechaActualizacion DATETIME2(3) NULL
    );
END;
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.AspNetUsers') AND name = N'UserNameIndex')
    CREATE UNIQUE INDEX UserNameIndex ON dbo.AspNetUsers(NormalizedUserName) WHERE NormalizedUserName IS NOT NULL;
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.AspNetUsers') AND name = N'EmailIndex')
    CREATE INDEX EmailIndex ON dbo.AspNetUsers(NormalizedEmail);
GO
