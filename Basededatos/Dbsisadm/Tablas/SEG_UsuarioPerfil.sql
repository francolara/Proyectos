-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Perfil complementario del usuario autenticado.
-- =============================================

IF OBJECT_ID(N'dbo.SEG_UsuarioPerfil', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SEG_UsuarioPerfil
    (
        IdUsuarioPerfil INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_SEG_UsuarioPerfil PRIMARY KEY,
        AspNetUserId NVARCHAR(450) NOT NULL,
        NombreCompleto NVARCHAR(180) NOT NULL,
        Telefono NVARCHAR(30) NULL,
        CorreoReferencia NVARCHAR(256) NULL,
        Estado BIT NOT NULL CONSTRAINT DF_SEG_UsuarioPerfil_Estado DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_SEG_UsuarioPerfil_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.SEG_UsuarioPerfil
        ADD CONSTRAINT FK_SEG_UsuarioPerfil_AspNetUsers
            FOREIGN KEY (AspNetUserId) REFERENCES dbo.AspNetUsers (Id);

    ALTER TABLE dbo.SEG_UsuarioPerfil
        ADD CONSTRAINT UQ_SEG_UsuarioPerfil_AspNetUserId
            UNIQUE (AspNetUserId);
END;
