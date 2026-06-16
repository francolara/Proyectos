-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Relaciona usuarios autenticados con las empresas a las que tienen acceso.
-- =============================================

IF OBJECT_ID(N'dbo.SEG_UsuarioEmpresa', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SEG_UsuarioEmpresa
    (
        IdUsuarioEmpresa INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_SEG_UsuarioEmpresa PRIMARY KEY,
        AspNetUserId NVARCHAR(450) NOT NULL,
        IdEmpresa INT NOT NULL,
        EsEmpresaPredeterminada BIT NOT NULL CONSTRAINT DF_SEG_UsuarioEmpresa_EsEmpresaPredeterminada DEFAULT (0),
        Estado BIT NOT NULL CONSTRAINT DF_SEG_UsuarioEmpresa_Estado DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_SEG_UsuarioEmpresa_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.SEG_UsuarioEmpresa
        ADD CONSTRAINT UQ_SEG_UsuarioEmpresa UNIQUE (AspNetUserId, IdEmpresa);

    ALTER TABLE dbo.SEG_UsuarioEmpresa
        ADD CONSTRAINT FK_SEG_UsuarioEmpresa_AspNetUsers
        FOREIGN KEY (AspNetUserId) REFERENCES dbo.AspNetUsers (Id);

    ALTER TABLE dbo.SEG_UsuarioEmpresa
        ADD CONSTRAINT FK_SEG_UsuarioEmpresa_SEG_Empresa
        FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);
END;
