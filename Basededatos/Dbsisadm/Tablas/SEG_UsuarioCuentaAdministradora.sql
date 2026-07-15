-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Relaciona usuarios autenticados con la cuenta administradora que gestiona multiples empresas.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   12/07/2026
-- Description:   Corrige el rol por defecto de la cuenta administradora a ADMINISTRADORCUENTA para evitar altas legacy con ADMINISTRADOR.
-- =============================================

IF OBJECT_ID(N'dbo.SEG_UsuarioCuentaAdministradora', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SEG_UsuarioCuentaAdministradora
    (
        IdUsuarioCuentaAdministradora INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_SEG_UsuarioCuentaAdministradora PRIMARY KEY,
        AspNetUserId NVARCHAR(450) NOT NULL,
        IdCuentaAdministradora INT NOT NULL,
        RolCuenta NVARCHAR(30) NOT NULL CONSTRAINT DF_SEG_UsuarioCuentaAdministradora_RolCuenta DEFAULT (N'ADMINISTRADORCUENTA'),
        EsCuentaPredeterminada BIT NOT NULL CONSTRAINT DF_SEG_UsuarioCuentaAdministradora_EsCuentaPredeterminada DEFAULT (0),
        Estado BIT NOT NULL CONSTRAINT DF_SEG_UsuarioCuentaAdministradora_Estado DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_SEG_UsuarioCuentaAdministradora_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.SEG_UsuarioCuentaAdministradora
        ADD CONSTRAINT UQ_SEG_UsuarioCuentaAdministradora UNIQUE (AspNetUserId, IdCuentaAdministradora);

    ALTER TABLE dbo.SEG_UsuarioCuentaAdministradora
        ADD CONSTRAINT FK_SEG_UsuarioCuentaAdministradora_AspNetUsers
        FOREIGN KEY (AspNetUserId) REFERENCES dbo.AspNetUsers (Id);

    ALTER TABLE dbo.SEG_UsuarioCuentaAdministradora
        ADD CONSTRAINT FK_SEG_UsuarioCuentaAdministradora_SEG_CuentaAdministradora
        FOREIGN KEY (IdCuentaAdministradora) REFERENCES dbo.SEG_CuentaAdministradora (IdCuentaAdministradora);
END;
