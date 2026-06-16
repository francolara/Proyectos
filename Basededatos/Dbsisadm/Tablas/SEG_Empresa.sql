-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Tabla de empresas del sistema administrativo multiempresa, ligada a la cuenta administradora titular.
-- =============================================

IF OBJECT_ID(N'dbo.SEG_Empresa', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SEG_Empresa
    (
        IdEmpresa INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_SEG_Empresa PRIMARY KEY,
        IdCuentaAdministradora INT NOT NULL,
        CodigoEmpresa VARCHAR(20) NOT NULL,
        RazonSocial NVARCHAR(200) NOT NULL,
        NombreComercial NVARCHAR(200) NULL,
        Ruc VARCHAR(11) NOT NULL,
        Estado BIT NOT NULL CONSTRAINT DF_SEG_Empresa_Estado DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_SEG_Empresa_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.SEG_Empresa
        ADD CONSTRAINT UQ_SEG_Empresa_CodigoEmpresa UNIQUE (CodigoEmpresa);

    ALTER TABLE dbo.SEG_Empresa
        ADD CONSTRAINT UQ_SEG_Empresa_Ruc UNIQUE (Ruc);

    IF OBJECT_ID(N'dbo.SEG_CuentaAdministradora', N'U') IS NOT NULL
    BEGIN
        ALTER TABLE dbo.SEG_Empresa
            ADD CONSTRAINT FK_SEG_Empresa_SEG_CuentaAdministradora
            FOREIGN KEY (IdCuentaAdministradora) REFERENCES dbo.SEG_CuentaAdministradora (IdCuentaAdministradora);
    END;
END;
