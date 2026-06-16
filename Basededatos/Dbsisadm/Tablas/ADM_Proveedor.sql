-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Maestro de proveedores por empresa.
-- =============================================

IF OBJECT_ID(N'dbo.ADM_Proveedor', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ADM_Proveedor
    (
        IdProveedor INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_ADM_Proveedor PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        IdPersona INT NOT NULL,
        CodigoProveedor VARCHAR(20) NOT NULL,
        Contacto NVARCHAR(150) NULL,
        CuentaDetraccion NVARCHAR(50) NULL,
        Observacion NVARCHAR(300) NULL,
        Estado BIT NOT NULL CONSTRAINT DF_ADM_Proveedor_Estado DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_ADM_Proveedor_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.ADM_Proveedor
        ADD CONSTRAINT FK_ADM_Proveedor_SEG_Empresa
        FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);

    ALTER TABLE dbo.ADM_Proveedor
        ADD CONSTRAINT FK_ADM_Proveedor_ADM_Persona
        FOREIGN KEY (IdPersona) REFERENCES dbo.ADM_Persona (IdPersona);

    ALTER TABLE dbo.ADM_Proveedor
        ADD CONSTRAINT UQ_ADM_Proveedor_EmpresaPersona UNIQUE (IdEmpresa, IdPersona);

    ALTER TABLE dbo.ADM_Proveedor
        ADD CONSTRAINT UQ_ADM_Proveedor_EmpresaCodigo UNIQUE (IdEmpresa, CodigoProveedor);
END;
