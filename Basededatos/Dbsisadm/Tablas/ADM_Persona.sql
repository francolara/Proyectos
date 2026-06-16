-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Maestro general de personas para clientes, proveedores y representantes de empresa.
-- =============================================

IF OBJECT_ID(N'dbo.ADM_Persona', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ADM_Persona
    (
        IdPersona INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_ADM_Persona PRIMARY KEY,
        TipoPersona CHAR(1) NOT NULL,
        TipoDocumento VARCHAR(3) NOT NULL,
        NumeroDocumento VARCHAR(20) NOT NULL,
        ApellidoPaterno NVARCHAR(100) NULL,
        ApellidoMaterno NVARCHAR(100) NULL,
        Nombres NVARCHAR(150) NULL,
        RazonSocial NVARCHAR(200) NULL,
        NombreCompleto AS
        (
            LTRIM(RTRIM(
                COALESCE(RazonSocial,
                    CONCAT(
                        COALESCE(ApellidoPaterno, N''),
                        N' ',
                        COALESCE(ApellidoMaterno, N''),
                        N' ',
                        COALESCE(Nombres, N'')
                    )
                )
            ))
        ) PERSISTED,
        CorreoElectronico NVARCHAR(200) NULL,
        Telefono NVARCHAR(50) NULL,
        Direccion NVARCHAR(250) NULL,
        Estado BIT NOT NULL CONSTRAINT DF_ADM_Persona_Estado DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_ADM_Persona_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.ADM_Persona
        ADD CONSTRAINT CK_ADM_Persona_TipoPersona
        CHECK (TipoPersona IN ('N', 'J'));

    ALTER TABLE dbo.ADM_Persona
        ADD CONSTRAINT UQ_ADM_Persona_Documento UNIQUE (TipoDocumento, NumeroDocumento);
END;
