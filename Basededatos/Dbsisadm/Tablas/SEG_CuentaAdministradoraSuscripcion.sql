-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Suscripcion comercial de la cuenta administradora, independiente de cada empresa.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/07/2026
-- Description:   Amplia la suscripcion por cuenta con tipo de cobro, dias de gracia y trazabilidad de actualizacion.
-- =============================================

IF OBJECT_ID(N'dbo.SEG_CuentaAdministradoraSuscripcion', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SEG_CuentaAdministradoraSuscripcion
    (
        IdCuentaAdministradoraSuscripcion INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_SEG_CuentaAdministradoraSuscripcion PRIMARY KEY,
        IdCuentaAdministradora INT NOT NULL,
        TipoPlan NVARCHAR(50) NOT NULL CONSTRAINT DF_SEG_CuentaAdministradoraSuscripcion_TipoPlan DEFAULT (N'TRIAL'),
        EstadoSuscripcion NVARCHAR(20) NOT NULL CONSTRAINT DF_SEG_CuentaAdministradoraSuscripcion_Estado DEFAULT (N'TRIAL'),
        EsPrueba BIT NOT NULL CONSTRAINT DF_SEG_CuentaAdministradoraSuscripcion_EsPrueba DEFAULT (1),
        FechaInicioPrueba DATE NULL,
        FechaFinPrueba DATE NULL,
        FechaInicioPlan DATE NULL,
        FechaFinPlan DATE NULL,
        TipoCobro NVARCHAR(20) NULL,
        DiasGracia INT NOT NULL CONSTRAINT DF_SEG_CuentaAdministradoraSuscripcion_DiasGracia DEFAULT (5),
        FechaFinGracia DATE NULL,
        EmpresasPermitidas INT NULL,
        UsuariosPermitidos INT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_SEG_CuentaAdministradoraSuscripcion_Activo DEFAULT (1),
        Observacion NVARCHAR(500) NULL,
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_SEG_CuentaAdministradoraSuscripcion_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL,
        FechaActualizacion DATETIME2(0) NULL,
        UsuarioActualizacion NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcion
        ADD CONSTRAINT UQ_SEG_CuentaAdministradoraSuscripcion_IdCuenta UNIQUE (IdCuentaAdministradora);

    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcion
        ADD CONSTRAINT FK_SEG_CuentaAdministradoraSuscripcion_SEG_CuentaAdministradora
        FOREIGN KEY (IdCuentaAdministradora) REFERENCES dbo.SEG_CuentaAdministradora (IdCuentaAdministradora);

    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcion
        ADD CONSTRAINT CK_SEG_CuentaAdministradoraSuscripcion_Estado
            CHECK (EstadoSuscripcion IN (N'TRIAL', N'ACTIVO', N'SUSPENDIDO', N'BAJA'));
END;
