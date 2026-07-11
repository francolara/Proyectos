-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Historial comercial de la suscripcion de la cuenta administradora.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/07/2026
-- Description:   Agrega al historial comercial por cuenta el tipo de cobro, dias de gracia y dias extra para contratos y ajustes.
-- =============================================

IF OBJECT_ID(N'dbo.SEG_CuentaAdministradoraSuscripcionMovimiento', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SEG_CuentaAdministradoraSuscripcionMovimiento
    (
        IdCuentaAdministradoraSuscripcionMovimiento INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_SEG_CuentaAdministradoraSuscripcionMovimiento PRIMARY KEY,
        IdCuentaAdministradora INT NOT NULL,
        IdCuentaAdministradoraSuscripcion INT NULL,
        TipoMovimiento NVARCHAR(40) NOT NULL,
        TipoPlanAnterior NVARCHAR(50) NULL,
        TipoPlanNuevo NVARCHAR(50) NOT NULL,
        EstadoSuscripcionAnterior NVARCHAR(20) NULL,
        EstadoSuscripcionNuevo NVARCHAR(20) NOT NULL,
        EsPruebaAnterior BIT NULL,
        EsPruebaNuevo BIT NOT NULL,
        TipoCobroAnterior NVARCHAR(20) NULL,
        TipoCobroNuevo NVARCHAR(20) NULL,
        FechaInicioReferencia DATE NULL,
        FechaFinReferencia DATE NULL,
        DiasGracia INT NULL,
        DiasExtra INT NULL,
        EmpresasPermitidasAnterior INT NULL,
        EmpresasPermitidasNuevo INT NULL,
        UsuariosPermitidosAnterior INT NULL,
        UsuariosPermitidosNuevo INT NULL,
        Observacion NVARCHAR(500) NULL,
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_SEG_CuentaAdministradoraSuscripcionMovimiento_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    CREATE NONCLUSTERED INDEX IX_SEG_CuentaAdministradoraSuscripcionMovimiento_Cuenta_Fecha
        ON dbo.SEG_CuentaAdministradoraSuscripcionMovimiento (IdCuentaAdministradora ASC, FechaRegistro DESC);

    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcionMovimiento
        ADD CONSTRAINT FK_SEG_CuentaAdministradoraSuscripcionMovimiento_SEG_CuentaAdministradora
        FOREIGN KEY (IdCuentaAdministradora) REFERENCES dbo.SEG_CuentaAdministradora (IdCuentaAdministradora);

    ALTER TABLE dbo.SEG_CuentaAdministradoraSuscripcionMovimiento
        ADD CONSTRAINT FK_SEG_CuentaAdministradoraSuscripcionMovimiento_SEG_CuentaAdministradoraSuscripcion
        FOREIGN KEY (IdCuentaAdministradoraSuscripcion) REFERENCES dbo.SEG_CuentaAdministradoraSuscripcion (IdCuentaAdministradoraSuscripcion);
END;
