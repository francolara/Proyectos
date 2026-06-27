-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Cabecera de configuracion contable automatica por empresa, modulo y escenario operativo.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Agrega escenario PROVISION para configuracion contable directa.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Amplia los modulos permitidos para provisiones futuras de egresos, ingresos y aplicaciones NC.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/06/2026
-- Description:   Incorpora el modulo DET para la provision automatica de detracciones en compras.
-- =============================================

IF OBJECT_ID(N'dbo.CON_ConfiguracionContabilizacion', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_ConfiguracionContabilizacion
    (
        IdConfiguracionContabilizacion INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_ConfiguracionContabilizacion PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        ModuloOperacion VARCHAR(10) NOT NULL,
        EscenarioOperacion VARCHAR(20) NOT NULL,
        IdOrigen INT NOT NULL,
        Descripcion NVARCHAR(200) NOT NULL,
        GeneraAsientoAutomatico BIT NOT NULL CONSTRAINT DF_CON_ConfiguracionContabilizacion_GeneraAsientoAutomatico DEFAULT (1),
        UsaTipoCambio BIT NOT NULL CONSTRAINT DF_CON_ConfiguracionContabilizacion_UsaTipoCambio DEFAULT (1),
        Activo BIT NOT NULL CONSTRAINT DF_CON_ConfiguracionContabilizacion_Activo DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_ConfiguracionContabilizacion_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_ConfiguracionContabilizacion
        ADD CONSTRAINT FK_CON_ConfiguracionContabilizacion_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);

    ALTER TABLE dbo.CON_ConfiguracionContabilizacion
        ADD CONSTRAINT FK_CON_ConfiguracionContabilizacion_CON_Origen
            FOREIGN KEY (IdOrigen) REFERENCES dbo.CON_Origen (IdOrigen);

    ALTER TABLE dbo.CON_ConfiguracionContabilizacion
        ADD CONSTRAINT CK_CON_ConfiguracionContabilizacion_ModuloOperacion
            CHECK (ModuloOperacion IN ('COM', 'VEN', 'EGR', 'ING', 'APNC', 'DET'));

    ALTER TABLE dbo.CON_ConfiguracionContabilizacion
        ADD CONSTRAINT CK_CON_ConfiguracionContabilizacion_EscenarioOperacion
            CHECK (EscenarioOperacion IN ('MERCADERIA', 'GASTO', 'SERVICIO', 'PROVISION'));

    ALTER TABLE dbo.CON_ConfiguracionContabilizacion
        ADD CONSTRAINT UQ_CON_ConfiguracionContabilizacion
            UNIQUE (IdEmpresa, ModuloOperacion, EscenarioOperacion);
END;
