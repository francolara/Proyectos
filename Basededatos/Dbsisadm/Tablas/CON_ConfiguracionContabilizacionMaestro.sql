-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   25/08/2026
-- Description:   Configuracion contable maestra por modulo y escenario, vinculada a origenes mediante CodigoOrigen portable entre empresas.
-- =============================================

IF OBJECT_ID(N'dbo.CON_ConfiguracionContabilizacionMaestro', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_ConfiguracionContabilizacionMaestro
    (
        IdConfiguracionContabilizacionMaestro INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_ConfiguracionContabilizacionMaestro PRIMARY KEY,
        ModuloOperacion VARCHAR(10) NOT NULL,
        EscenarioOperacion VARCHAR(20) NOT NULL,
        CodigoOrigen VARCHAR(10) NOT NULL,
        Descripcion NVARCHAR(200) NOT NULL,
        GeneraAsientoAutomatico BIT NOT NULL CONSTRAINT DF_CON_ConfiguracionContabilizacionMaestro_GeneraAsientoAutomatico DEFAULT (1),
        UsaTipoCambio BIT NOT NULL CONSTRAINT DF_CON_ConfiguracionContabilizacionMaestro_UsaTipoCambio DEFAULT (1),
        Activo BIT NOT NULL CONSTRAINT DF_CON_ConfiguracionContabilizacionMaestro_Activo DEFAULT (1),
        Orden INT NOT NULL CONSTRAINT DF_CON_ConfiguracionContabilizacionMaestro_Orden DEFAULT (0),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_ConfiguracionContabilizacionMaestro_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_ConfiguracionContabilizacionMaestro
        ADD CONSTRAINT CK_CON_ConfiguracionContabilizacionMaestro_ModuloOperacion
            CHECK (ModuloOperacion IN ('COM', 'VEN', 'EGR', 'ING', 'APNC', 'DET', 'PER', 'DIF', 'AJU', 'APR', 'CIE'));

    ALTER TABLE dbo.CON_ConfiguracionContabilizacionMaestro
        ADD CONSTRAINT CK_CON_ConfiguracionContabilizacionMaestro_EscenarioOperacion
            CHECK (EscenarioOperacion IN ('MERCADERIA', 'GASTO', 'SERVICIO', 'PROVISION'));

    ALTER TABLE dbo.CON_ConfiguracionContabilizacionMaestro
        ADD CONSTRAINT UQ_CON_ConfiguracionContabilizacionMaestro
            UNIQUE (ModuloOperacion, EscenarioOperacion);
END;

IF OBJECT_ID(N'dbo.CON_OrigenMaestro', N'U') IS NOT NULL
   AND NOT EXISTS
   (
       SELECT 1
       FROM sys.foreign_keys AS fk
       WHERE fk.name = N'FK_CON_ConfiguracionContabilizacionMaestro_CON_OrigenMaestro'
         AND fk.parent_object_id = OBJECT_ID(N'dbo.CON_ConfiguracionContabilizacionMaestro')
   )
BEGIN
    ALTER TABLE dbo.CON_ConfiguracionContabilizacionMaestro
        ADD CONSTRAINT FK_CON_ConfiguracionContabilizacionMaestro_CON_OrigenMaestro
            FOREIGN KEY (CodigoOrigen) REFERENCES dbo.CON_OrigenMaestro (CodigoOrigen);
END;

