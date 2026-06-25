-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Cabecera maestra interna de reglas de cuentas destino. No pertenece a una empresa.
-- =============================================

IF OBJECT_ID(N'dbo.CON_CuentaDestinoReglaMaestro', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_CuentaDestinoReglaMaestro
    (
        IdCuentaDestinoReglaMaestro INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_CuentaDestinoReglaMaestro PRIMARY KEY,
        Ejercicio SMALLINT NOT NULL,
        CodigoCuentaOrigen VARCHAR(20) NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_CON_CuentaDestinoReglaMaestro_Activo DEFAULT (1),
        Observacion NVARCHAR(500) NULL,
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_CuentaDestinoReglaMaestro_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_CuentaDestinoReglaMaestro
        ADD CONSTRAINT UQ_CON_CuentaDestinoReglaMaestro
            UNIQUE (Ejercicio, CodigoCuentaOrigen);
END;
