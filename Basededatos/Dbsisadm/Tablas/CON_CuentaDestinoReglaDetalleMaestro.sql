-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Detalle maestro interno de cuentas destino. No pertenece a una empresa.
-- =============================================

IF OBJECT_ID(N'dbo.CON_CuentaDestinoReglaDetalleMaestro', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_CuentaDestinoReglaDetalleMaestro
    (
        IdCuentaDestinoReglaDetalleMaestro INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_CuentaDestinoReglaDetalleMaestro PRIMARY KEY,
        IdCuentaDestinoReglaMaestro INT NOT NULL,
        Orden SMALLINT NOT NULL,
        CodigoCuentaDestinoCargo VARCHAR(20) NOT NULL,
        CodigoCuentaDestinoAbono VARCHAR(20) NOT NULL,
        Porcentaje DECIMAL(7,4) NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_CON_CuentaDestinoReglaDetalleMaestro_Activo DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_CuentaDestinoReglaDetalleMaestro_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_CuentaDestinoReglaDetalleMaestro
        ADD CONSTRAINT FK_CON_CuentaDestinoReglaDetalleMaestro_Cabecera
            FOREIGN KEY (IdCuentaDestinoReglaMaestro) REFERENCES dbo.CON_CuentaDestinoReglaMaestro (IdCuentaDestinoReglaMaestro);

    ALTER TABLE dbo.CON_CuentaDestinoReglaDetalleMaestro
        ADD CONSTRAINT UQ_CON_CuentaDestinoReglaDetalleMaestro
            UNIQUE (IdCuentaDestinoReglaMaestro, Orden);
END;
