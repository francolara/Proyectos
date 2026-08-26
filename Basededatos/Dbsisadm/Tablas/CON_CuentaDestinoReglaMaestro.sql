-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Cabecera maestra interna de reglas de cuentas destino. No pertenece a una empresa.
-- =============================================
-- Firma: FRANCO LARA - 25/08/2026 | Elimina Ejercicio, mantiene una regla por CodigoCuentaOrigen y define el identity desde cero.

IF OBJECT_ID(N'dbo.CON_CuentaDestinoReglaMaestro', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_CuentaDestinoReglaMaestro
    (
        IdCuentaDestinoReglaMaestro INT IDENTITY(0,1) NOT NULL CONSTRAINT PK_CON_CuentaDestinoReglaMaestro PRIMARY KEY,
        CodigoCuentaOrigen VARCHAR(20) NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_CON_CuentaDestinoReglaMaestro_Activo DEFAULT (1),
        Observacion NVARCHAR(500) NULL,
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_CuentaDestinoReglaMaestro_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_CuentaDestinoReglaMaestro
        ADD CONSTRAINT UQ_CON_CuentaDestinoReglaMaestro
            UNIQUE (CodigoCuentaOrigen);
END;

IF COL_LENGTH(N'dbo.CON_CuentaDestinoReglaMaestro', N'Ejercicio') IS NOT NULL
BEGIN
    EXEC sys.sp_executesql N'
        ;WITH ReglasDuplicadas AS
        (
            SELECT
                regla.IdCuentaDestinoReglaMaestro,
                ROW_NUMBER() OVER
                (
                    PARTITION BY regla.CodigoCuentaOrigen
                    ORDER BY regla.Ejercicio DESC, regla.IdCuentaDestinoReglaMaestro DESC
                ) AS NumeroFila
            FROM dbo.CON_CuentaDestinoReglaMaestro AS regla
        )
        DELETE detalle
        FROM dbo.CON_CuentaDestinoReglaDetalleMaestro AS detalle
        INNER JOIN ReglasDuplicadas AS duplicada
            ON duplicada.IdCuentaDestinoReglaMaestro = detalle.IdCuentaDestinoReglaMaestro
        WHERE duplicada.NumeroFila > 1;

        ;WITH ReglasDuplicadas AS
        (
            SELECT
                regla.IdCuentaDestinoReglaMaestro,
                ROW_NUMBER() OVER
                (
                    PARTITION BY regla.CodigoCuentaOrigen
                    ORDER BY regla.Ejercicio DESC, regla.IdCuentaDestinoReglaMaestro DESC
                ) AS NumeroFila
            FROM dbo.CON_CuentaDestinoReglaMaestro AS regla
        )
        DELETE regla
        FROM dbo.CON_CuentaDestinoReglaMaestro AS regla
        INNER JOIN ReglasDuplicadas AS duplicada
            ON duplicada.IdCuentaDestinoReglaMaestro = regla.IdCuentaDestinoReglaMaestro
        WHERE duplicada.NumeroFila > 1;';

    IF EXISTS
    (
        SELECT 1
        FROM sys.key_constraints AS restriccion
        WHERE restriccion.name = N'UQ_CON_CuentaDestinoReglaMaestro'
          AND restriccion.parent_object_id = OBJECT_ID(N'dbo.CON_CuentaDestinoReglaMaestro')
    )
    BEGIN
        ALTER TABLE dbo.CON_CuentaDestinoReglaMaestro
            DROP CONSTRAINT UQ_CON_CuentaDestinoReglaMaestro;
    END;

    EXEC sys.sp_executesql N'
        ALTER TABLE dbo.CON_CuentaDestinoReglaMaestro DROP COLUMN Ejercicio;';
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.key_constraints AS restriccion
    WHERE restriccion.name = N'UQ_CON_CuentaDestinoReglaMaestro'
      AND restriccion.parent_object_id = OBJECT_ID(N'dbo.CON_CuentaDestinoReglaMaestro')
)
BEGIN
    ALTER TABLE dbo.CON_CuentaDestinoReglaMaestro
        ADD CONSTRAINT UQ_CON_CuentaDestinoReglaMaestro
            UNIQUE (CodigoCuentaOrigen);
END;
