-- =============================================
-- Author:        FRANCO LARA
-- Create date:   25/06/2026
-- Description:   Unifica las cuentas destino por empresa y cuenta origen, eliminando duplicados por ejercicio y recreando la clave unica.
-- =============================================

IF OBJECT_ID(N'dbo.CON_CuentaDestinoRegla', N'U') IS NOT NULL
BEGIN
    IF EXISTS
    (
        SELECT 1
        FROM sys.key_constraints
        WHERE [type] = 'UQ'
          AND [name] = 'UQ_CON_CuentaDestinoRegla'
          AND parent_object_id = OBJECT_ID(N'dbo.CON_CuentaDestinoRegla')
    )
    BEGIN
        ALTER TABLE dbo.CON_CuentaDestinoRegla
            DROP CONSTRAINT UQ_CON_CuentaDestinoRegla;
    END;

    ;WITH Duplicados AS
    (
        SELECT
            r.IdCuentaDestinoRegla,
            ROW_NUMBER() OVER
            (
                PARTITION BY r.IdEmpresa, r.IdPlanCuentaOrigen
                ORDER BY r.FechaRegistro DESC, r.IdCuentaDestinoRegla DESC
            ) AS OrdenDuplicado
        FROM dbo.CON_CuentaDestinoRegla AS r
    )
    DELETE d
    FROM dbo.CON_CuentaDestinoReglaDetalle AS d
    INNER JOIN Duplicados AS x
        ON x.IdCuentaDestinoRegla = d.IdCuentaDestinoRegla
    WHERE x.OrdenDuplicado > 1;

    ;WITH Duplicados AS
    (
        SELECT
            r.IdCuentaDestinoRegla,
            ROW_NUMBER() OVER
            (
                PARTITION BY r.IdEmpresa, r.IdPlanCuentaOrigen
                ORDER BY r.FechaRegistro DESC, r.IdCuentaDestinoRegla DESC
            ) AS OrdenDuplicado
        FROM dbo.CON_CuentaDestinoRegla AS r
    )
    DELETE r
    FROM dbo.CON_CuentaDestinoRegla AS r
    INNER JOIN Duplicados AS x
        ON x.IdCuentaDestinoRegla = r.IdCuentaDestinoRegla
    WHERE x.OrdenDuplicado > 1;

    IF NOT EXISTS
    (
        SELECT 1
        FROM sys.key_constraints
        WHERE [type] = 'UQ'
          AND [name] = 'UQ_CON_CuentaDestinoRegla'
          AND parent_object_id = OBJECT_ID(N'dbo.CON_CuentaDestinoRegla')
    )
    BEGIN
        ALTER TABLE dbo.CON_CuentaDestinoRegla
            ADD CONSTRAINT UQ_CON_CuentaDestinoRegla
                UNIQUE (IdEmpresa, IdPlanCuentaOrigen);
    END;
END;
