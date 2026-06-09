

-- Firma: FRANCO LARA - 09/06/2026 | Agrega HorasMaximasReservaCliente en Negocios para controlar el limite configurable de duracion en la reserva publica.

IF COL_LENGTH('dbo.Negocios', 'HorasMaximasReservaCliente') IS NULL
BEGIN
    ALTER TABLE dbo.Negocios
        ADD HorasMaximasReservaCliente INT NOT NULL
            CONSTRAINT DF_Negocios_HorasMaximasReservaCliente DEFAULT (1);
END;

IF COL_LENGTH('dbo.Negocios', 'HorasMaximasReservaCliente') IS NOT NULL
   AND NOT EXISTS (
       SELECT 1
       FROM sys.default_constraints dc
       INNER JOIN sys.columns c
           ON c.object_id = dc.parent_object_id
          AND c.column_id = dc.parent_column_id
       WHERE dc.parent_object_id = OBJECT_ID('dbo.Negocios')
         AND c.name = 'HorasMaximasReservaCliente'
   )
BEGIN
    ALTER TABLE dbo.Negocios
        ADD CONSTRAINT DF_Negocios_HorasMaximasReservaCliente DEFAULT (1) FOR HorasMaximasReservaCliente;
END;

IF EXISTS (
    SELECT 1
    FROM sys.columns
    WHERE object_id = OBJECT_ID('dbo.Negocios')
      AND name = 'HorasMaximasReservaCliente'
      AND is_nullable = 1
)
BEGIN
    UPDATE dbo.Negocios
    SET HorasMaximasReservaCliente = 1
    WHERE HorasMaximasReservaCliente IS NULL
       OR HorasMaximasReservaCliente < 1
       OR HorasMaximasReservaCliente > 12;

    ALTER TABLE dbo.Negocios
        ALTER COLUMN HorasMaximasReservaCliente INT NOT NULL;
END;

UPDATE dbo.Negocios
SET HorasMaximasReservaCliente = 1
WHERE HorasMaximasReservaCliente IS NULL
   OR HorasMaximasReservaCliente < 1
   OR HorasMaximasReservaCliente > 12;

IF EXISTS (
    SELECT 1
    FROM sys.check_constraints
    WHERE name = 'CK_Negocios_HorasMaximasReservaCliente'
      AND parent_object_id = OBJECT_ID('dbo.Negocios')
)
BEGIN
    ALTER TABLE dbo.Negocios DROP CONSTRAINT CK_Negocios_HorasMaximasReservaCliente;
END;

IF NOT EXISTS (
    SELECT 1
    FROM sys.check_constraints
    WHERE name = 'CK_Negocios_HorasMaximasReservaCliente'
      AND parent_object_id = OBJECT_ID('dbo.Negocios')
)
BEGIN
    ALTER TABLE dbo.Negocios WITH CHECK
        ADD CONSTRAINT CK_Negocios_HorasMaximasReservaCliente
            CHECK (HorasMaximasReservaCliente BETWEEN 1 AND 12);
END;
