USE [DbSportCenter];

-- Firma: Codex - 06/04/2026 | Estructura para politica de confirmacion de reservas por pago y porcentaje minimo de adelanto por negocio.

IF COL_LENGTH('dbo.Negocios', 'PoliticaConfirmacionPago') IS NULL
BEGIN
    ALTER TABLE dbo.Negocios
        ADD PoliticaConfirmacionPago TINYINT NOT NULL
            CONSTRAINT DF_Negocios_PoliticaConfirmacionPago DEFAULT (0);
END;

IF COL_LENGTH('dbo.Negocios', 'PoliticaConfirmacionPago') IS NOT NULL
   AND NOT EXISTS (
       SELECT 1
       FROM sys.default_constraints dc
       INNER JOIN sys.columns c
           ON c.object_id = dc.parent_object_id
          AND c.column_id = dc.parent_column_id
       WHERE dc.parent_object_id = OBJECT_ID('dbo.Negocios')
         AND c.name = 'PoliticaConfirmacionPago'
   )
BEGIN
    ALTER TABLE dbo.Negocios
        ADD CONSTRAINT DF_Negocios_PoliticaConfirmacionPago DEFAULT (0) FOR PoliticaConfirmacionPago;
END;

IF COL_LENGTH('dbo.Negocios', 'PorcentajeAdelantoMinimo') IS NULL
BEGIN
    ALTER TABLE dbo.Negocios
        ADD PorcentajeAdelantoMinimo DECIMAL(5,2) NULL;
END;

IF EXISTS (
    SELECT 1
    FROM sys.columns
    WHERE object_id = OBJECT_ID('dbo.Negocios')
      AND name = 'PoliticaConfirmacionPago'
      AND is_nullable = 1
)
BEGIN
    UPDATE dbo.Negocios
    SET PoliticaConfirmacionPago = 0
    WHERE PoliticaConfirmacionPago IS NULL;

    ALTER TABLE dbo.Negocios
        ALTER COLUMN PoliticaConfirmacionPago TINYINT NOT NULL;
END;

UPDATE dbo.Negocios
SET PoliticaConfirmacionPago = 0
WHERE PoliticaConfirmacionPago NOT IN (0, 1, 2)
   OR PoliticaConfirmacionPago IS NULL;

UPDATE dbo.Negocios
SET PorcentajeAdelantoMinimo = NULL
WHERE PorcentajeAdelantoMinimo <= 0
   OR PorcentajeAdelantoMinimo > 100
   OR PorcentajeAdelantoMinimo <> FLOOR(PorcentajeAdelantoMinimo);

IF NOT EXISTS (
    SELECT 1
    FROM sys.check_constraints
    WHERE name = 'CK_Negocios_PoliticaConfirmacionPago'
      AND parent_object_id = OBJECT_ID('dbo.Negocios')
)
BEGIN
    ALTER TABLE dbo.Negocios WITH CHECK
        ADD CONSTRAINT CK_Negocios_PoliticaConfirmacionPago
            CHECK (PoliticaConfirmacionPago IN (0, 1, 2));
END;

IF EXISTS (
    SELECT 1
    FROM sys.check_constraints
    WHERE name = 'CK_Negocios_PorcentajeAdelantoMinimo'
      AND parent_object_id = OBJECT_ID('dbo.Negocios')
)
BEGIN
    ALTER TABLE dbo.Negocios DROP CONSTRAINT CK_Negocios_PorcentajeAdelantoMinimo;
END;

IF NOT EXISTS (
    SELECT 1
    FROM sys.check_constraints
    WHERE name = 'CK_Negocios_PorcentajeAdelantoMinimo'
      AND parent_object_id = OBJECT_ID('dbo.Negocios')
)
BEGIN
    ALTER TABLE dbo.Negocios WITH CHECK
        ADD CONSTRAINT CK_Negocios_PorcentajeAdelantoMinimo
            CHECK (PorcentajeAdelantoMinimo IS NULL OR (PorcentajeAdelantoMinimo >= 1 AND PorcentajeAdelantoMinimo <= 100 AND PorcentajeAdelantoMinimo = FLOOR(PorcentajeAdelantoMinimo)));
END;
