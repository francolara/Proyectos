-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   23/06/2026
-- Description:   Agrega correlativo interno por periodo a Caja y Bancos y rellena datos existentes por empresa y mes de fecha emision.
-- =============================================

IF COL_LENGTH('dbo.BAN_MovimientoBanco', 'NumeroMovimiento') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBanco
        ADD NumeroMovimiento INT NULL;
END;

;WITH Correlativos AS
(
    SELECT
        m.IdMovimientoBanco,
        ROW_NUMBER() OVER
        (
            PARTITION BY
                m.IdEmpresa,
                YEAR(m.FechaEmision),
                MONTH(m.FechaEmision)
            ORDER BY m.FechaEmision, m.IdMovimientoBanco
        ) AS NumeroMovimientoNuevo
    FROM dbo.BAN_MovimientoBanco AS m
)
UPDATE m
SET NumeroMovimiento = c.NumeroMovimientoNuevo
FROM dbo.BAN_MovimientoBanco AS m
INNER JOIN Correlativos AS c
    ON c.IdMovimientoBanco = m.IdMovimientoBanco
WHERE m.NumeroMovimiento IS NULL
   OR m.NumeroMovimiento <= 0;

IF EXISTS
(
    SELECT 1
    FROM sys.check_constraints
    WHERE name = 'CK_BAN_MovimientoBanco_NumeroMovimiento'
)
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBanco
        DROP CONSTRAINT CK_BAN_MovimientoBanco_NumeroMovimiento;
END;

ALTER TABLE dbo.BAN_MovimientoBanco
    ALTER COLUMN NumeroMovimiento INT NOT NULL;

ALTER TABLE dbo.BAN_MovimientoBanco
    ADD CONSTRAINT CK_BAN_MovimientoBanco_NumeroMovimiento
        CHECK (NumeroMovimiento > 0);
