-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Agrega TipoCambio y Observacion a la cabecera de Caja y Bancos e inicializa TipoCambio con 1 para registros existentes.
-- =============================================

IF COL_LENGTH('dbo.BAN_MovimientoBanco', 'TipoCambio') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBanco
    ADD TipoCambio DECIMAL(18, 6) NOT NULL
        CONSTRAINT DF_BAN_MovimientoBanco_TipoCambio DEFAULT (1);
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.check_constraints
    WHERE name = 'CK_BAN_MovimientoBanco_TipoCambio'
)
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBanco
        ADD CONSTRAINT CK_BAN_MovimientoBanco_TipoCambio
            CHECK (TipoCambio > 0);
END;

IF COL_LENGTH('dbo.BAN_MovimientoBanco', 'Observacion') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBanco
    ADD Observacion NVARCHAR(500) NULL;
END;

UPDATE dbo.BAN_MovimientoBanco
SET TipoCambio = 1
WHERE TipoCambio IS NULL;
