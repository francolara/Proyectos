-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/07/2026
-- Description:   Agrega y rellena el Periodo persistido de BAN_MovimientoBanco desde FechaEmision para listar y resumir Caja y Bancos por periodo contable grabado.
-- =============================================

IF COL_LENGTH(N'dbo.BAN_MovimientoBanco', N'Periodo') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBanco
        ADD Periodo CHAR(6) NULL;
END;

UPDATE dbo.BAN_MovimientoBanco
SET Periodo = CONVERT(CHAR(4), YEAR(FechaEmision)) + RIGHT('0' + CONVERT(VARCHAR(2), MONTH(FechaEmision)), 2)
WHERE Periodo IS NULL;

IF EXISTS
(
    SELECT 1
    FROM sys.columns
    WHERE object_id = OBJECT_ID(N'dbo.BAN_MovimientoBanco')
      AND name = N'Periodo'
      AND is_nullable = 1
)
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBanco
        ALTER COLUMN Periodo CHAR(6) NOT NULL;
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.check_constraints
    WHERE name = N'CK_BAN_MovimientoBanco_Periodo'
)
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBanco
        ADD CONSTRAINT CK_BAN_MovimientoBanco_Periodo
            CHECK (
                Periodo LIKE '[1-2][0-9][0-9][0-9][0-1][0-9]'
                AND RIGHT(Periodo, 2) BETWEEN '01' AND '12'
            );
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.indexes
    WHERE name = N'IX_BAN_MovimientoBanco_IdEmpresa_Periodo'
      AND object_id = OBJECT_ID(N'dbo.BAN_MovimientoBanco')
)
BEGIN
    CREATE INDEX IX_BAN_MovimientoBanco_IdEmpresa_Periodo
        ON dbo.BAN_MovimientoBanco (IdEmpresa, Periodo, IdBancoConfiguracionEmpresa, Activo);
END;
