-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Agrega el enlace funcional de transferencias entre cuentas en BAN_MovimientoBanco y crea los procedimientos para guardar, listar y eliminar la transferencia completa.
-- =============================================

IF COL_LENGTH(N'dbo.BAN_MovimientoBanco', N'IdTransferenciaCuenta') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBanco
        ADD IdTransferenciaCuenta UNIQUEIDENTIFIER NULL;
END;

IF COL_LENGTH(N'dbo.BAN_MovimientoBanco', N'RolTransferencia') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBanco
        ADD RolTransferencia CHAR(1) NULL;
END;

IF COL_LENGTH(N'dbo.BAN_MovimientoBanco', N'IdMovimientoBancoRelacionado') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBanco
        ADD IdMovimientoBancoRelacionado INT NULL;
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.check_constraints
    WHERE name = N'CK_BAN_MovimientoBanco_RolTransferencia'
)
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBanco
        ADD CONSTRAINT CK_BAN_MovimientoBanco_RolTransferencia
            CHECK (RolTransferencia IS NULL OR RolTransferencia IN ('E', 'I'));
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.foreign_keys
    WHERE name = N'FK_BAN_MovimientoBanco_MovimientoRelacionado'
)
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBanco
        ADD CONSTRAINT FK_BAN_MovimientoBanco_MovimientoRelacionado
            FOREIGN KEY (IdMovimientoBancoRelacionado) REFERENCES dbo.BAN_MovimientoBanco (IdMovimientoBanco);
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.indexes
    WHERE name = N'IX_BAN_MovimientoBanco_IdTransferenciaCuenta'
      AND object_id = OBJECT_ID(N'dbo.BAN_MovimientoBanco')
)
BEGIN
    CREATE INDEX IX_BAN_MovimientoBanco_IdTransferenciaCuenta
        ON dbo.BAN_MovimientoBanco (IdTransferenciaCuenta, RolTransferencia);
END;
