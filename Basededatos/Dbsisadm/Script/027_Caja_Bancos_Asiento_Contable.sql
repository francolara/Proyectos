-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Agrega IdAsiento en BAN_MovimientoBanco para vincular el movimiento bancario con su asiento contable automatico.
-- =============================================

IF COL_LENGTH(N'dbo.BAN_MovimientoBanco', N'IdAsiento') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBanco
        ADD IdAsiento INT NULL;
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.foreign_keys
    WHERE name = N'FK_BAN_MovimientoBanco_CON_Asiento'
)
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBanco
        ADD CONSTRAINT FK_BAN_MovimientoBanco_CON_Asiento
            FOREIGN KEY (IdAsiento) REFERENCES dbo.CON_Asiento (IdAsiento);
END;
