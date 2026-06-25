-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Agrega la persona por linea al detalle de Caja y Bancos para enlazar comprobantes por cada registro.
-- =============================================

IF COL_LENGTH('dbo.BAN_MovimientoBancoDetalle', 'IdPersona') IS NULL
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBancoDetalle
    ADD IdPersona INT NULL;
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.foreign_keys
    WHERE name = 'FK_BAN_MovimientoBancoDetalle_ADM_Persona'
)
BEGIN
    ALTER TABLE dbo.BAN_MovimientoBancoDetalle
        ADD CONSTRAINT FK_BAN_MovimientoBancoDetalle_ADM_Persona
            FOREIGN KEY (IdPersona) REFERENCES dbo.ADM_Persona (IdPersona);
END;
