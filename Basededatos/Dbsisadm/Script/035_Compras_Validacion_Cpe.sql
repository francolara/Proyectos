-- =============================================
-- Author:        FRANCO LARA
-- Create date:   29/06/2026
-- Description:   Agrega columnas de fecha, estado y mensaje para registrar la validacion CPE de compras.
-- =============================================

IF COL_LENGTH(N'dbo.COM_Compra', N'FechaValidacionCpe') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD FechaValidacionCpe DATETIME2(0) NULL;
END;

IF COL_LENGTH(N'dbo.COM_Compra', N'EstadoValidacionCpe') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD EstadoValidacionCpe NVARCHAR(50) NULL;
END;

IF COL_LENGTH(N'dbo.COM_Compra', N'MensajeValidacionCpe') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD MensajeValidacionCpe NVARCHAR(500) NULL;
END;
