-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Agrega campos documentarios opcionales al detalle del asiento manual.
-- =============================================

IF COL_LENGTH('dbo.CON_AsientoDetalle', 'CodigoCentroCosto') IS NULL
BEGIN
    ALTER TABLE dbo.CON_AsientoDetalle
        ADD CodigoCentroCosto NVARCHAR(50) NULL;
END;

IF COL_LENGTH('dbo.CON_AsientoDetalle', 'TipoDocumento') IS NULL
BEGIN
    ALTER TABLE dbo.CON_AsientoDetalle
        ADD TipoDocumento VARCHAR(3) NULL;
END;

IF COL_LENGTH('dbo.CON_AsientoDetalle', 'NumeroDocumento') IS NULL
BEGIN
    ALTER TABLE dbo.CON_AsientoDetalle
        ADD NumeroDocumento VARCHAR(20) NULL;
END;

IF COL_LENGTH('dbo.CON_AsientoDetalle', 'Serie') IS NULL
BEGIN
    ALTER TABLE dbo.CON_AsientoDetalle
        ADD Serie VARCHAR(10) NULL;
END;

IF COL_LENGTH('dbo.CON_AsientoDetalle', 'TipoCambioLinea') IS NULL
BEGIN
    ALTER TABLE dbo.CON_AsientoDetalle
        ADD TipoCambioLinea DECIMAL(18,6) NULL;
END;
