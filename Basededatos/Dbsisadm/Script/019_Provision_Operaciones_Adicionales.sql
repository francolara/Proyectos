-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Amplia la configuracion de provision para compras, ventas, egresos, ingresos y aplicaciones NC.
-- =============================================

IF OBJECT_ID(N'dbo.CK_CON_ConfiguracionContabilizacion_ModuloOperacion', N'C') IS NOT NULL
BEGIN
    ALTER TABLE dbo.CON_ConfiguracionContabilizacion
        DROP CONSTRAINT CK_CON_ConfiguracionContabilizacion_ModuloOperacion;
END;

ALTER TABLE dbo.CON_ConfiguracionContabilizacion
    ADD CONSTRAINT CK_CON_ConfiguracionContabilizacion_ModuloOperacion
        CHECK (ModuloOperacion IN ('COM', 'VEN', 'EGR', 'ING', 'APNC'));
