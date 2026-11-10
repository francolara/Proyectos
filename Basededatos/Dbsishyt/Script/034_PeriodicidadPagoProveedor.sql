-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Agrega frecuencia y periodo de cobertura a los abonos de proveedor.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/11/2026
-- Description:   Permite guardar frecuencia sin fechas de cobertura y conserva los rangos historicos.
-- =============================================
USE [SistemaWebHYT];
GO
IF COL_LENGTH(N'dbo.PagoProveedorHYT', N'Frecuencia') IS NULL
BEGIN
    ALTER TABLE dbo.PagoProveedorHYT ADD Frecuencia VARCHAR(10) NULL;
    ALTER TABLE dbo.PagoProveedorHYT ADD PeriodoDesde DATE NULL;
    ALTER TABLE dbo.PagoProveedorHYT ADD PeriodoHasta DATE NULL;
END;
GO
IF OBJECT_ID(N'dbo.CK_PagoProveedorHYT_CoberturaOpcional', N'C') IS NULL
BEGIN
    ALTER TABLE dbo.PagoProveedorHYT WITH CHECK ADD CONSTRAINT CK_PagoProveedorHYT_CoberturaOpcional CHECK (
        (Frecuencia IS NULL OR Frecuencia IN ('SEMANAL', 'QUINCENAL', 'MENSUAL'))
        AND ((PeriodoDesde IS NULL AND PeriodoHasta IS NULL)
             OR (Frecuencia IS NOT NULL AND PeriodoDesde IS NOT NULL AND PeriodoHasta IS NOT NULL
                 AND PeriodoDesde <= PeriodoHasta)));
END;
IF OBJECT_ID(N'dbo.CK_PagoProveedorHYT_Periodo', N'C') IS NOT NULL
    ALTER TABLE dbo.PagoProveedorHYT DROP CONSTRAINT CK_PagoProveedorHYT_Periodo;
GO
