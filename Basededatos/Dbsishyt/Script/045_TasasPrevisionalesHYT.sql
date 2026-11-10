-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/11/2026
-- Description:   Agrega porcentajes de estimacion al catalogo previsional.
-- =============================================
USE [SistemaWebHYT];
GO
IF COL_LENGTH(N'dbo.EntidadPrevisionalHYT', N'PorcentajeEstimado') IS NULL
BEGIN
    ALTER TABLE dbo.EntidadPrevisionalHYT ADD PorcentajeEstimado DECIMAL(7,4) NULL;
END;
GO
-- AFP: aporte al fondo segun SBS. ONP: tasa publicada por SUNAT.
UPDATE dbo.EntidadPrevisionalHYT
SET PorcentajeEstimado = CASE WHEN Regimen = 'AFP' THEN 10.0000 ELSE 13.0000 END
WHERE PorcentajeEstimado IS NULL;
GO
IF OBJECT_ID(N'dbo.CK_EntidadPrevisionalHYT_Porcentaje', N'C') IS NULL
    ALTER TABLE dbo.EntidadPrevisionalHYT ADD CONSTRAINT CK_EntidadPrevisionalHYT_Porcentaje
        CHECK (PorcentajeEstimado > 0 AND PorcentajeEstimado <= 100);
GO
