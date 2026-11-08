-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/11/2026
-- Description:   Subtítulo configurable de cotización con texto inicial indicado por el usuario.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF COL_LENGTH(N'dbo.ConfiguracionHYT',N'SubtituloCotizacion') IS NULL
    ALTER TABLE dbo.ConfiguracionHYT ADD SubtituloCotizacion NVARCHAR(2000) NULL CONSTRAINT DF_ConfiguracionHYT_Subtitulo DEFAULT(N'GASFITERÍA, EBANISTERÍA, ELECTRICIDAD, PINTURA LÁTEX Y LAQUEADOS, TRABAJOS EN DRYWALL, FALSOS CIELOS,  CARPINTERÍA METÁLICA, REMODELACIONES, DEMOLICIONES Y ABASTECIMIENTO DE COMBUSTIBLE / R.P.C: 991346699 E-mail: hytserviciosg@gmail.com - rene_torres35@hotmail.com') WITH VALUES;
GO
