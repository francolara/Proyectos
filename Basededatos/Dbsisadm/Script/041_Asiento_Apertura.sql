-- =============================================
-- Author:        FRANCO LARA
-- Create date:   02/07/2026
-- Description:   Habilita el modulo APR y sus objetos base para asiento de apertura.
-- =============================================
-- Firma: FRANCO LARA - 02/07/2026 | Agrega el modulo APR en configuracion contable, parametriza el origen sugerido 00 y prepara las estructuras persistentes del proceso de asiento de apertura.

IF OBJECT_ID(N'dbo.CK_CON_ConfiguracionContabilizacion_ModuloOperacion', N'C') IS NOT NULL
BEGIN
    ALTER TABLE dbo.CON_ConfiguracionContabilizacion
        DROP CONSTRAINT CK_CON_ConfiguracionContabilizacion_ModuloOperacion;
END;

ALTER TABLE dbo.CON_ConfiguracionContabilizacion
    ADD CONSTRAINT CK_CON_ConfiguracionContabilizacion_ModuloOperacion
        CHECK (ModuloOperacion IN ('COM', 'VEN', 'EGR', 'ING', 'APNC', 'DET', 'PER', 'DIF', 'AJU', 'APR', 'CIE'));

MERGE dbo.ADM_ParametroMaestro AS destino
USING
(
    SELECT
        CAST('CONTABLE' AS VARCHAR(50)) AS TipoParametro,
        CAST('ORIGEN_ASIENTO_APERTURA' AS VARCHAR(50)) AS CodigoParametro,
        CAST(N'00' AS NVARCHAR(200)) AS ValorParametro,
        CAST(N'Origen sugerido para el proceso de asiento de apertura' AS NVARCHAR(500)) AS DescripcionParametro,
        CAST(NULL AS DATE) AS FecIni,
        CAST(NULL AS DATE) AS FecFin,
        CAST(56 AS INT) AS Orden
) AS fuente
    ON destino.TipoParametro = fuente.TipoParametro
   AND destino.CodigoParametro = fuente.CodigoParametro
WHEN MATCHED THEN
    UPDATE
    SET destino.ValorParametro = fuente.ValorParametro,
        destino.DescripcionParametro = fuente.DescripcionParametro,
        destino.FecIni = fuente.FecIni,
        destino.FecFin = fuente.FecFin,
        destino.Orden = fuente.Orden,
        destino.Activo = 1
WHEN NOT MATCHED BY TARGET THEN
    INSERT (TipoParametro, CodigoParametro, ValorParametro, DescripcionParametro, FecIni, FecFin, Orden, Activo)
    VALUES (fuente.TipoParametro, fuente.CodigoParametro, fuente.ValorParametro, fuente.DescripcionParametro, fuente.FecIni, fuente.FecFin, fuente.Orden, 1);
