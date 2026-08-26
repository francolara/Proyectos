-- =============================================
-- Author:        FRANCO LARA
-- Create date:   29/06/2026
-- Description:   Agrega maestro de tipos de percepcion, documento pendiente por compra, modulo contable PER, origenes y parametro CTADEPERCEPCION para provisiones de compras.
-- =============================================
-- Firma: FRANCO LARA - 25/08/2026 | Mantiene el CHECK de modulos alineado con el catalogo contable vigente al reejecutar la migracion.

IF COL_LENGTH(N'dbo.COM_Compra', N'TienePercepcion') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD TienePercepcion BIT NOT NULL CONSTRAINT DF_COM_Compra_TienePercepcion DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.COM_Compra', N'IdTipoPercepcion') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD IdTipoPercepcion INT NULL;
END;

IF COL_LENGTH(N'dbo.COM_Compra', N'PorcentajePercepcion') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD PorcentajePercepcion DECIMAL(7,4) NOT NULL CONSTRAINT DF_COM_Compra_PorcentajePercepcion DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.COM_Compra', N'BasePercepcion') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD BasePercepcion DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_Compra_BasePercepcion DEFAULT (0);
END;

IF COL_LENGTH(N'dbo.COM_Compra', N'ImportePercepcion') IS NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD ImportePercepcion DECIMAL(18,2) NOT NULL CONSTRAINT DF_COM_Compra_ImportePercepcion DEFAULT (0);
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.foreign_keys
    WHERE name = N'FK_COM_Compra_ADM_TipoPercepcion'
)
AND COL_LENGTH(N'dbo.COM_Compra', N'IdTipoPercepcion') IS NOT NULL
BEGIN
    ALTER TABLE dbo.COM_Compra
        ADD CONSTRAINT FK_COM_Compra_ADM_TipoPercepcion
            FOREIGN KEY (IdTipoPercepcion) REFERENCES dbo.ADM_TipoPercepcion (IdTipoPercepcion);
END;

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
        CAST('ADMINISTRATIVO' AS VARCHAR(50)) AS TipoParametro,
        CAST('CTADEPERCEPCION' AS VARCHAR(50)) AS CodigoParametro,
        CAST(N'' AS NVARCHAR(200)) AS ValorParametro,
        CAST(N'Cuenta contable para percepciones de compras' AS NVARCHAR(500)) AS DescripcionParametro,
        CAST(NULL AS DATE) AS FecIni,
        CAST(NULL AS DATE) AS FecFin,
        CAST(285 AS INT) AS Orden
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

MERGE dbo.CON_OrigenMaestro AS destino
USING
(
    SELECT
        CAST('73' AS VARCHAR(2)) AS CodigoOrigen,
        CAST(N'PERCEPCIONES COMPRAS' AS NVARCHAR(150)) AS NombreOrigen,
        CAST(N'COMPRAS' AS NVARCHAR(100)) AS ModuloOrigen,
        CAST(1 AS BIT) AS PermiteRegistroManual,
        CAST(255 AS INT) AS Orden
) AS fuente
    ON destino.CodigoOrigen = fuente.CodigoOrigen
WHEN MATCHED THEN
    UPDATE
    SET destino.NombreOrigen = fuente.NombreOrigen,
        destino.ModuloOrigen = fuente.ModuloOrigen,
        destino.PermiteRegistroManual = fuente.PermiteRegistroManual,
        destino.Orden = fuente.Orden,
        destino.Estado = 1
WHEN NOT MATCHED BY TARGET THEN
    INSERT (CodigoOrigen, NombreOrigen, ModuloOrigen, PermiteRegistroManual, Estado, Orden)
    VALUES (fuente.CodigoOrigen, fuente.NombreOrigen, fuente.ModuloOrigen, fuente.PermiteRegistroManual, 1, fuente.Orden);

MERGE dbo.ADM_TipoPercepcion AS target
USING
(
    SELECT '01' AS Codigo, N'Venta interna general' AS Descripcion, CAST(2.0000 AS DECIMAL(7,4)) AS Porcentaje UNION ALL
    SELECT '02', N'Cliente agente percepcion', 0.5000 UNION ALL
    SELECT '03', N'Combustible', 1.0000 UNION ALL
    SELECT '04', N'Importacion general', 3.5000 UNION ALL
    SELECT '05', N'Importacion bienes usados', 5.0000 UNION ALL
    SELECT '06', N'Importacion supuesto especial', 10.0000
) AS source
    ON target.Codigo = source.Codigo
WHEN MATCHED THEN
    UPDATE SET
        target.Descripcion = source.Descripcion,
        target.Porcentaje = source.Porcentaje,
        target.Estado = 1
WHEN NOT MATCHED BY TARGET THEN
    INSERT (Codigo, Descripcion, Porcentaje, Estado)
    VALUES (source.Codigo, source.Descripcion, source.Porcentaje, 1);

DECLARE @IdEmpresa INT;

DECLARE empresa_cursor CURSOR LOCAL FAST_FORWARD FOR
SELECT e.IdEmpresa
FROM dbo.SEG_Empresa AS e;

OPEN empresa_cursor;

FETCH NEXT FROM empresa_cursor INTO @IdEmpresa;

WHILE @@FETCH_STATUS = 0
BEGIN
    EXEC dbo.usp_CON_GenerarOrigenesBaseEmpresa
        @IdEmpresa = @IdEmpresa,
        @UsuarioRegistro = N'SISTEMA';

    FETCH NEXT FROM empresa_cursor INTO @IdEmpresa;
END;

CLOSE empresa_cursor;
DEALLOCATE empresa_cursor;
