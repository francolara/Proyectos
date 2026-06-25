-- =============================================
-- Author:        FRANCO LARA
-- Create date:   21/06/2026
-- Description:   Catalogo SUNAT de tipos de afectacion al IGV para registros de compras y ventas.
-- =============================================

IF OBJECT_ID(N'dbo.CON_TipoAfectacionIGV', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_TipoAfectacionIGV
    (
        IdTipoAfectacionIGV INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_TipoAfectacionIGV PRIMARY KEY,
        CodigoSunat VARCHAR(10) NOT NULL,
        NombreAfectacion NVARCHAR(120) NOT NULL,
        Estado BIT NOT NULL CONSTRAINT DF_CON_TipoAfectacionIGV_Estado DEFAULT (1)
    );

    ALTER TABLE dbo.CON_TipoAfectacionIGV
        ADD CONSTRAINT UQ_CON_TipoAfectacionIGV_CodigoSunat
            UNIQUE (CodigoSunat);
END;

MERGE dbo.CON_TipoAfectacionIGV AS destino
USING
(
    VALUES
        ('10', N'Gravado - Operacion Onerosa', 1),
        ('20', N'Exonerado - Operacion Onerosa', 1),
        ('30', N'Inafecto - Operacion Onerosa', 1),
        ('40', N'Exportacion', 1),
        ('21', N'Exonerado - Transferencia Gratuita', 1),
        ('31', N'Inafecto - Transferencia Gratuita', 1)
) AS fuente (CodigoSunat, NombreAfectacion, Estado)
    ON destino.CodigoSunat = fuente.CodigoSunat
WHEN MATCHED THEN
    UPDATE SET
        NombreAfectacion = fuente.NombreAfectacion,
        Estado = fuente.Estado
WHEN NOT MATCHED BY TARGET THEN
    INSERT
    (
        CodigoSunat,
        NombreAfectacion,
        Estado
    )
    VALUES
    (
        fuente.CodigoSunat,
        fuente.NombreAfectacion,
        fuente.Estado
    );
