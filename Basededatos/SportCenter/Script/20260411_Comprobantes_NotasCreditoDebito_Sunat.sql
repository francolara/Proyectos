/*
Firma: Codex - 11/04/2026
Descripcion: Soporte de Notas de Credito/Debito (NC/ND) para comprobantes, con referencia al documento origen y tipos SUNAT.
Fuente SUNAT:
  - Catalogo N. 09 (tipos nota credito): https://www.sunat.gob.pe/legislacion/superin/2021/anexo-165-2021.pdf
  - Catalogo N. 10 (tipos nota debito):  https://www.sunat.gob.pe/legislacion/superin/2017/anexosV-318-2017.pdf
*/
USE [DbSportCenter]
GO

SET NOCOUNT ON;

IF OBJECT_ID(N'dbo.TiposNotaComprobanteSunat', N'U') IS NULL
BEGIN
    RAISERROR('No existe la tabla dbo.TiposNotaComprobanteSunat. Ejecuta primero su script de estructura.', 16, 1);
    RETURN;
END;

IF COL_LENGTH('dbo.ComprobantesElectronicos', 'ComprobanteReferenciaId') IS NULL
BEGIN
    ALTER TABLE dbo.ComprobantesElectronicos
        ADD ComprobanteReferenciaId INT NULL;
END;

IF COL_LENGTH('dbo.ComprobantesElectronicos', 'TipoNota') IS NULL
BEGIN
    ALTER TABLE dbo.ComprobantesElectronicos
        ADD TipoNota CHAR(2) NULL;
END;

IF COL_LENGTH('dbo.ComprobantesElectronicos', 'TipoNotaCodigoSunat') IS NULL
BEGIN
    ALTER TABLE dbo.ComprobantesElectronicos
        ADD TipoNotaCodigoSunat NVARCHAR(2) NULL;
END;

UPDATE dbo.ComprobantesElectronicos
SET TipoNota =
    CASE
        WHEN TipoNota = 'NC' THEN '07'
        WHEN TipoNota = 'ND' THEN '08'
        ELSE TipoNota
    END
WHERE TipoNota IN ('NC', 'ND');

IF EXISTS (SELECT 1 FROM sys.check_constraints WHERE name = 'CK_ComprobantesElectronicos_TipoNota')
BEGIN
    ALTER TABLE dbo.ComprobantesElectronicos DROP CONSTRAINT CK_ComprobantesElectronicos_TipoNota;
END;

ALTER TABLE dbo.ComprobantesElectronicos
WITH CHECK ADD CONSTRAINT CK_ComprobantesElectronicos_TipoNota
CHECK (TipoNota IS NULL OR TipoNota IN ('07', '08'));

IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = 'FK_ComprobantesElectronicos_Referencia')
BEGIN
    ALTER TABLE dbo.ComprobantesElectronicos
    WITH CHECK ADD CONSTRAINT FK_ComprobantesElectronicos_Referencia
    FOREIGN KEY (ComprobanteReferenciaId) REFERENCES dbo.ComprobantesElectronicos(Id);
END;

IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = 'FK_ComprobantesElectronicos_TipoNotaSunat')
BEGIN
    ALTER TABLE dbo.ComprobantesElectronicos
    WITH CHECK ADD CONSTRAINT FK_ComprobantesElectronicos_TipoNotaSunat
    FOREIGN KEY (TipoNota, TipoNotaCodigoSunat) REFERENCES dbo.TiposNotaComprobanteSunat(TipoNota, CodigoSunat);
END;

UPDATE dbo.TiposNotaComprobanteSunat
SET TipoNota =
    CASE
        WHEN TipoNota = 'NC' THEN '07'
        WHEN TipoNota = 'ND' THEN '08'
        ELSE TipoNota
    END
WHERE TipoNota IN ('NC', 'ND');

MERGE dbo.TiposNotaComprobanteSunat AS tgt
USING
(
    VALUES
        ('07', N'01', N'Anulacion de la operacion', 1, 1),
        ('07', N'02', N'Anulacion por error en el RUC', 2, 1),
        ('07', N'03', N'Correccion por error en la descripcion o atencion de reclamo', 3, 1),
        ('07', N'04', N'Descuento global', 4, 1),
        ('07', N'05', N'Descuento por item', 5, 1),
        ('07', N'06', N'Devolucion total', 6, 1),
        ('07', N'07', N'Devolucion por item', 7, 1),
        ('07', N'08', N'Bonificacion', 8, 1),
        ('07', N'09', N'Disminucion en el valor', 9, 1),
        ('07', N'10', N'Otros conceptos', 10, 1),
        ('07', N'11', N'Ajustes de operaciones de exportacion', 11, 1),
        ('07', N'12', N'Ajustes afectos al IVAP', 12, 1),
        ('07', N'13', N'Correccion o modificacion del monto neto pendiente de pago y/o cuotas', 13, 1),
        ('08', N'01', N'Intereses por mora', 1, 1),
        ('08', N'02', N'Aumento en el valor', 2, 1),
        ('08', N'03', N'Penalidades u otros conceptos', 3, 1)
) AS src (TipoNota, CodigoSunat, Nombre, Orden, Activo)
ON tgt.TipoNota = src.TipoNota
AND tgt.CodigoSunat = src.CodigoSunat
WHEN MATCHED THEN
    UPDATE SET
        tgt.Nombre = src.Nombre,
        tgt.Orden = src.Orden,
        tgt.Activo = src.Activo
WHEN NOT MATCHED BY TARGET THEN
    INSERT (TipoNota, CodigoSunat, Nombre, Orden, Activo)
    VALUES (src.TipoNota, src.CodigoSunat, src.Nombre, src.Orden, src.Activo);

UPDATE dbo.TiposDocumentoComprobanteSuperMaestro
SET Habilitado = 1
WHERE CodigoSunat IN (N'07', N'08')
  AND Activo = 1;
GO
