-- Firma: FRANCO LARA | Create date: 11/11/2026 | Migra motivos de Proveeperu_071126 y complementa los códigos 10 a 13 del catálogo 09 de SUNAT.
USE [SistemaWebHYT];
GO
IF COL_LENGTH(N'dbo.ComprobanteHYT',N'CodigoMotivoNota') IS NULL
    ALTER TABLE dbo.ComprobanteHYT ADD CodigoMotivoNota NVARCHAR(2) NULL;
GO
INSERT INTO dbo.MotivoNotaCreditoHYT (Codigo,Descripcion)
SELECT origen.Codigo,origen.Descripcion
FROM (VALUES
    ('01',N'Anulación de la operación'),
    ('02',N'Anulación por error en el RUC'),
    ('03',N'Corrección por error en la descripción'),
    ('04',N'Descuento global'),
    ('05',N'Descuento por ítem'),
    ('06',N'Devolución total'),
    ('07',N'Devolución por ítem'),
    ('08',N'Bonificación'),
    ('09',N'Disminución en el valor'),
    ('10',N'Otros conceptos'),
    ('11',N'Ajustes de operaciones de exportación'),
    ('12',N'Ajustes afectos al IVAP'),
    ('13',N'Corrección del monto neto pendiente de pago o de las fechas e importes de las cuotas')
) AS origen(Codigo,Descripcion)
WHERE NOT EXISTS (SELECT 1 FROM dbo.MotivoNotaCreditoHYT AS actual WITH (UPDLOCK,HOLDLOCK) WHERE actual.Codigo=origen.Codigo);
GO
