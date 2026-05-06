USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Firma: Codex - 05/05/2026 | Permite multiples series por sede para NC(07) y ND(08) en SedesSeriesDocumentoComprobante.
IF EXISTS (
    SELECT 1
    FROM sys.key_constraints
    WHERE [type] = N'UQ'
      AND [name] = N'UX_SedesSeriesDocumentoComprobante_Sede_Documento'
      AND [parent_object_id] = OBJECT_ID(N'dbo.SedesSeriesDocumentoComprobante')
)
BEGIN
    ALTER TABLE [dbo].[SedesSeriesDocumentoComprobante]
    DROP CONSTRAINT [UX_SedesSeriesDocumentoComprobante_Sede_Documento];
END
GO
IF NOT EXISTS (
    SELECT 1
    FROM sys.indexes
    WHERE name = N'UX_SedesSeriesDocumentoComprobante_Sede_Documento_Activo_NoNotas'
      AND object_id = OBJECT_ID(N'dbo.SedesSeriesDocumentoComprobante')
)
BEGIN
    CREATE UNIQUE NONCLUSTERED INDEX [UX_SedesSeriesDocumentoComprobante_Sede_Documento_Activo_NoNotas]
    ON [dbo].[SedesSeriesDocumentoComprobante]
    (
        [SedeId] ASC,
        [CodigoSunat] ASC
    )
    WHERE [Activo] = 1
      AND [CodigoSunat] <> N'07'
      AND [CodigoSunat] <> N'08';
END
GO
