USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/04/2026
-- Description:   Series configuradas por sede para cada tipo de documento habilitado.
-- Firma:         Codex - 05/05/2026 | Ajusta indices para permitir multiples series activas por sede en NC/ND (07/08).
-- =============================================
CREATE TABLE [dbo].[SedesSeriesDocumentoComprobante](
    [Id] [int] IDENTITY(1,1) NOT NULL,
    [SedeId] [int] NOT NULL,
    [CodigoSunat] [nvarchar](4) NOT NULL,
    [NegocioSerieId] [int] NOT NULL,
    [Activo] [bit] NOT NULL,
    [FechaCreacion] [datetime2](7) NOT NULL,
    [FechaActualizacion] [datetime2](7) NULL,
    [UsuarioCreacion] [nvarchar](200) NULL,
    [UsuarioActualizacion] [nvarchar](200) NULL,
 CONSTRAINT [PK_SedesSeriesDocumentoComprobante] PRIMARY KEY CLUSTERED
(
    [Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY],
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[SedesSeriesDocumentoComprobante] ADD CONSTRAINT [DF_SedesSeriesDocumentoComprobante_Activo] DEFAULT ((1)) FOR [Activo]
GO
ALTER TABLE [dbo].[SedesSeriesDocumentoComprobante] ADD CONSTRAINT [DF_SedesSeriesDocumentoComprobante_FechaCreacion] DEFAULT (SYSUTCDATETIME()) FOR [FechaCreacion]
GO
ALTER TABLE [dbo].[SedesSeriesDocumentoComprobante]  WITH CHECK ADD  CONSTRAINT [FK_SedesSeriesDocumentoComprobante_Sedes] FOREIGN KEY([SedeId])
REFERENCES [dbo].[Sedes] ([Id])
GO
ALTER TABLE [dbo].[SedesSeriesDocumentoComprobante] CHECK CONSTRAINT [FK_SedesSeriesDocumentoComprobante_Sedes]
GO
ALTER TABLE [dbo].[SedesSeriesDocumentoComprobante]  WITH CHECK ADD  CONSTRAINT [FK_SedesSeriesDocumentoComprobante_TiposDocumentoComprobanteSuperMaestro] FOREIGN KEY([CodigoSunat])
REFERENCES [dbo].[TiposDocumentoComprobanteSuperMaestro] ([CodigoSunat])
GO
ALTER TABLE [dbo].[SedesSeriesDocumentoComprobante] CHECK CONSTRAINT [FK_SedesSeriesDocumentoComprobante_TiposDocumentoComprobanteSuperMaestro]
GO
ALTER TABLE [dbo].[SedesSeriesDocumentoComprobante]  WITH CHECK ADD  CONSTRAINT [FK_SedesSeriesDocumentoComprobante_NegociosSeriesDocumentoComprobante] FOREIGN KEY([NegocioSerieId])
REFERENCES [dbo].[NegociosSeriesDocumentoComprobante] ([Id])
GO
ALTER TABLE [dbo].[SedesSeriesDocumentoComprobante] CHECK CONSTRAINT [FK_SedesSeriesDocumentoComprobante_NegociosSeriesDocumentoComprobante]
GO
CREATE UNIQUE NONCLUSTERED INDEX [UX_SedesSeriesDocumentoComprobante_Sede_Documento_Activo_NoNotas]
ON [dbo].[SedesSeriesDocumentoComprobante]
(
    [SedeId] ASC,
    [CodigoSunat] ASC
)
WHERE [Activo] = 1 AND [CodigoSunat] NOT IN (N'07', N'08')
WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
ON [PRIMARY]
GO
