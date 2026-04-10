USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/04/2026
-- Description:   Series de comprobantes configuradas por negocio y tipo de documento.
-- =============================================
CREATE TABLE [dbo].[NegociosSeriesDocumentoComprobante](
    [Id] [int] IDENTITY(1,1) NOT NULL,
    [NegocioId] [int] NOT NULL,
    [CodigoSunat] [nvarchar](4) NOT NULL,
    [Serie] [nvarchar](4) NOT NULL,
    [Activo] [bit] NOT NULL,
    [FechaCreacion] [datetime2](7) NOT NULL,
    [FechaActualizacion] [datetime2](7) NULL,
    [UsuarioCreacion] [nvarchar](200) NULL,
    [UsuarioActualizacion] [nvarchar](200) NULL,
 CONSTRAINT [PK_NegociosSeriesDocumentoComprobante] PRIMARY KEY CLUSTERED
(
    [Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY],
 CONSTRAINT [UX_NegociosSeriesDocumentoComprobante_Negocio_Documento_Serie] UNIQUE NONCLUSTERED
(
    [NegocioId] ASC,
    [CodigoSunat] ASC,
    [Serie] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[NegociosSeriesDocumentoComprobante] ADD CONSTRAINT [DF_NegociosSeriesDocumentoComprobante_Activo] DEFAULT ((1)) FOR [Activo]
GO
ALTER TABLE [dbo].[NegociosSeriesDocumentoComprobante] ADD CONSTRAINT [DF_NegociosSeriesDocumentoComprobante_FechaCreacion] DEFAULT (SYSUTCDATETIME()) FOR [FechaCreacion]
GO
ALTER TABLE [dbo].[NegociosSeriesDocumentoComprobante]  WITH CHECK ADD  CONSTRAINT [FK_NegociosSeriesDocumentoComprobante_Negocios] FOREIGN KEY([NegocioId])
REFERENCES [dbo].[Negocios] ([Id])
GO
ALTER TABLE [dbo].[NegociosSeriesDocumentoComprobante] CHECK CONSTRAINT [FK_NegociosSeriesDocumentoComprobante_Negocios]
GO
ALTER TABLE [dbo].[NegociosSeriesDocumentoComprobante]  WITH CHECK ADD  CONSTRAINT [FK_NegociosSeriesDocumentoComprobante_TiposDocumentoComprobanteSuperMaestro] FOREIGN KEY([CodigoSunat])
REFERENCES [dbo].[TiposDocumentoComprobanteSuperMaestro] ([CodigoSunat])
GO
ALTER TABLE [dbo].[NegociosSeriesDocumentoComprobante] CHECK CONSTRAINT [FK_NegociosSeriesDocumentoComprobante_TiposDocumentoComprobanteSuperMaestro]
GO
