USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/04/2026
-- Description:   Tipos de documento de comprobante habilitados por negocio.
-- =============================================
CREATE TABLE [dbo].[NegociosTiposDocumentoComprobante](
    [Id] [int] IDENTITY(1,1) NOT NULL,
    [NegocioId] [int] NOT NULL,
    [CodigoSunat] [nvarchar](4) NOT NULL,
    [Activo] [bit] NOT NULL,
    [FechaCreacion] [datetime2](7) NOT NULL,
    [FechaActualizacion] [datetime2](7) NULL,
    [UsuarioCreacion] [nvarchar](200) NULL,
    [UsuarioActualizacion] [nvarchar](200) NULL,
 CONSTRAINT [PK_NegociosTiposDocumentoComprobante] PRIMARY KEY CLUSTERED
(
    [Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY],
 CONSTRAINT [UX_NegociosTiposDocumentoComprobante_Negocio_Documento] UNIQUE NONCLUSTERED
(
    [NegocioId] ASC,
    [CodigoSunat] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[NegociosTiposDocumentoComprobante] ADD CONSTRAINT [DF_NegociosTiposDocumentoComprobante_Activo] DEFAULT ((1)) FOR [Activo]
GO
ALTER TABLE [dbo].[NegociosTiposDocumentoComprobante] ADD CONSTRAINT [DF_NegociosTiposDocumentoComprobante_FechaCreacion] DEFAULT (SYSUTCDATETIME()) FOR [FechaCreacion]
GO
ALTER TABLE [dbo].[NegociosTiposDocumentoComprobante]  WITH CHECK ADD  CONSTRAINT [FK_NegociosTiposDocumentoComprobante_Negocios] FOREIGN KEY([NegocioId])
REFERENCES [dbo].[Negocios] ([Id])
GO
ALTER TABLE [dbo].[NegociosTiposDocumentoComprobante] CHECK CONSTRAINT [FK_NegociosTiposDocumentoComprobante_Negocios]
GO
ALTER TABLE [dbo].[NegociosTiposDocumentoComprobante]  WITH CHECK ADD  CONSTRAINT [FK_NegociosTiposDocumentoComprobante_TiposDocumentoComprobanteSuperMaestro] FOREIGN KEY([CodigoSunat])
REFERENCES [dbo].[TiposDocumentoComprobanteSuperMaestro] ([CodigoSunat])
GO
ALTER TABLE [dbo].[NegociosTiposDocumentoComprobante] CHECK CONSTRAINT [FK_NegociosTiposDocumentoComprobante_TiposDocumentoComprobanteSuperMaestro]
GO
