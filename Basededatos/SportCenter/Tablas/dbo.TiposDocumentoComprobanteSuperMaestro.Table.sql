USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/04/2026
-- Description:   Super maestro de tipos de documento para emision (SUNAT + Recibo Interno).
-- =============================================
CREATE TABLE [dbo].[TiposDocumentoComprobanteSuperMaestro](
    [CodigoSunat] [nvarchar](4) NOT NULL,
    [Nombre] [nvarchar](150) NOT NULL,
    [Tributario] [bit] NOT NULL,
    [Habilitado] [bit] NOT NULL,
    [Orden] [tinyint] NOT NULL,
    [Activo] [bit] NOT NULL,
 CONSTRAINT [PK_TiposDocumentoComprobanteSuperMaestro] PRIMARY KEY CLUSTERED
(
    [CodigoSunat] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[TiposDocumentoComprobanteSuperMaestro] ADD CONSTRAINT [DF_TiposDocumentoComprobanteSuperMaestro_Tributario] DEFAULT ((1)) FOR [Tributario]
GO
ALTER TABLE [dbo].[TiposDocumentoComprobanteSuperMaestro] ADD CONSTRAINT [DF_TiposDocumentoComprobanteSuperMaestro_Habilitado] DEFAULT ((0)) FOR [Habilitado]
GO
ALTER TABLE [dbo].[TiposDocumentoComprobanteSuperMaestro] ADD CONSTRAINT [DF_TiposDocumentoComprobanteSuperMaestro_Activo] DEFAULT ((1)) FOR [Activo]
GO
