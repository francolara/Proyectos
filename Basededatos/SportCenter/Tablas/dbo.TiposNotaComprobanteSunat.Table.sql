USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   11/04/2026
-- Description:   Maestro SUNAT de tipos de nota de credito y debito (07/08).
-- =============================================
CREATE TABLE [dbo].[TiposNotaComprobanteSunat](
    [Id] [int] IDENTITY(1,1) NOT NULL,
    [TipoNota] [char](2) NOT NULL,
    [CodigoSunat] [nvarchar](2) NOT NULL,
    [Nombre] [nvarchar](250) NOT NULL,
    [Orden] [smallint] NOT NULL,
    [Activo] [bit] NOT NULL,
 CONSTRAINT [PK_TiposNotaComprobanteSunat] PRIMARY KEY CLUSTERED
(
    [Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY],
 CONSTRAINT [UQ_TiposNotaComprobanteSunat_TipoCodigo] UNIQUE NONCLUSTERED
(
    [TipoNota] ASC,
    [CodigoSunat] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[TiposNotaComprobanteSunat] ADD CONSTRAINT [DF_TiposNotaComprobanteSunat_Activo] DEFAULT ((1)) FOR [Activo]
GO
ALTER TABLE [dbo].[TiposNotaComprobanteSunat] WITH CHECK ADD CONSTRAINT [CK_TiposNotaComprobanteSunat_TipoNota]
CHECK ([TipoNota] IN ('07', '08'))
GO
