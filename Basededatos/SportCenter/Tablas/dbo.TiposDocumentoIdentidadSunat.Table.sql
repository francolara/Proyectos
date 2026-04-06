USE [DbSportCenter]
GO
/****** Object:  Table [dbo].[TiposDocumentoIdentidadSunat]    Script Date: 04/04/2026 ******/
-- Firma: Codex - 04/04/2026 | Tabla maestra centralizada de tipos de documento SUNAT para clientes y configuracion.
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TiposDocumentoIdentidadSunat](
    [CodigoSunat] [nvarchar](2) NOT NULL,
    [CodigoInterno] [nvarchar](20) NOT NULL,
    [Nombre] [nvarchar](150) NOT NULL,
    [Activo] [bit] NOT NULL,
    [Orden] [tinyint] NOT NULL,
 CONSTRAINT [PK_TiposDocumentoIdentidadSunat] PRIMARY KEY CLUSTERED
(
    [CodigoSunat] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY],
 CONSTRAINT [UQ_TiposDocumentoIdentidadSunat_CodigoInterno] UNIQUE NONCLUSTERED
(
    [CodigoInterno] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[TiposDocumentoIdentidadSunat] ADD CONSTRAINT [DF_TiposDocumentoIdentidadSunat_Activo] DEFAULT ((1)) FOR [Activo]
GO
